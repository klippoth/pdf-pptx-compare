from __future__ import annotations

import base64
from collections import Counter
from dataclasses import dataclass
from difflib import SequenceMatcher
from io import BytesIO
import json
from pathlib import Path
import re
import hashlib
from typing import Literal, Optional
import unicodedata

from PIL import Image, ImageDraw, ImageFont
from pydantic import BaseModel, Field

from app.services.models import (
    ParagraphLayout,
    TextLayout,
    PageImage,
    QcFindingSeverity,
    QcFindingType,
    SlideQcFinding,
    SlideQcResult,
    SlideQcStatus,
    TextSource,
)

DEFAULT_GENERAL_SYSTEM_PROMPT = (
    "You are a slide quality-control reviewer using a visual-first workflow. Compare a recreated "
    "PowerPoint slide against the original PDF reference. The PDF reference is always the source of truth. "
    "You will receive a labeled side-by-side comparison panel plus the original PPTX candidate image and "
    "the original PDF reference image. Use the side-by-side panel to orient yourself, then use the original "
    "images to verify details and return bounding boxes in normalized PPTX candidate-image coordinates "
    "[x0, y0, x1, y1] between 0 and 1. Report only obvious discrepancies in these categories: "
    "missing_content, extra_content, wrong_text, and wrong_color. Treat wrong_color only as a significant "
    "and obvious font/text color mismatch on readable text. Do not report general fill, accent, chart, or "
    "image color differences. Do not report line breaks or wrapping in this version. Do not report chart "
    "size/position, alignment, or general positioning issues. Work methodically from top to bottom and left "
    "to right. Re-check every finding against the original images before returning it. Ignore subtle layout "
    "drift, tiny spacing changes, and negligible rendering noise."
)

DEFAULT_GENERAL_USER_PROMPT = (
    "Image 1 is a labeled side-by-side comparison panel: the left side is the PPTX candidate slide and the "
    "right side is the PDF reference page. Image 2 is the original PPTX candidate slide. Image 3 is the "
    "original PDF reference page. Use the original PPTX candidate image as the coordinate system for any "
    "returned bounding box. Focus on obvious missing elements, extra elements, visible text differences, and "
    "significant obvious font/text color mismatches only. Do not report line-wrap issues. Ignore chart "
    "size/alignment/position changes. Ignore tiny edge-only discrepancies unless there is a clearly real border "
    "element or footer element missing or extra there."
)

DEFAULT_TEXT_SYSTEM_PROMPT = (
    "Supplemental text verification rules for the same single visual pass: use extracted PPTX slide text only "
    "as supporting context for the candidate side. Do not rely on extracted PDF text. Treat wrong_text as a "
    "clearly visible wording, spelling, typo, name, date, number, symbol, punctuation, currency marker, or "
    "percentage difference between the PPTX candidate image and the PDF reference image. Be especially careful "
    "with short labels, chart labels, superscripts, footnotes, and small numeric text. Do not proofread, "
    "rewrite grammar, normalize wording, or invent corrections that are not visibly supported by the images."
)

DEFAULT_TEXT_USER_PROMPT = (
    "Candidate PowerPoint text extracted from the PPTX file will be provided as supporting context. Use it only "
    "to sanity-check shaky or small regions on the candidate side. The final decision for wrong_text must still "
    "be based on visible evidence in the images."
)

PROMPT_CONFIG_KEYS = (
    "general_system_prompt",
    "general_user_prompt",
    "text_system_prompt",
    "text_user_prompt",
)


class _FindingSchema(BaseModel):
    type: Literal["missing_content", "extra_content", "wrong_text", "wrong_color", "line_break_issue", "size_position_issue"]
    severity: Literal["high", "medium", "low"] = "medium"
    message: str
    confidence: float = Field(ge=0.0, le=1.0)
    bbox: list[float] = Field(min_length=4, max_length=4)


class _SlideQcSchema(BaseModel):
    status: Literal["ok", "findings", "manual_review"]
    summary: Optional[str] = None
    bullets: list[str] = Field(default_factory=list)
    note: Optional[str] = None
    comparison_confidence: float = Field(default=0.8, ge=0.0, le=1.0)
    findings: list[_FindingSchema] = Field(default_factory=list)


@dataclass
class _TextDiscrepancySupport:
    candidate_bbox: tuple[float, float, float, float]
    reference_bbox: tuple[float, float, float, float]
    reference_text: str
    candidate_text: str
    text_score: float
    position_score: float
    label: str


@dataclass
class _CanvasFit:
    source_size: tuple[int, int]
    target_size: tuple[int, int]
    fitted_size: tuple[int, int]
    offset: tuple[int, int]


class OpenAIQCEvaluator:
    def __init__(
        self,
        *,
        api_key: Optional[str],
        model: str = "gpt-5.3-chat-latest",
        timeout_seconds: float = 90.0,
        max_image_dimension: int = 0,
        client=None,
    ):
        self.api_key = api_key
        self.model = model
        self.timeout_seconds = timeout_seconds
        self.max_image_dimension = max_image_dimension
        self._client = client

    def is_available(self) -> bool:
        return bool(self.api_key or self._client is not None)

    @classmethod
    def get_default_prompt_config(cls) -> dict[str, str]:
        return {
            "general_system_prompt": DEFAULT_GENERAL_SYSTEM_PROMPT,
            "general_user_prompt": DEFAULT_GENERAL_USER_PROMPT,
            "text_system_prompt": DEFAULT_TEXT_SYSTEM_PROMPT,
            "text_user_prompt": DEFAULT_TEXT_USER_PROMPT,
        }

    def compare_pages(
        self,
        *,
        slide_index: int,
        page_index: int,
        reference_page: PageImage,
        candidate_page: PageImage,
        reference_layout: Optional[TextLayout] = None,
        candidate_layout: Optional[TextLayout] = None,
        debug_output_dir: Optional[Path] = None,
        prompt_override: Optional[str] = None,
        prompt_config: Optional[dict[str, str]] = None,
    ) -> SlideQcResult:
        candidate_image, reference_image = self._prepare_comparison_images(
            candidate_page=candidate_page,
            reference_page=reference_page,
        )
        comparison_panel = self._build_side_by_side_comparison(
            candidate_image=candidate_image,
            reference_image=reference_image,
        )
        candidate_text_regions = self._collect_candidate_text_regions(candidate_layout)
        text_context = self._build_text_context(candidate_layout=candidate_layout)
        if debug_output_dir is not None:
            self._save_debug_inputs(
                debug_output_dir=debug_output_dir,
                comparison_panel=comparison_panel,
                candidate_image=candidate_image,
                reference_image=reference_image,
                prompt_override=prompt_override,
                prompt_config=prompt_config,
            )

        parsed = self._invoke_visual_model(
            comparison_panel=comparison_panel,
            candidate_image=candidate_image,
            reference_image=reference_image,
            text_context=text_context,
            prompt_override=prompt_override,
            prompt_config=prompt_config,
        )

        findings = self._materialize_findings(
            parsed.findings,
            allowed_types={
                QcFindingType.MISSING_CONTENT,
                QcFindingType.EXTRA_CONTENT,
                QcFindingType.WRONG_TEXT,
                QcFindingType.WRONG_COLOR,
            },
            next_finding_id=1,
        )
        findings = self._filter_visual_findings(
            findings,
            candidate_text_regions=candidate_text_regions,
        )
        findings = self._deduplicate_findings(findings)
        for finding_id, finding in enumerate(findings, start=1):
            finding.finding_id = finding_id

        status = self._merge_statuses(parsed.status, parsed.status, findings)
        summary, bullets, note = self._comments_for_output(parsed, findings)
        comparison_confidence = float(parsed.comparison_confidence)

        return SlideQcResult(
            slide_index=slide_index,
            page_index=page_index,
            status=status,
            findings=findings,
            alignment_confidence=max(0.0, min(1.0, comparison_confidence)),
            reference_source=TextSource.MODEL,
            candidate_source=TextSource.MODEL,
            summary=summary,
            comment_bullets=bullets,
            note=note,
        )

    def _invoke_visual_model(
        self,
        *,
        comparison_panel: Image.Image,
        candidate_image: Image.Image,
        reference_image: Image.Image,
        text_context: str,
        prompt_override: Optional[str] = None,
        prompt_config: Optional[dict[str, str]] = None,
    ) -> _SlideQcSchema:
        additional_instructions = self._format_prompt_override(prompt_override)
        resolved_prompt_config = self._resolve_prompt_config(prompt_config)
        return self._parse_model_response(
            input_content=[
                {
                    "role": "system",
                    "content": (
                        resolved_prompt_config["general_system_prompt"]
                        + "\n\n"
                        + resolved_prompt_config["text_system_prompt"]
                        + additional_instructions
                    ),
                },
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "input_text",
                            "text": (
                                resolved_prompt_config["general_user_prompt"]
                                + "\n\n"
                                + resolved_prompt_config["text_user_prompt"]
                            ),
                        },
                        *self._prompt_override_content(prompt_override),
                        {
                            "type": "input_text",
                            "text": text_context,
                        },
                        {"type": "input_image", "image_url": self._image_to_data_url(comparison_panel)},
                        {"type": "input_image", "image_url": self._image_to_data_url(candidate_image)},
                        {"type": "input_image", "image_url": self._image_to_data_url(reference_image)},
                    ],
                },
            ]
        )

    def _parse_model_response(
        self,
        *,
        input_content: list[dict],
    ) -> _SlideQcSchema:
        client = self._get_client(fresh=True)
        response = client.responses.parse(
            model=self.model,
            input=input_content,
            text_format=_SlideQcSchema,
            timeout=self.timeout_seconds,
        )
        return response.output_parsed

    def _get_client(self, *, fresh: bool = False):
        if self._client is not None:
            return self._client
        if not fresh and getattr(self, "_cached_client", None) is not None:
            return self._cached_client
        from openai import OpenAI

        client = OpenAI(api_key=self.api_key, timeout=self.timeout_seconds, max_retries=2)
        if not fresh:
            self._cached_client = client
        return client

    @classmethod
    def _resolve_prompt_config(cls, prompt_config: Optional[dict[str, str]]) -> dict[str, str]:
        defaults = cls.get_default_prompt_config()
        if not prompt_config:
            return defaults
        resolved = defaults.copy()
        for key in PROMPT_CONFIG_KEYS:
            value = prompt_config.get(key)
            if isinstance(value, str) and value.strip():
                resolved[key] = value.strip()
        return resolved

    def _prepare_comparison_images(
        self,
        *,
        candidate_page: PageImage,
        reference_page: PageImage,
    ) -> tuple[Image.Image, Image.Image]:
        candidate_source = Image.fromarray(candidate_page.image)
        reference_source = Image.fromarray(reference_page.image)

        candidate_scaled = self._scale_image(candidate_source)
        reference_scaled = self._scale_image(reference_source)
        return (candidate_scaled, reference_scaled)

    def _build_side_by_side_comparison(
        self,
        *,
        candidate_image: Image.Image,
        reference_image: Image.Image,
    ) -> Image.Image:
        gutter = max(28, int(round(max(candidate_image.height, reference_image.height) * 0.02)))
        header_height = max(72, int(round(max(candidate_image.height, reference_image.height) * 0.08)))
        outer_padding = max(24, gutter)
        background = (244, 244, 244)
        pane_background = (255, 255, 255)
        label_color = (40, 40, 40)
        border_color = (214, 214, 214)
        candidate_label = "PPTX CANDIDATE"
        reference_label = "PDF REFERENCE"
        target_height = max(candidate_image.height, reference_image.height)
        candidate_resized = self._contain_with_height(candidate_image, target_height)
        reference_resized = self._contain_with_height(reference_image, target_height)
        panel_width = (
            outer_padding * 2
            + candidate_resized.width
            + reference_resized.width
            + gutter
        )
        panel_height = outer_padding * 2 + header_height + target_height
        panel = Image.new("RGB", (panel_width, panel_height), background)
        draw = ImageDraw.Draw(panel)
        font = self._label_font(max(18, int(round(header_height * 0.32))))

        candidate_left = outer_padding
        reference_left = outer_padding + candidate_resized.width + gutter
        pane_top = outer_padding + header_height

        draw.rectangle(
            (candidate_left - 1, pane_top - 1, candidate_left + candidate_resized.width, pane_top + target_height),
            fill=pane_background,
            outline=border_color,
            width=1,
        )
        draw.rectangle(
            (reference_left - 1, pane_top - 1, reference_left + reference_resized.width, pane_top + target_height),
            fill=pane_background,
            outline=border_color,
            width=1,
        )
        panel.paste(candidate_resized, (candidate_left, pane_top))
        panel.paste(reference_resized, (reference_left, pane_top))

        candidate_label_bbox = draw.textbbox((0, 0), candidate_label, font=font)
        reference_label_bbox = draw.textbbox((0, 0), reference_label, font=font)
        draw.text(
            (candidate_left, outer_padding + (header_height - (candidate_label_bbox[3] - candidate_label_bbox[1])) / 2),
            candidate_label,
            fill=label_color,
            font=font,
        )
        draw.text(
            (reference_left, outer_padding + (header_height - (reference_label_bbox[3] - reference_label_bbox[1])) / 2),
            reference_label,
            fill=label_color,
            font=font,
        )
        return panel

    @staticmethod
    def _contain_with_height(image: Image.Image, target_height: int) -> Image.Image:
        if image.height == target_height:
            return image
        scale = target_height / float(max(image.height, 1))
        target_width = max(1, int(round(image.width * scale)))
        return image.resize((target_width, target_height), Image.Resampling.LANCZOS)

    @staticmethod
    def _label_font(size: int):
        try:
            return ImageFont.truetype("DejaVuSans-Bold.ttf", size=size)
        except OSError:  # pragma: no cover - environment fallback
            return ImageFont.load_default()

    def _scale_image(self, image: Image.Image) -> Image.Image:
        if self.max_image_dimension <= 0:
            return image
        max_side = max(image.size)
        if max_side <= self.max_image_dimension:
            return image
        scale = self.max_image_dimension / float(max_side)
        target_size = (
            max(1, int(round(image.width * scale))),
            max(1, int(round(image.height * scale))),
        )
        return image.resize(target_size, Image.Resampling.LANCZOS)

    @staticmethod
    def _fit_image_to_canvas(image: Image.Image, target_size: tuple[int, int]) -> tuple[Image.Image, _CanvasFit]:
        canvas = Image.new("RGB", target_size, (255, 255, 255))
        scale = min(target_size[0] / image.width, target_size[1] / image.height)
        fitted_size = (
            max(1, int(round(image.width * scale))),
            max(1, int(round(image.height * scale))),
        )
        fitted = image.resize(fitted_size, Image.Resampling.LANCZOS)
        paste_x = (target_size[0] - fitted_size[0]) // 2
        paste_y = (target_size[1] - fitted_size[1]) // 2
        canvas.paste(fitted, (paste_x, paste_y))
        return canvas, _CanvasFit(
            source_size=image.size,
            target_size=target_size,
            fitted_size=fitted_size,
            offset=(paste_x, paste_y),
        )

    def _image_to_data_url(self, image: Image.Image) -> str:
        encoded = base64.b64encode(self._image_to_png_bytes(image)).decode("ascii")
        return f"data:image/png;base64,{encoded}"

    @staticmethod
    def _image_to_png_bytes(image: Image.Image) -> bytes:
        buffer = BytesIO()
        image.save(buffer, format="PNG", optimize=True)
        return buffer.getvalue()

    def _build_text_discrepancy_content(
        self,
        *,
        candidate_image: Image.Image,
        reference_image: Image.Image,
        text_discrepancies: list[_TextDiscrepancySupport],
        reference_fit: _CanvasFit,
        max_items: int = 8,
    ) -> list[dict]:
        content: list[dict] = []
        for index, discrepancy in enumerate(text_discrepancies[:max_items], start=1):
            candidate_crop = self._crop_to_bbox(candidate_image, self._expand_bbox(discrepancy.candidate_bbox, margin=0.02, min_size=0.08))
            reference_canvas_bbox = self._map_bbox_to_canvas(discrepancy.reference_bbox, reference_fit)
            reference_crop = self._crop_to_bbox(reference_image, self._expand_bbox(reference_canvas_bbox, margin=0.02, min_size=0.08))
            if discrepancy.reference_text and discrepancy.candidate_text:
                discrepancy_summary = (
                    f'Reference reads "{self._trim_text(discrepancy.reference_text, max_length=100)}". '
                    f'Candidate reads "{self._trim_text(discrepancy.candidate_text, max_length=100)}". '
                )
            elif discrepancy.candidate_text:
                discrepancy_summary = (
                    "Reference has no matching extracted text item in this region. "
                    f'Candidate reads "{self._trim_text(discrepancy.candidate_text, max_length=100)}". '
                )
            else:
                discrepancy_summary = (
                    f'Reference reads "{self._trim_text(discrepancy.reference_text, max_length=100)}". '
                    "Candidate has no matching extracted text item in this region. "
                )
            content.extend(
                [
                    {
                        "type": "input_text",
                        "text": (
                            f"Suspicious text diff {index} ({discrepancy.label}). "
                            f"{discrepancy_summary}"
                            "The next image is the candidate text crop and the image after that is the reference text crop for the matching region."
                        ),
                    },
                    {"type": "input_image", "image_url": self._image_to_data_url(candidate_crop)},
                    {"type": "input_image", "image_url": self._image_to_data_url(reference_crop)},
                ]
            )
        return content

    def _materialize_findings(
        self,
        raw_findings: list[_FindingSchema],
        *,
        allowed_types: set[QcFindingType],
        next_finding_id: int,
    ) -> list[SlideQcFinding]:
        findings: list[SlideQcFinding] = []
        current_id = next_finding_id
        for finding in raw_findings:
            finding_type = QcFindingType(finding.type)
            if finding_type not in allowed_types:
                continue
            findings.append(
                SlideQcFinding(
                    finding_id=current_id,
                    finding_type=finding_type,
                    severity=QcFindingSeverity(finding.severity),
                    bbox=self._normalize_bbox(finding.bbox),
                    message=finding.message.strip(),
                    confidence=max(0.0, min(1.0, float(finding.confidence))),
                )
            )
            current_id += 1
        return findings

    def _filter_visual_findings(
        self,
        findings: list[SlideQcFinding],
        *,
        candidate_text_regions: list[tuple[float, float, float, float]],
    ) -> list[SlideQcFinding]:
        kept: list[SlideQcFinding] = []
        for finding in findings:
            if finding.finding_type == QcFindingType.SIZE_POSITION_ISSUE:
                continue
            if finding.finding_type == QcFindingType.LINE_BREAK_ISSUE:
                continue
            if finding.finding_type in {QcFindingType.MISSING_CONTENT, QcFindingType.EXTRA_CONTENT}:
                if self._bbox_is_within_edge_band(finding.bbox):
                    continue
            if finding.finding_type == QcFindingType.WRONG_COLOR:
                if finding.confidence < 0.9:
                    continue
                if not candidate_text_regions:
                    continue
                if not any(self._text_color_finding_matches_region(finding, bbox) for bbox in candidate_text_regions):
                    continue
            kept.append(finding)
        return kept

    @staticmethod
    def _bbox_is_within_edge_band(
        bbox: tuple[float, float, float, float],
        *,
        edge_band: float = 0.045,
    ) -> bool:
        return (
            bbox[2] <= edge_band
            or bbox[0] >= 1.0 - edge_band
            or bbox[3] <= edge_band
            or bbox[1] >= 1.0 - edge_band
        )

    def _text_color_finding_matches_region(
        self,
        finding: SlideQcFinding,
        bbox: tuple[float, float, float, float],
    ) -> bool:
        if self._bbox_iou(finding.bbox, bbox) >= 0.08:
            return True
        return self._bbox_contains_point(bbox, self._bbox_center(finding.bbox))

    def _filter_text_findings_against_support(
        self,
        findings: list[SlideQcFinding],
        *,
        text_discrepancies: list[_TextDiscrepancySupport],
        candidate_text_regions: list[tuple[float, float, float, float]],
    ) -> list[SlideQcFinding]:
        kept: list[SlideQcFinding] = []
        for finding in findings:
            if any(self._text_finding_matches_support(finding, support) for support in text_discrepancies):
                kept.append(finding)
                continue
            if self._allow_visual_text_fallback(
                finding,
                candidate_text_regions=candidate_text_regions,
            ):
                kept.append(finding)
        return kept

    def _allow_visual_text_fallback(
        self,
        finding: SlideQcFinding,
        *,
        candidate_text_regions: list[tuple[float, float, float, float]],
    ) -> bool:
        if finding.confidence < 0.9:
            return False
        if any(self._text_color_finding_matches_region(finding, bbox) for bbox in candidate_text_regions):
            return False
        return True

    def _text_finding_matches_support(
        self,
        finding: SlideQcFinding,
        support: _TextDiscrepancySupport,
    ) -> bool:
        if self._bbox_iou(finding.bbox, support.candidate_bbox) >= 0.08:
            return True
        if self._bbox_contains_point(support.candidate_bbox, self._bbox_center(finding.bbox)):
            return True
        return self._region_label(finding.bbox) == self._region_label(support.candidate_bbox)

    def _deduplicate_findings(self, findings: list[SlideQcFinding]) -> list[SlideQcFinding]:
        deduped: list[SlideQcFinding] = []
        for finding in findings:
            match_index = next(
                (
                    index
                    for index, existing in enumerate(deduped)
                    if existing.finding_type == finding.finding_type and self._bbox_iou(existing.bbox, finding.bbox) >= 0.82
                ),
                None,
            )
            if match_index is None:
                deduped.append(finding)
                continue
            existing = deduped[match_index]
            if finding.confidence > existing.confidence:
                deduped[match_index] = finding
        return deduped

    def _build_text_context(
        self,
        *,
        candidate_layout: Optional[TextLayout],
    ) -> str:
        return (
            "Supporting candidate-side text extracted from the PPTX file is provided below. "
            "Use it to verify small or shaky text regions on the PPTX candidate side only. "
            "Do not assume any PDF text extraction is available.\n\n"
            f"Candidate PowerPoint text:\n{self._format_layout(candidate_layout)}"
        )

    def _format_layout(self, layout: Optional[TextLayout], *, max_items: int = 36) -> str:
        text_items = self._text_items(layout)
        if not text_items:
            return "- No extracted text available."

        ordered = sorted(text_items, key=lambda item: (item.bbox[1], item.bbox[0]))[:max_items]
        lines: list[str] = []
        for item in ordered:
            lines.append(f"- {self._region_label(item.bbox)}: {self._trim_text(item.text)}")
        if len(text_items) > max_items:
            lines.append(f"- ... {len(text_items) - max_items} more text blocks omitted")
        return "\n".join(lines)

    @staticmethod
    def _region_label(bbox) -> str:
        x_center = (bbox[0] + bbox[2]) / 2.0
        y_center = (bbox[1] + bbox[3]) / 2.0
        vertical = "top" if y_center < 0.33 else "middle" if y_center < 0.66 else "bottom"
        horizontal = "left" if x_center < 0.33 else "center" if x_center < 0.66 else "right"
        return f"{vertical}-{horizontal}"

    @staticmethod
    def _trim_text(text: str, *, max_length: int = 180) -> str:
        compact = " ".join(text.split())
        if len(compact) <= max_length:
            return compact
        return compact[: max_length - 1].rstrip() + "…"

    def _format_text_discrepancies(
        self,
        text_discrepancies: list[_TextDiscrepancySupport],
    ) -> str:
        items: list[str] = []
        for discrepancy in text_discrepancies:
            items.append(
                f"- {self._region_label(discrepancy.reference_bbox)} ({discrepancy.label}): "
                f'reference="{self._trim_text(discrepancy.reference_text, max_length=120)}" | '
                f'candidate="{self._trim_text(discrepancy.candidate_text, max_length=120)}"'
            )
        return "\n".join(items) if items else "- No obvious deterministic text mismatches detected."

    def _collect_text_discrepancies(
        self,
        *,
        reference_layout: Optional[TextLayout],
        candidate_layout: Optional[TextLayout],
        max_items: int = 12,
    ) -> list[_TextDiscrepancySupport]:
        if reference_layout is None or candidate_layout is None:
            return []
        reference_items = self._text_items(reference_layout)
        candidate_items = self._text_items(candidate_layout)
        if not reference_items or not candidate_items:
            return []

        ordered_reference = sorted(reference_items, key=lambda item: (item.bbox[1], item.bbox[0]))
        ordered_candidate = sorted(candidate_items, key=lambda item: (item.bbox[1], item.bbox[0]))
        matches, matched_reference_indexes, matched_candidate_indexes = self._match_text_blocks(
            ordered_reference,
            ordered_candidate,
        )
        items: list[_TextDiscrepancySupport] = []
        for reference_index, candidate_index, text_score, position_score in matches:
            reference_item = ordered_reference[reference_index]
            candidate_item = ordered_candidate[candidate_index]
            reference_text = self._normalize_text(reference_item.text)
            candidate_text = self._normalize_text(candidate_item.text)
            if not reference_text or not candidate_text or reference_text == candidate_text:
                continue
            if text_score < 0.58 and position_score < 0.82:
                continue
            label = self._label_for_text_mismatch(reference_item.text, candidate_item.text, text_score=text_score)
            items.append(
                _TextDiscrepancySupport(
                    candidate_bbox=self._normalize_bbox(list(candidate_item.bbox)),
                    reference_bbox=self._normalize_bbox(list(reference_item.bbox)),
                    reference_text=reference_item.text,
                    candidate_text=candidate_item.text,
                    text_score=text_score,
                    position_score=position_score,
                    label=label,
                )
            )
            if len(items) >= max_items:
                break

        if len(items) < max_items:
            items.extend(
                self._collect_candidate_only_text_discrepancies(
                    ordered_reference=ordered_reference,
                    ordered_candidate=ordered_candidate,
                    matched_candidate_indexes=matched_candidate_indexes,
                    max_items=max_items - len(items),
                )
            )

        if len(items) < max_items:
            items.extend(
                self._collect_reference_only_text_discrepancies(
                    ordered_reference=ordered_reference,
                    ordered_candidate=ordered_candidate,
                    matched_reference_indexes=matched_reference_indexes,
                    max_items=max_items - len(items),
                )
            )
        return items

    def _collect_candidate_text_regions(
        self,
        candidate_layout: Optional[TextLayout],
    ) -> list[tuple[float, float, float, float]]:
        if candidate_layout is None:
            return []
        return [self._normalize_bbox(list(item.bbox)) for item in self._text_items(candidate_layout) if item.text.strip()]

    def _match_text_blocks(
        self,
        reference_paragraphs: list[ParagraphLayout | TextBox],
        candidate_paragraphs: list[ParagraphLayout | TextBox],
    ) -> tuple[list[tuple[int, int, float, float]], set[int], set[int]]:
        used_candidate_indexes: set[int] = set()
        used_reference_indexes: set[int] = set()
        matches: list[tuple[int, int, float, float]] = []

        for reference_index, reference_paragraph in enumerate(reference_paragraphs):
            normalized_reference = self._normalize_text(reference_paragraph.text)
            if not normalized_reference:
                continue

            best_index = -1
            best_score = -1.0
            best_text_score = 0.0
            best_position_score = 0.0

            for candidate_index, candidate_paragraph in enumerate(candidate_paragraphs):
                if candidate_index in used_candidate_indexes:
                    continue
                normalized_candidate = self._normalize_text(candidate_paragraph.text)
                if not normalized_candidate:
                    continue

                text_score = SequenceMatcher(None, normalized_reference, normalized_candidate).ratio()
                position_score = self._position_similarity(reference_paragraph.bbox, candidate_paragraph.bbox)
                combined_score = (text_score * 0.55) + (position_score * 0.45)
                if self._region_label(reference_paragraph.bbox) == self._region_label(candidate_paragraph.bbox):
                    combined_score += 0.04

                if combined_score > best_score:
                    best_index = candidate_index
                    best_score = combined_score
                    best_text_score = text_score
                    best_position_score = position_score

            if best_index >= 0:
                used_candidate_indexes.add(best_index)
                used_reference_indexes.add(reference_index)
                matches.append((reference_index, best_index, best_text_score, best_position_score))

        return matches, used_reference_indexes, used_candidate_indexes

    def _collect_candidate_only_text_discrepancies(
        self,
        *,
        ordered_reference: list[ParagraphLayout | TextBox],
        ordered_candidate: list[ParagraphLayout | TextBox],
        matched_candidate_indexes: set[int],
        max_items: int,
    ) -> list[_TextDiscrepancySupport]:
        if max_items <= 0:
            return []
        reference_counts = Counter(
            self._normalize_text(item.text)
            for item in ordered_reference
            if self._normalize_text(item.text)
        )
        supports: list[_TextDiscrepancySupport] = []
        for candidate_index, candidate_item in enumerate(ordered_candidate):
            normalized_candidate = self._normalize_text(candidate_item.text)
            if not normalized_candidate:
                continue
            if candidate_index in matched_candidate_indexes:
                if reference_counts[normalized_candidate] > 0:
                    reference_counts[normalized_candidate] -= 1
                continue
            if reference_counts[normalized_candidate] > 0:
                reference_counts[normalized_candidate] -= 1
                continue
            similar_match = self._best_similarity_against_items(candidate_item, ordered_reference)
            if similar_match and (similar_match[0] >= 0.55 or similar_match[1] >= 0.78):
                continue
            supports.append(
                _TextDiscrepancySupport(
                    candidate_bbox=self._normalize_bbox(list(candidate_item.bbox)),
                    reference_bbox=self._normalize_bbox(list(candidate_item.bbox)),
                    reference_text="",
                    candidate_text=candidate_item.text,
                    text_score=0.0,
                    position_score=0.0,
                    label="candidate-only extra text item",
                )
            )
            if len(supports) >= max_items:
                break
        return supports

    def _collect_reference_only_text_discrepancies(
        self,
        *,
        ordered_reference: list[ParagraphLayout | TextBox],
        ordered_candidate: list[ParagraphLayout | TextBox],
        matched_reference_indexes: set[int],
        max_items: int,
    ) -> list[_TextDiscrepancySupport]:
        if max_items <= 0:
            return []
        candidate_counts = Counter(
            self._normalize_text(item.text)
            for item in ordered_candidate
            if self._normalize_text(item.text)
        )
        supports: list[_TextDiscrepancySupport] = []
        for reference_index, reference_item in enumerate(ordered_reference):
            normalized_reference = self._normalize_text(reference_item.text)
            if not normalized_reference:
                continue
            if reference_index in matched_reference_indexes:
                if candidate_counts[normalized_reference] > 0:
                    candidate_counts[normalized_reference] -= 1
                continue
            if candidate_counts[normalized_reference] > 0:
                candidate_counts[normalized_reference] -= 1
                continue
            similar_match = self._best_similarity_against_items(reference_item, ordered_candidate)
            if similar_match and (similar_match[0] >= 0.55 or similar_match[1] >= 0.78):
                continue
            supports.append(
                _TextDiscrepancySupport(
                    candidate_bbox=self._normalize_bbox(list(reference_item.bbox)),
                    reference_bbox=self._normalize_bbox(list(reference_item.bbox)),
                    reference_text=reference_item.text,
                    candidate_text="",
                    text_score=0.0,
                    position_score=0.0,
                    label="reference-only missing text item",
                )
            )
            if len(supports) >= max_items:
                break
        return supports

    def _best_similarity_against_items(
        self,
        anchor_item: ParagraphLayout | TextBox,
        other_items: list[ParagraphLayout | TextBox],
    ) -> tuple[float, float] | None:
        normalized_anchor = self._normalize_text(anchor_item.text)
        if not normalized_anchor:
            return None
        best_text_score = 0.0
        best_position_score = 0.0
        found = False
        for other_item in other_items:
            normalized_other = self._normalize_text(other_item.text)
            if not normalized_other:
                continue
            found = True
            text_score = SequenceMatcher(None, normalized_anchor, normalized_other).ratio()
            position_score = self._position_similarity(anchor_item.bbox, other_item.bbox)
            if (text_score, position_score) > (best_text_score, best_position_score):
                best_text_score = text_score
                best_position_score = position_score
        if not found:
            return None
        return best_text_score, best_position_score

    @staticmethod
    def _text_items(layout: Optional[TextLayout]) -> list[ParagraphLayout | TextBox]:
        if layout is None:
            return []
        if layout.lines:
            return [line for line in layout.lines if line.text.strip()]
        return [paragraph for paragraph in layout.paragraphs if paragraph.text.strip()]

    @staticmethod
    def _normalize_text(value: str) -> str:
        normalized = unicodedata.normalize("NFKC", value or "")
        normalized = (
            normalized
            .replace("\u00ad", "")
            .replace("\u200b", "")
            .replace("\u200c", "")
            .replace("\u200d", "")
            .replace("\ufeff", "")
            .replace("’", "'")
            .replace("“", '"')
            .replace("”", '"')
            .replace("–", "-")
            .replace("—", "-")
        )
        normalized = re.sub(r"([A-Za-z])-\s+([A-Za-z])", r"\1-\2", normalized)
        normalized = re.sub(r"\s+", " ", normalized)
        return normalized.strip().casefold()

    def _label_for_text_mismatch(self, reference_text: str, candidate_text: str, *, text_score: float) -> str:
        if self._number_tokens(reference_text) != self._number_tokens(candidate_text):
            return "numeric-token mismatch"
        if self._has_non_ascii_or_confusable_difference(reference_text, candidate_text):
            return "possible unicode/confusable-character mismatch"
        if self._is_short_symbol_heavy_label(reference_text, candidate_text):
            return "short label or symbol-level text mismatch"
        return "likely line-level typo/spelling difference" if text_score >= 0.82 else "line-level text mismatch"

    @staticmethod
    def _number_tokens(text: str) -> list[str]:
        return re.findall(r"[\$\(]?-?\d[\d,]*\.?\d*[%\)BMK]?", text or "", flags=re.IGNORECASE)

    def _has_non_ascii_or_confusable_difference(self, reference_text: str, candidate_text: str) -> bool:
        return (
            self._contains_non_ascii(reference_text)
            or self._contains_non_ascii(candidate_text)
            or self._looks_like_confusable_mismatch(reference_text, candidate_text)
        )

    @staticmethod
    def _contains_non_ascii(text: str) -> bool:
        return any(ord(character) > 127 for character in (text or ""))

    @staticmethod
    def _looks_like_confusable_mismatch(reference_text: str, candidate_text: str) -> bool:
        confusable_map = str.maketrans(
            {
                "А": "A", "В": "B", "С": "C", "Е": "E", "Н": "H", "К": "K", "М": "M", "О": "O", "Р": "P", "Т": "T", "Х": "X",
                "а": "a", "е": "e", "о": "o", "р": "p", "с": "c", "у": "y", "х": "x", "і": "i", "Ι": "I", "Β": "B",
                "Α": "A", "Ε": "E", "Ζ": "Z", "Η": "H", "Ι": "I", "Κ": "K", "Μ": "M", "Ν": "N", "Ο": "O", "Ρ": "P",
                "Τ": "T", "Υ": "Y", "Χ": "X",
            }
        )
        normalized_reference = unicodedata.normalize("NFKC", reference_text or "").translate(confusable_map)
        normalized_candidate = unicodedata.normalize("NFKC", candidate_text or "").translate(confusable_map)
        return normalized_reference.casefold() == normalized_candidate.casefold() and reference_text != candidate_text

    @staticmethod
    def _is_short_symbol_heavy_label(reference_text: str, candidate_text: str) -> bool:
        combined = f"{reference_text}{candidate_text}"
        compact = re.sub(r"\s+", "", combined)
        return len(compact) <= 12 and any(character in compact for character in "$%()")

    @staticmethod
    def _position_similarity(reference_bbox, candidate_bbox) -> float:
        reference_center_x = (reference_bbox[0] + reference_bbox[2]) / 2.0
        reference_center_y = (reference_bbox[1] + reference_bbox[3]) / 2.0
        candidate_center_x = (candidate_bbox[0] + candidate_bbox[2]) / 2.0
        candidate_center_y = (candidate_bbox[1] + candidate_bbox[3]) / 2.0
        distance = ((reference_center_x - candidate_center_x) ** 2 + (reference_center_y - candidate_center_y) ** 2) ** 0.5
        return max(0.0, 1.0 - (distance / 1.2))

    @staticmethod
    def _map_bbox_to_canvas(bbox: tuple[float, float, float, float], fit: _CanvasFit) -> tuple[float, float, float, float]:
        x0 = ((bbox[0] * fit.fitted_size[0]) + fit.offset[0]) / max(fit.target_size[0], 1)
        y0 = ((bbox[1] * fit.fitted_size[1]) + fit.offset[1]) / max(fit.target_size[1], 1)
        x1 = ((bbox[2] * fit.fitted_size[0]) + fit.offset[0]) / max(fit.target_size[0], 1)
        y1 = ((bbox[3] * fit.fitted_size[1]) + fit.offset[1]) / max(fit.target_size[1], 1)
        return (
            max(0.0, min(1.0, x0)),
            max(0.0, min(1.0, y0)),
            max(0.0, min(1.0, x1)),
            max(0.0, min(1.0, y1)),
        )

    def _save_debug_inputs(
        self,
        *,
        debug_output_dir: Path,
        comparison_panel: Image.Image,
        candidate_image: Image.Image,
        reference_image: Image.Image,
        prompt_override: Optional[str] = None,
        prompt_config: Optional[dict[str, str]] = None,
    ) -> None:
        debug_output_dir.mkdir(parents=True, exist_ok=True)
        comparison_panel_png_bytes = OpenAIQCEvaluator._image_to_png_bytes(comparison_panel)
        candidate_png_bytes = OpenAIQCEvaluator._image_to_png_bytes(candidate_image)
        reference_png_bytes = OpenAIQCEvaluator._image_to_png_bytes(reference_image)
        (debug_output_dir / "01-comparison-panel.png").write_bytes(comparison_panel_png_bytes)
        (debug_output_dir / "02-candidate-slide.png").write_bytes(candidate_png_bytes)
        (debug_output_dir / "03-reference-page.png").write_bytes(reference_png_bytes)
        metadata = {
            "comparison_panel": {
                "width": comparison_panel.width,
                "height": comparison_panel.height,
                "png_bytes": len(comparison_panel_png_bytes),
                "sha256": hashlib.sha256(comparison_panel_png_bytes).hexdigest(),
                "labels": {
                    "left": "PPTX candidate",
                    "right": "PDF reference",
                },
            },
            "candidate": {
                "width": candidate_image.width,
                "height": candidate_image.height,
                "png_bytes": len(candidate_png_bytes),
                "sha256": hashlib.sha256(candidate_png_bytes).hexdigest(),
            },
            "reference": {
                "width": reference_image.width,
                "height": reference_image.height,
                "png_bytes": len(reference_png_bytes),
                "sha256": hashlib.sha256(reference_png_bytes).hexdigest(),
            },
            "normalization": {
                "max_image_dimension": self.max_image_dimension,
                "candidate_image_size": [candidate_image.width, candidate_image.height],
                "reference_image_size": [reference_image.width, reference_image.height],
                "note": (
                    "These PNG files are the exact bytes uploaded to OpenAI by this app. "
                    "Any further internal model-side processing is not visible from the client."
                ),
            },
            "prompt_override": (prompt_override or "").strip(),
            "prompt_config": self._resolve_prompt_config(prompt_config),
        }
        (debug_output_dir / "00-upload-metadata.json").write_text(json.dumps(metadata, indent=2), encoding="utf-8")

    def _comments_for_output(
        self,
        result: _SlideQcSchema,
        findings: list[SlideQcFinding],
    ) -> tuple[Optional[str], list[str], Optional[str]]:
        if not findings:
            return None, [], None
        active_types = {finding.finding_type for finding in findings}
        summary = self._summary_for_output(result, has_kept_findings=True) or "AI review findings"
        bullets = self._deduplicate_bullets(
            self._filter_general_bullets(
                self._bullets_for_output(result, has_kept_findings=True),
                active_general_types=active_types,
            )
        )
        if not bullets:
            bullets = [finding.message.strip() for finding in findings if finding.message.strip()]
        note_parts = [summary]
        note_parts.extend(f"- {bullet}" for bullet in bullets)
        return summary, bullets, "\n".join(note_parts)

    @staticmethod
    def _format_prompt_override(prompt_override: Optional[str]) -> str:
        normalized = (prompt_override or "").strip()
        if not normalized:
            return ""
        return f"\n\nAdditional run-specific user QC instructions:\n{normalized}"

    @staticmethod
    def _prompt_override_content(prompt_override: Optional[str]) -> list[dict]:
        normalized = (prompt_override or "").strip()
        if not normalized:
            return []
        return [
            {
                "type": "input_text",
                "text": (
                    "Additional run-specific QC instructions from the user. Follow these if they do not conflict "
                    f"with the base comparison rules:\n{normalized}"
                ),
            }
        ]

    def _merge_comments(
        self,
        general_result: _SlideQcSchema,
        text_result: _SlideQcSchema,
        general_kept_findings: list[SlideQcFinding],
        text_kept_findings: list[SlideQcFinding],
        findings: list[SlideQcFinding],
    ) -> tuple[Optional[str], list[str], Optional[str]]:
        general_summary = self._summary_for_output(general_result, has_kept_findings=bool(general_kept_findings))
        text_summary = self._summary_for_output(text_result, has_kept_findings=bool(text_kept_findings))
        if general_summary and text_summary and general_summary != text_summary:
            summary = "AI review findings"
        else:
            summary = general_summary or text_summary

        active_general_types = {finding.finding_type for finding in general_kept_findings}
        bullets = self._deduplicate_bullets(
            self._filter_general_bullets(
                self._bullets_for_output(general_result, has_kept_findings=bool(general_kept_findings)),
                active_general_types=active_general_types,
            )
            + self._bullets_for_output(text_result, has_kept_findings=bool(text_kept_findings))
        )
        if not bullets and findings:
            bullets = [finding.message.strip() for finding in findings if finding.message.strip()]
        if summary is None and bullets:
            summary = "AI review findings"
        note_parts: list[str] = []
        if summary:
            note_parts.append(summary)
        note_parts.extend(f"- {bullet}" for bullet in bullets)
        note = "\n".join(note_parts) if note_parts else None
        return summary, bullets, note

    @staticmethod
    def _summary_for_output(result: _SlideQcSchema, *, has_kept_findings: bool) -> Optional[str]:
        if not has_kept_findings:
            return None
        if result.summary and result.summary.strip():
            return result.summary.strip()
        if result.note and result.note.strip():
            return result.note.strip()
        return None

    @staticmethod
    def _bullets_for_output(result: _SlideQcSchema, *, has_kept_findings: bool) -> list[str]:
        if not has_kept_findings:
            return []
        return [bullet.strip() for bullet in result.bullets if bullet and bullet.strip()]

    @staticmethod
    def _deduplicate_bullets(bullets: list[str]) -> list[str]:
        deduped: list[str] = []
        seen: set[str] = set()
        for bullet in bullets:
            key = bullet.casefold()
            if key in seen:
                continue
            seen.add(key)
            deduped.append(bullet)
        return deduped

    @staticmethod
    def _filter_general_bullets(
        bullets: list[str],
        *,
        active_general_types: set[QcFindingType],
    ) -> list[str]:
        filtered: list[str] = []
        for bullet in bullets:
            lower = bullet.casefold()
            if QcFindingType.WRONG_COLOR not in active_general_types and any(
                token in lower for token in ("color", "colour", "shade", "font color", "text color")
            ):
                continue
            if any(
                token in lower
                for token in (
                    "position",
                    "alignment",
                    "aligned",
                    "shifted",
                    "too far",
                    "narrower",
                    "wider",
                    "wrap",
                    "wrapped",
                    "line break",
                    "line-break",
                )
            ):
                continue
            filtered.append(bullet)
        return filtered

    @staticmethod
    def _merge_statuses(
        general_status: str,
        text_status: str,
        findings: list[SlideQcFinding],
    ) -> SlideQcStatus:
        if findings:
            return SlideQcStatus.FINDINGS
        if general_status == SlideQcStatus.MANUAL_REVIEW.value or text_status == SlideQcStatus.MANUAL_REVIEW.value:
            return SlideQcStatus.MANUAL_REVIEW
        return SlideQcStatus.OK

    @staticmethod
    def _expand_bbox(
        bbox: tuple[float, float, float, float],
        *,
        margin: float = 0.03,
        min_size: float = 0.12,
    ) -> tuple[float, float, float, float]:
        x0, y0, x1, y1 = bbox
        x0 = max(0.0, x0 - margin)
        y0 = max(0.0, y0 - margin)
        x1 = min(1.0, x1 + margin)
        y1 = min(1.0, y1 + margin)
        width = x1 - x0
        height = y1 - y0
        if width < min_size:
            center_x = (x0 + x1) / 2.0
            half = min_size / 2.0
            x0 = max(0.0, center_x - half)
            x1 = min(1.0, center_x + half)
        if height < min_size:
            center_y = (y0 + y1) / 2.0
            half = min_size / 2.0
            y0 = max(0.0, center_y - half)
            y1 = min(1.0, center_y + half)
        return (x0, y0, x1, y1)

    @staticmethod
    def _crop_to_bbox(image: Image.Image, bbox: tuple[float, float, float, float]) -> Image.Image:
        x0 = max(0, min(image.width - 1, int(round(bbox[0] * image.width))))
        y0 = max(0, min(image.height - 1, int(round(bbox[1] * image.height))))
        x1 = max(x0 + 1, min(image.width, int(round(bbox[2] * image.width))))
        y1 = max(y0 + 1, min(image.height, int(round(bbox[3] * image.height))))
        return image.crop((x0, y0, x1, y1))

    @staticmethod
    def _bbox_iou(left: tuple[float, float, float, float], right: tuple[float, float, float, float]) -> float:
        inter_x0 = max(left[0], right[0])
        inter_y0 = max(left[1], right[1])
        inter_x1 = min(left[2], right[2])
        inter_y1 = min(left[3], right[3])
        inter_width = max(0.0, inter_x1 - inter_x0)
        inter_height = max(0.0, inter_y1 - inter_y0)
        inter_area = inter_width * inter_height
        if inter_area <= 0.0:
            return 0.0
        left_area = max(0.0, left[2] - left[0]) * max(0.0, left[3] - left[1])
        right_area = max(0.0, right[2] - right[0]) * max(0.0, right[3] - right[1])
        union_area = left_area + right_area - inter_area
        if union_area <= 0.0:
            return 0.0
        return inter_area / union_area

    @staticmethod
    def _bbox_center(bbox: tuple[float, float, float, float]) -> tuple[float, float]:
        return ((bbox[0] + bbox[2]) / 2.0, (bbox[1] + bbox[3]) / 2.0)

    @staticmethod
    def _bbox_contains_point(bbox: tuple[float, float, float, float], point: tuple[float, float]) -> bool:
        return bbox[0] <= point[0] <= bbox[2] and bbox[1] <= point[1] <= bbox[3]

    @staticmethod
    def _normalize_bbox(raw_bbox: list[float]) -> tuple[float, float, float, float]:
        if len(raw_bbox) != 4:
            return (0.0, 0.0, 1.0, 1.0)
        x0, y0, x1, y1 = [max(0.0, min(1.0, float(value))) for value in raw_bbox]
        return (min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1))
