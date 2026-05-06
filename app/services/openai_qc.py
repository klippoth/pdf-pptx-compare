from __future__ import annotations

import base64
from io import BytesIO
from pathlib import Path
from typing import Literal, Optional

from PIL import Image
from pydantic import BaseModel, Field

from app.services.models import (
    PageImage,
    QcFindingSeverity,
    QcFindingType,
    SlideQcFinding,
    SlideQcResult,
    SlideQcStatus,
    TextSource,
)


class _FindingSchema(BaseModel):
    type: Literal["missing_content", "extra_content", "wrong_text", "wrong_color", "line_break_issue"]
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


class OpenAIQCEvaluator:
    def __init__(
        self,
        *,
        api_key: Optional[str],
        model: str = "gpt-5.3-chat-latest",
        timeout_seconds: float = 90.0,
        max_image_dimension: int = 1400,
        client=None,
    ):
        self.api_key = api_key
        self.model = model
        self.timeout_seconds = timeout_seconds
        self.max_image_dimension = max_image_dimension
        self._client = client

    def is_available(self) -> bool:
        return bool(self.api_key or self._client is not None)

    def compare_pages(
        self,
        *,
        slide_index: int,
        page_index: int,
        reference_page: PageImage,
        candidate_page: PageImage,
        debug_output_dir: Optional[Path] = None,
    ) -> SlideQcResult:
        parsed = self._invoke_model(
            reference_page=reference_page,
            candidate_page=candidate_page,
            debug_output_dir=debug_output_dir,
        )
        findings: list[SlideQcFinding] = []
        for finding in parsed.findings:
            finding_type = QcFindingType(finding.type)
            severity = QcFindingSeverity(finding.severity)
            bbox = self._normalize_bbox(finding.bbox)
            findings.append(
                SlideQcFinding(
                    finding_id=len(findings) + 1,
                    finding_type=finding_type,
                    severity=severity,
                    bbox=bbox,
                    message=finding.message.strip(),
                    confidence=max(0.0, min(1.0, float(finding.confidence))),
                )
            )

        status = SlideQcStatus(parsed.status)
        if status == SlideQcStatus.OK and findings:
            status = SlideQcStatus.FINDINGS
        if status == SlideQcStatus.FINDINGS and not findings:
            status = SlideQcStatus.OK
        summary, bullets, note = self._final_comment(parsed.summary, parsed.bullets, parsed.note, findings)

        return SlideQcResult(
            slide_index=slide_index,
            page_index=page_index,
            status=status,
            findings=findings,
            alignment_confidence=max(0.0, min(1.0, float(parsed.comparison_confidence))),
            reference_source=TextSource.MODEL,
            candidate_source=TextSource.MODEL,
            summary=summary,
            comment_bullets=bullets,
            note=note,
        )

    def _invoke_model(
        self,
        *,
        reference_page: PageImage,
        candidate_page: PageImage,
        debug_output_dir: Optional[Path] = None,
    ) -> _SlideQcSchema:
        client = self._get_client()
        candidate_image, reference_image = self._prepare_comparison_images(
            candidate_page=candidate_page,
            reference_page=reference_page,
        )
        if debug_output_dir is not None:
            self._save_debug_inputs(
                debug_output_dir=debug_output_dir,
                candidate_image=candidate_image,
                reference_image=reference_image,
            )
        response = client.responses.parse(
            model=self.model,
            input=[
                {
                    "role": "system",
                    "content": (
                        "You are a slide quality-control reviewer. Compare a reference PDF page image against a "
                        "candidate PowerPoint slide render. The PDF reference is always the source of truth. "
                        "Your job is to comment only on obvious content mistakes. Only report findings in these "
                        "five categories: missing_content, extra_content, wrong_text, wrong_color, and "
                        "line_break_issue. Treat missing_content as shapes or elements that are present in the "
                        "reference but absent in the candidate. Treat extra_content as shapes or elements that "
                        "appear in the candidate but should not be there. Treat wrong_text as clearly incorrect "
                        "wording, numbers, or labels. Treat line_break_issue only as an obvious wrap/split where "
                        "the same text should stay on one line or block in the candidate. "
                        "Do not report size_position_issue, alignment drift, subtle layout drift, tiny spacing "
                        "changes, minor font differences, tiny kerning changes, mild rendering noise, or negligible "
                        "sub-pixel shifts. If a difference is borderline, ignore it. "
                        "Return bounding boxes in normalized candidate-image coordinates [x0, y0, x1, y1] "
                        "between 0 and 1. Provide a short summary plus a bullet list. Each bullet should describe "
                        "one obvious issue with a bit of location or context. Use manual_review only if the images "
                        "are too ambiguous to compare."
                    ),
                },
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "input_text",
                            "text": (
                                "Image 1 is the candidate PowerPoint slide. Image 2 is the reference PDF page. "
                                "Both images have been normalized to the same canvas size for easier comparison. "
                                "Compare them visually, page by page. "
                                "Focus only on: missing shapes/elements, shapes that should not be there, clearly "
                                "wrong text, obviously wrong colors, and obvious line breaks where the text should "
                                "stay together. Ignore all subtle alignment, positioning, sizing, spacing, or font "
                                "differences. If a line break call is not extremely clear visually, do not report it. "
                                "Use Image 1 for the returned bounding boxes. If there are obvious "
                                "issues, write them as concise bullets with a little more detail than a label alone. "
                                "If there are no obvious issues, say so."
                            ),
                        },
                        {"type": "input_image", "image_url": self._image_to_data_url(candidate_image)},
                        {"type": "input_image", "image_url": self._image_to_data_url(reference_image)},
                    ],
                },
            ],
            text_format=_SlideQcSchema,
            timeout=self.timeout_seconds,
        )
        return response.output_parsed

    def _get_client(self):
        if self._client is not None:
            return self._client
        from openai import OpenAI

        self._client = OpenAI(api_key=self.api_key, timeout=self.timeout_seconds, max_retries=2)
        return self._client

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
        return (
            candidate_scaled,
            self._fit_image_to_canvas(reference_scaled, candidate_scaled.size),
        )

    def _scale_image(self, image: Image.Image) -> Image.Image:
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
    def _fit_image_to_canvas(image: Image.Image, target_size: tuple[int, int]) -> Image.Image:
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
        return canvas

    def _image_to_data_url(self, image: Image.Image) -> str:
        buffer = BytesIO()
        image.save(buffer, format="PNG", optimize=True)
        encoded = base64.b64encode(buffer.getvalue()).decode("ascii")
        return f"data:image/png;base64,{encoded}"

    @staticmethod
    def _save_debug_inputs(
        *,
        debug_output_dir: Path,
        candidate_image: Image.Image,
        reference_image: Image.Image,
    ) -> None:
        debug_output_dir.mkdir(parents=True, exist_ok=True)
        candidate_image.save(debug_output_dir / "01-candidate-slide.png")
        reference_image.save(debug_output_dir / "02-reference-page.png")

    @staticmethod
    def _final_comment(
        raw_summary: Optional[str],
        raw_bullets: list[str],
        raw_note: Optional[str],
        findings: list[SlideQcFinding],
    ) -> tuple[Optional[str], list[str], Optional[str]]:
        summary = raw_summary.strip() if raw_summary and raw_summary.strip() else None

        bullets = [bullet.strip() for bullet in raw_bullets if bullet and bullet.strip()]
        if not bullets and findings:
            bullets = [finding.message.strip() for finding in findings if finding.message.strip()]
        if summary is None and raw_note and raw_note.strip():
            summary = raw_note.strip()
        if summary is None and bullets:
            summary = "AI review findings"
        note_parts: list[str] = []
        if summary:
            note_parts.append(summary)
        note_parts.extend(f"- {bullet}" for bullet in bullets)
        note = "\n".join(note_parts) if note_parts else None
        return summary, bullets, note

    @staticmethod
    def _normalize_bbox(raw_bbox: list[float]) -> tuple[float, float, float, float]:
        if len(raw_bbox) != 4:
            return (0.0, 0.0, 1.0, 1.0)
        x0, y0, x1, y1 = [max(0.0, min(1.0, float(value))) for value in raw_bbox]
        return (min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1))
