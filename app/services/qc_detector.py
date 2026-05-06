from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from difflib import SequenceMatcher
from math import hypot

from app.services.models import (
    NormalizedBBox,
    ParagraphLayout,
    QcFindingSeverity,
    QcFindingType,
    SlideQcFinding,
    SlideQcResult,
    SlideQcStatus,
    TextLayout,
    VisualDiff,
)


@dataclass
class _ParagraphMatch:
    reference: ParagraphLayout
    candidate: ParagraphLayout
    score: float


class QcDetector:
    def __init__(
        self,
        paragraph_match_threshold: float = 0.72,
        strong_text_match_threshold: float = 0.92,
        line_break_exact_threshold: float = 0.98,
        min_alignment_confidence: float = 0.28,
        min_ocr_confidence: float = 0.55,
        drift_threshold_x: float = 0.026,
        drift_threshold_y: float = 0.028,
        drift_threshold_center: float = 0.036,
        drift_threshold_width: float = 0.05,
    ):
        self.paragraph_match_threshold = paragraph_match_threshold
        self.strong_text_match_threshold = strong_text_match_threshold
        self.line_break_exact_threshold = line_break_exact_threshold
        self.min_alignment_confidence = min_alignment_confidence
        self.min_ocr_confidence = min_ocr_confidence
        self.drift_threshold_x = drift_threshold_x
        self.drift_threshold_y = drift_threshold_y
        self.drift_threshold_center = drift_threshold_center
        self.drift_threshold_width = drift_threshold_width

    def detect(
        self,
        reference_layout: TextLayout,
        candidate_layout: TextLayout,
        visual_diff: VisualDiff,
    ) -> SlideQcResult:
        findings: list[SlideQcFinding] = []
        manual_review_note: str | None = None
        advisory_note: str | None = None
        next_finding_id = 1
        reference_is_visual_only = reference_layout.total_characters == 0 and not reference_layout.extracted_with_ocr

        if visual_diff.alignment.similarity_score < self.min_alignment_confidence:
            manual_review_note = "Page alignment confidence is too low for automated QC."

        if reference_layout.extracted_with_ocr and reference_layout.average_confidence < self.min_ocr_confidence:
            manual_review_note = "Reference OCR confidence is too low for reliable automated QC."

        if reference_is_visual_only:
            advisory_note = "Visual-only QC was used because the reference PDF does not expose text."

        matches, unmatched_reference = self._match_paragraphs(reference_layout.paragraphs, candidate_layout.paragraphs)
        mapped_reference_boxes = [self._map_reference_bbox(paragraph.bbox, visual_diff) for paragraph in reference_layout.paragraphs]

        for paragraph in unmatched_reference:
            mapped_bbox = self._map_reference_bbox(paragraph.bbox, visual_diff)
            findings.append(
                self._build_finding(
                    finding_id=next_finding_id,
                    finding_type=QcFindingType.MISSING_CONTENT,
                    severity=QcFindingSeverity.HIGH,
                    bbox=mapped_bbox,
                    message="Text or content from the PDF appears to be missing on this slide.",
                    confidence=0.88,
                )
            )
            next_finding_id += 1

        for match in matches:
            reference_text = self._normalize_text(match.reference.text)
            candidate_text = self._normalize_text(match.candidate.text)
            reference_line_count = len([line for line in match.reference.lines if line.text.strip()])
            candidate_line_count = len([line for line in match.candidate.lines if line.text.strip()])

            if (
                reference_text
                and candidate_text
                and match.score >= self.line_break_exact_threshold
                and candidate_line_count > reference_line_count
            ):
                findings.append(
                    self._build_finding(
                        finding_id=next_finding_id,
                        finding_type=QcFindingType.LINE_BREAK_ISSUE,
                        severity=QcFindingSeverity.MEDIUM,
                        bbox=match.candidate.bbox,
                        message=(
                            f"Paragraph wraps into {candidate_line_count} lines on the slide versus "
                            f"{reference_line_count} lines in the PDF."
                        ),
                        confidence=min(0.96, 0.80 + ((candidate_line_count - reference_line_count) * 0.06)),
                    )
                )
                next_finding_id += 1
                continue

            if match.score >= self.strong_text_match_threshold:
                reference_bbox = self._map_reference_bbox(match.reference.bbox, visual_diff)
                if self._is_alignment_drift(reference_bbox, match.candidate.bbox):
                    findings.append(
                        self._build_finding(
                            finding_id=next_finding_id,
                            finding_type=QcFindingType.SIZE_POSITION_ISSUE,
                            severity=QcFindingSeverity.MEDIUM,
                            bbox=self._union_bbox(reference_bbox, match.candidate.bbox),
                            message="Matched content is in the wrong size or position compared with the PDF reference.",
                            confidence=min(0.94, max(match.score, 0.82)),
                        )
                    )
                    next_finding_id += 1

        for region in visual_diff.missing_regions:
            if any(self._bbox_overlap_ratio(region.bbox, bbox) >= 0.35 for bbox in mapped_reference_boxes):
                continue
            findings.append(
                self._build_finding(
                    finding_id=next_finding_id,
                    finding_type=QcFindingType.MISSING_CONTENT,
                    severity=QcFindingSeverity.HIGH,
                    bbox=region.bbox,
                    message="A visual element from the PDF appears to be missing on this slide.",
                    confidence=min(0.95, 0.72 + (region.area_ratio * 18)),
                )
            )
            next_finding_id += 1

        existing_boxes = [finding.bbox for finding in findings]
        for region in visual_diff.geometry_regions:
            if any(self._bbox_overlap_ratio(region.bbox, bbox) >= 0.55 for bbox in existing_boxes):
                continue
            findings.append(
                self._build_finding(
                    finding_id=next_finding_id,
                    finding_type=QcFindingType.SIZE_POSITION_ISSUE,
                    severity=QcFindingSeverity.MEDIUM,
                    bbox=region.bbox,
                    message="An element appears to be in the wrong size or position compared with the PDF reference.",
                    confidence=min(0.93, 0.70 + (region.area_ratio * 16)),
                )
            )
            existing_boxes.append(region.bbox)
            next_finding_id += 1

        for region in visual_diff.color_regions:
            if any(self._bbox_overlap_ratio(region.bbox, bbox) >= 0.55 for bbox in existing_boxes):
                continue
            findings.append(
                self._build_finding(
                    finding_id=next_finding_id,
                    finding_type=QcFindingType.WRONG_COLOR,
                    severity=QcFindingSeverity.MEDIUM,
                    bbox=region.bbox,
                    message="An element appears to use the wrong color compared with the PDF reference.",
                    confidence=min(0.92, 0.72 + (region.area_ratio * 16)),
                )
            )
            existing_boxes.append(region.bbox)
            next_finding_id += 1

        deduped_findings = self._dedupe_findings(findings)
        status = SlideQcStatus.OK
        note = advisory_note
        if manual_review_note is not None:
            status = SlideQcStatus.MANUAL_REVIEW
            note = manual_review_note
        elif deduped_findings:
            status = SlideQcStatus.FINDINGS

        return SlideQcResult(
            slide_index=candidate_layout.page_number - 1,
            page_index=reference_layout.page_number - 1,
            status=status,
            findings=deduped_findings,
            alignment_confidence=visual_diff.alignment.similarity_score,
            reference_source=reference_layout.source,
            candidate_source=candidate_layout.source,
            note=note,
        )

    def _match_paragraphs(
        self,
        reference_paragraphs: list[ParagraphLayout],
        candidate_paragraphs: list[ParagraphLayout],
    ) -> tuple[list[_ParagraphMatch], list[ParagraphLayout]]:
        matches: list[_ParagraphMatch] = []
        used_candidate_indexes: set[int] = set()
        unmatched_reference: list[ParagraphLayout] = []

        for reference_paragraph in reference_paragraphs:
            best_index = -1
            best_score = 0.0
            normalized_reference = self._normalize_text(reference_paragraph.text)

            for candidate_index, candidate_paragraph in enumerate(candidate_paragraphs):
                if candidate_index in used_candidate_indexes:
                    continue
                normalized_candidate = self._normalize_text(candidate_paragraph.text)
                if not normalized_reference or not normalized_candidate:
                    continue
                score = SequenceMatcher(None, normalized_reference, normalized_candidate).ratio()
                if score > best_score:
                    best_score = score
                    best_index = candidate_index

            if best_index >= 0 and best_score >= self.paragraph_match_threshold:
                used_candidate_indexes.add(best_index)
                matches.append(
                    _ParagraphMatch(
                        reference=reference_paragraph,
                        candidate=candidate_paragraphs[best_index],
                        score=best_score,
                    )
                )
            else:
                unmatched_reference.append(reference_paragraph)

        return matches, unmatched_reference

    def _is_alignment_drift(self, reference_bbox: NormalizedBBox, candidate_bbox: NormalizedBBox) -> bool:
        reference_center = self._bbox_center(reference_bbox)
        candidate_center = self._bbox_center(candidate_bbox)
        left_delta = abs(reference_bbox[0] - candidate_bbox[0])
        top_delta = abs(reference_bbox[1] - candidate_bbox[1])
        width_delta = abs((reference_bbox[2] - reference_bbox[0]) - (candidate_bbox[2] - candidate_bbox[0]))
        center_delta = hypot(reference_center[0] - candidate_center[0], reference_center[1] - candidate_center[1])
        return (
            left_delta >= self.drift_threshold_x
            or top_delta >= self.drift_threshold_y
            or width_delta >= self.drift_threshold_width
            or center_delta >= self.drift_threshold_center
        )

    @staticmethod
    def _normalize_text(value: str) -> str:
        normalized = unicodedata.normalize("NFKC", value or "")
        normalized = normalized.replace("’", "'").replace("“", '"').replace("”", '"').replace("–", "-").replace("—", "-")
        normalized = re.sub(r"\s+", " ", normalized)
        return normalized.strip().casefold()

    @staticmethod
    def _map_reference_bbox(bbox: NormalizedBBox, visual_diff: VisualDiff) -> NormalizedBBox:
        angle = visual_diff.alignment.angle
        target_width, target_height = visual_diff.alignment.target_size
        offset_x, offset_y = visual_diff.alignment.offset
        scaled_width, scaled_height = visual_diff.alignment.scaled_size

        corners = [
            (bbox[0], bbox[1]),
            (bbox[2], bbox[1]),
            (bbox[2], bbox[3]),
            (bbox[0], bbox[3]),
        ]
        rotated_corners = [QcDetector._rotate_point(x, y, angle) for x, y in corners]
        xs = [(offset_x + (x * scaled_width)) / max(target_width, 1) for x, _y in rotated_corners]
        ys = [(offset_y + (y * scaled_height)) / max(target_height, 1) for _x, y in rotated_corners]
        return (
            QcDetector._clamp(min(xs)),
            QcDetector._clamp(min(ys)),
            QcDetector._clamp(max(xs)),
            QcDetector._clamp(max(ys)),
        )

    @staticmethod
    def _rotate_point(x: float, y: float, angle: int) -> tuple[float, float]:
        normalized_angle = angle % 360
        if normalized_angle == 0:
            return (x, y)
        if normalized_angle == 90:
            return (y, 1.0 - x)
        if normalized_angle == 180:
            return (1.0 - x, 1.0 - y)
        if normalized_angle == 270:
            return (1.0 - y, x)
        raise ValueError(f"Unsupported rotation angle: {angle}")

    @staticmethod
    def _build_finding(
        finding_id: int,
        finding_type: QcFindingType,
        severity: QcFindingSeverity,
        bbox: NormalizedBBox,
        message: str,
        confidence: float,
    ) -> SlideQcFinding:
        return SlideQcFinding(
            finding_id=finding_id,
            finding_type=finding_type,
            severity=severity,
            bbox=bbox,
            message=message,
            confidence=max(0.0, min(1.0, confidence)),
        )

    @staticmethod
    def _bbox_center(bbox: NormalizedBBox) -> tuple[float, float]:
        return ((bbox[0] + bbox[2]) / 2.0, (bbox[1] + bbox[3]) / 2.0)

    @staticmethod
    def _bbox_overlap_ratio(left: NormalizedBBox, right: NormalizedBBox) -> float:
        overlap_x0 = max(left[0], right[0])
        overlap_y0 = max(left[1], right[1])
        overlap_x1 = min(left[2], right[2])
        overlap_y1 = min(left[3], right[3])
        overlap_width = max(0.0, overlap_x1 - overlap_x0)
        overlap_height = max(0.0, overlap_y1 - overlap_y0)
        overlap_area = overlap_width * overlap_height
        left_area = max(0.0, left[2] - left[0]) * max(0.0, left[3] - left[1])
        if left_area <= 0:
            return 0.0
        return overlap_area / left_area

    @staticmethod
    def _bbox_iou(left: NormalizedBBox, right: NormalizedBBox) -> float:
        overlap_x0 = max(left[0], right[0])
        overlap_y0 = max(left[1], right[1])
        overlap_x1 = min(left[2], right[2])
        overlap_y1 = min(left[3], right[3])
        overlap_width = max(0.0, overlap_x1 - overlap_x0)
        overlap_height = max(0.0, overlap_y1 - overlap_y0)
        overlap_area = overlap_width * overlap_height
        left_area = max(0.0, left[2] - left[0]) * max(0.0, left[3] - left[1])
        right_area = max(0.0, right[2] - right[0]) * max(0.0, right[3] - right[1])
        denominator = left_area + right_area - overlap_area
        if denominator <= 0:
            return 0.0
        return overlap_area / denominator

    @staticmethod
    def _union_bbox(left: NormalizedBBox, right: NormalizedBBox) -> NormalizedBBox:
        return (
            min(left[0], right[0]),
            min(left[1], right[1]),
            max(left[2], right[2]),
            max(left[3], right[3]),
        )

    def _dedupe_findings(self, findings: list[SlideQcFinding]) -> list[SlideQcFinding]:
        deduped: list[SlideQcFinding] = []
        for finding in findings:
            duplicate = False
            for existing in deduped:
                if existing.finding_type != finding.finding_type:
                    continue
                if self._bbox_iou(existing.bbox, finding.bbox) >= 0.72:
                    duplicate = True
                    break
            if not duplicate:
                deduped.append(finding)
        return deduped

    @staticmethod
    def _clamp(value: float) -> float:
        return max(0.0, min(1.0, value))
