from __future__ import annotations

import numpy as np

from app.services.models import (
    PageAlignment,
    ParagraphLayout,
    QcFindingType,
    TextBox,
    TextLayout,
    TextSource,
    VisualDiff,
    VisualDiffRegion,
)
from app.services.qc_detector import QcDetector


def _line(text: str, bbox: tuple[float, float, float, float]) -> TextBox:
    return TextBox(text=text, bbox=bbox, page_number=1, confidence=1.0, source=TextSource.NATIVE)


def _paragraph(text: str, bbox: tuple[float, float, float, float], line_count: int = 1) -> ParagraphLayout:
    lines = []
    for index in range(line_count):
        offset = index * 0.03
        lines.append(_line(text if line_count == 1 else f"{text} {index + 1}", (bbox[0], bbox[1] + offset, bbox[2], bbox[3] + offset)))
    return ParagraphLayout(
        text=text,
        lines=lines,
        bbox=bbox,
        page_number=1,
        confidence=1.0,
        source=TextSource.NATIVE,
    )


def _layout(paragraphs: list[ParagraphLayout]) -> TextLayout:
    return TextLayout(
        page_number=1,
        page_size=(1000.0, 750.0),
        paragraphs=paragraphs,
        lines=[line for paragraph in paragraphs for line in paragraph.lines],
        source=TextSource.NATIVE,
        total_characters=sum(1 for paragraph in paragraphs for character in paragraph.text if not character.isspace()),
        average_confidence=1.0,
        extracted_with_ocr=False,
    )


def _visual_diff(*regions: VisualDiffRegion, similarity_score: float = 0.94) -> VisualDiff:
    return VisualDiff(
        page_index=0,
        alignment=PageAlignment(
            page_index=0,
            angle=0,
            aligned_reference_image=np.zeros((10, 10, 3), dtype=np.uint8),
            similarity_score=similarity_score,
            target_size=(1000, 750),
            offset=(0, 0),
            scaled_size=(1000, 750),
        ),
        missing_regions=list(regions),
    )


def test_qc_detector_flags_line_break_issues() -> None:
    detector = QcDetector()
    reference_layout = _layout([_paragraph("Alpha beta gamma", (0.1, 0.1, 0.45, 0.18), line_count=1)])
    candidate_layout = _layout([_paragraph("Alpha beta gamma", (0.1, 0.1, 0.45, 0.24), line_count=2)])

    result = detector.detect(reference_layout, candidate_layout, _visual_diff())

    assert result.status.value == "findings"
    assert any(finding.finding_type == QcFindingType.LINE_BREAK_ISSUE for finding in result.findings)


def test_qc_detector_flags_alignment_drift() -> None:
    detector = QcDetector()
    reference_layout = _layout([_paragraph("Aligned headline", (0.10, 0.10, 0.42, 0.18))])
    candidate_layout = _layout([_paragraph("Aligned headline", (0.16, 0.10, 0.48, 0.18))])

    result = detector.detect(reference_layout, candidate_layout, _visual_diff())

    assert any(finding.finding_type == QcFindingType.SIZE_POSITION_ISSUE for finding in result.findings)


def test_qc_detector_flags_missing_visual_content() -> None:
    detector = QcDetector()
    reference_layout = _layout([])
    candidate_layout = _layout([])

    result = detector.detect(
        reference_layout,
        candidate_layout,
        _visual_diff(VisualDiffRegion(bbox=(0.60, 0.20, 0.82, 0.42), area_ratio=0.02)),
    )

    assert any(finding.finding_type == QcFindingType.MISSING_CONTENT for finding in result.findings)


def test_qc_detector_flags_geometry_difference_regions() -> None:
    detector = QcDetector()
    reference_layout = _layout([])
    candidate_layout = _layout([])

    result = detector.detect(
        reference_layout,
        candidate_layout,
        VisualDiff(
            page_index=0,
            alignment=PageAlignment(
                page_index=0,
                angle=0,
                aligned_reference_image=np.zeros((10, 10, 3), dtype=np.uint8),
                similarity_score=0.94,
                target_size=(1000, 750),
                offset=(0, 0),
                scaled_size=(1000, 750),
            ),
            geometry_regions=[VisualDiffRegion(bbox=(0.20, 0.22, 0.48, 0.38), area_ratio=0.018)],
        ),
    )

    assert any(finding.finding_type == QcFindingType.SIZE_POSITION_ISSUE for finding in result.findings)


def test_qc_detector_flags_wrong_color_regions() -> None:
    detector = QcDetector()
    reference_layout = _layout([])
    candidate_layout = _layout([])

    result = detector.detect(
        reference_layout,
        candidate_layout,
        VisualDiff(
            page_index=0,
            alignment=PageAlignment(
                page_index=0,
                angle=0,
                aligned_reference_image=np.zeros((10, 10, 3), dtype=np.uint8),
                similarity_score=0.94,
                target_size=(1000, 750),
                offset=(0, 0),
                scaled_size=(1000, 750),
            ),
            color_regions=[VisualDiffRegion(bbox=(0.52, 0.22, 0.70, 0.36), area_ratio=0.015)],
        ),
    )

    assert any(finding.finding_type == QcFindingType.WRONG_COLOR for finding in result.findings)


def test_qc_detector_marks_manual_review_when_alignment_is_too_low() -> None:
    detector = QcDetector()
    reference_layout = _layout([_paragraph("Alpha beta gamma", (0.1, 0.1, 0.45, 0.18))])
    candidate_layout = _layout([_paragraph("Alpha beta gamma", (0.1, 0.1, 0.45, 0.18))])

    result = detector.detect(reference_layout, candidate_layout, _visual_diff(similarity_score=0.12))

    assert result.status.value == "manual_review"
    assert result.note


def test_qc_detector_keeps_image_only_reference_as_advisory_not_manual_review() -> None:
    detector = QcDetector()
    reference_layout = _layout([])
    candidate_layout = _layout([_paragraph("Alpha beta gamma", (0.1, 0.1, 0.45, 0.18))])

    result = detector.detect(reference_layout, candidate_layout, _visual_diff())

    assert result.status.value == "ok"
    assert result.note == "Visual-only QC was used because the reference PDF does not expose text."
