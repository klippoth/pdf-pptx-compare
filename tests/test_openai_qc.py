from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import numpy as np

from app.services.models import PageImage, QcFindingType, SlideQcStatus
from app.services.openai_qc import OpenAIQCEvaluator


class _FakeResponses:
    def __init__(self, parsed_payload):
        self.parsed_payload = parsed_payload
        self.last_kwargs = None

    def parse(self, **kwargs):
        self.last_kwargs = kwargs
        return SimpleNamespace(output_parsed=self.parsed_payload)


class _FakeClient:
    def __init__(self, parsed_payload):
        self.responses = _FakeResponses(parsed_payload)


def _page(page_index: int = 0) -> PageImage:
    image = np.full((240, 320, 3), 255, dtype=np.uint8)
    return PageImage(page_index=page_index, image=image, image_path=Path(f"/tmp/page-{page_index}.png"))


def _page_with_size(width: int, height: int, page_index: int = 0) -> PageImage:
    image = np.full((height, width, 3), 255, dtype=np.uint8)
    return PageImage(page_index=page_index, image=image, image_path=Path(f"/tmp/page-{page_index}-{width}x{height}.png"))


def test_openai_qc_evaluator_maps_structured_response_to_slide_qc_result() -> None:
    parsed_payload = SimpleNamespace(
        status="findings",
        summary="Slide differences found",
        bullets=[
            "Missing logo in the top-right corner",
            "Headline text differs from the PDF reference",
            "Extra shape appears near the footer",
        ],
        note="",
        comparison_confidence=0.91,
        findings=[
            SimpleNamespace(
                type="missing_content",
                severity="high",
                message="Logo is missing.",
                confidence=0.96,
                bbox=[0.11, 0.14, 0.23, 0.28],
            ),
            SimpleNamespace(
                type="wrong_text",
                severity="medium",
                message="Headline text differs from the PDF reference.",
                confidence=0.88,
                bbox=[0.24, 0.18, 0.88, 0.32],
            ),
            SimpleNamespace(
                type="extra_content",
                severity="medium",
                message="Extra shape appears near the footer.",
                confidence=0.81,
                bbox=[0.62, 0.72, 0.94, 0.84],
            ),
        ],
    )
    fake_client = _FakeClient(parsed_payload)
    evaluator = OpenAIQCEvaluator(api_key="test", client=fake_client)

    result = evaluator.compare_pages(slide_index=2, page_index=2, reference_page=_page(2), candidate_page=_page(2))

    assert result.status == SlideQcStatus.FINDINGS
    assert result.alignment_confidence == 0.91
    assert result.summary == "Slide differences found"
    assert result.comment_bullets == [
        "Missing logo in the top-right corner",
        "Headline text differs from the PDF reference",
        "Extra shape appears near the footer",
    ]
    assert [finding.finding_type for finding in result.findings] == [
        QcFindingType.MISSING_CONTENT,
        QcFindingType.WRONG_TEXT,
        QcFindingType.EXTRA_CONTENT,
    ]
    assert fake_client.responses.last_kwargs["model"] == "gpt-5.3-chat-latest"
    assert "temperature" not in fake_client.responses.last_kwargs
    content = fake_client.responses.last_kwargs["input"][1]["content"]
    image_items = [item for item in content if item["type"] == "input_image"]
    assert len(image_items) == 2


def test_openai_qc_evaluator_exposes_unavailability_without_key() -> None:
    evaluator = OpenAIQCEvaluator(api_key=None)
    assert evaluator.is_available() is False


def test_openai_qc_evaluator_keeps_obvious_line_break_findings_and_comments() -> None:
    parsed_payload = SimpleNamespace(
        status="findings",
        summary="Major text wrapping issues compared to the reference",
        bullets=[
            "The main title is incorrectly split across many short lines.",
            "The subtitle wraps to two lines.",
        ],
        note="",
        comparison_confidence=0.82,
        findings=[
            SimpleNamespace(
                type="line_break_issue",
                severity="medium",
                message="Title wraps into multiple lines.",
                confidence=0.84,
                bbox=[0.10, 0.12, 0.78, 0.55],
            ),
        ],
    )
    fake_client = _FakeClient(parsed_payload)
    evaluator = OpenAIQCEvaluator(api_key="test", client=fake_client)

    result = evaluator.compare_pages(slide_index=0, page_index=0, reference_page=_page(0), candidate_page=_page(0))

    assert result.status == SlideQcStatus.FINDINGS
    assert [finding.finding_type for finding in result.findings] == [QcFindingType.LINE_BREAK_ISSUE]
    assert result.summary == "Major text wrapping issues compared to the reference"
    assert result.comment_bullets == [
        "The main title is incorrectly split across many short lines.",
        "The subtitle wraps to two lines.",
    ]


def test_openai_qc_evaluator_saves_exact_model_input_images(tmp_path: Path) -> None:
    parsed_payload = SimpleNamespace(
        status="ok",
        summary=None,
        bullets=[],
        note="",
        comparison_confidence=0.9,
        findings=[],
    )
    fake_client = _FakeClient(parsed_payload)
    evaluator = OpenAIQCEvaluator(api_key="test", client=fake_client)

    debug_dir = tmp_path / "model-inputs" / "slide-001"
    evaluator.compare_pages(
        slide_index=0,
        page_index=0,
        reference_page=_page(0),
        candidate_page=_page(0),
        debug_output_dir=debug_dir,
    )

    assert (debug_dir / "01-candidate-slide.png").exists()
    assert (debug_dir / "02-reference-page.png").exists()
    assert (debug_dir / "03-reference-page.png").exists() is False


def test_openai_qc_evaluator_normalizes_both_images_to_same_dimensions() -> None:
    parsed_payload = SimpleNamespace(
        status="ok",
        summary=None,
        bullets=[],
        note="",
        comparison_confidence=0.9,
        findings=[],
    )
    fake_client = _FakeClient(parsed_payload)
    evaluator = OpenAIQCEvaluator(api_key="test", client=fake_client, max_image_dimension=1400)

    candidate_image, reference_image = evaluator._prepare_comparison_images(
        candidate_page=_page_with_size(960, 540, page_index=0),
        reference_page=_page_with_size(1400, 788, page_index=0),
    )

    assert candidate_image.size == reference_image.size
    assert candidate_image.size == (960, 540)
