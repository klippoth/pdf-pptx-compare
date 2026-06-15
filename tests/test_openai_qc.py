from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import numpy as np

from app.services.models import PageImage, ParagraphLayout, QcFindingType, SlideQcStatus, TextBox, TextLayout, TextSource
from app.services.openai_qc import OpenAIQCEvaluator


class _FakeResponses:
    def __init__(self, parsed_payload):
        self.parsed_payload = parsed_payload
        self.calls: list[dict] = []

    def parse(self, **kwargs):
        self.calls.append(kwargs)
        return SimpleNamespace(output_parsed=self.parsed_payload)


class _FakeClient:
    def __init__(self, parsed_payload):
        self.responses = _FakeResponses(parsed_payload)


def _page(page_index: int = 0, *, width: int = 320, height: int = 240) -> PageImage:
    image = np.full((height, width, 3), 255, dtype=np.uint8)
    return PageImage(page_index=page_index, image=image, image_path=Path(f"/tmp/page-{page_index}.png"))


def _layout(text: str, bbox=(0.1, 0.1, 0.6, 0.2), page_number: int = 1) -> TextLayout:
    line = TextBox(
        text=text,
        bbox=bbox,
        page_number=page_number,
        confidence=1.0,
        source=TextSource.NATIVE,
    )
    paragraph = ParagraphLayout(
        text=text,
        lines=[line],
        bbox=bbox,
        page_number=page_number,
        confidence=1.0,
        source=TextSource.NATIVE,
    )
    return TextLayout(
        page_number=page_number,
        page_size=(1000.0, 700.0),
        paragraphs=[paragraph],
        lines=[line],
        source=TextSource.NATIVE,
        total_characters=len(text),
        average_confidence=1.0,
        extracted_with_ocr=False,
    )


def test_openai_qc_single_visual_pass_maps_structured_response_to_slide_qc_result() -> None:
    payload = SimpleNamespace(
        status="findings",
        summary="Visible discrepancies found",
        bullets=[
            "Contribution margin label is missing.",
            'Chart label includes an extra "$" before 3%.',
        ],
        note="",
        comparison_confidence=0.91,
        findings=[
            SimpleNamespace(
                type="missing_content",
                severity="high",
                message="Contribution margin label is missing.",
                confidence=0.97,
                bbox=[0.78, 0.12, 0.92, 0.24],
            ),
            SimpleNamespace(
                type="wrong_text",
                severity="medium",
                message='Chart label includes an extra "$" before 3%.',
                confidence=0.92,
                bbox=[0.28, 0.56, 0.39, 0.68],
            ),
        ],
    )
    fake_client = _FakeClient(payload)
    evaluator = OpenAIQCEvaluator(api_key="test", client=fake_client)

    result = evaluator.compare_pages(
        slide_index=2,
        page_index=2,
        reference_page=_page(2),
        candidate_page=_page(2),
        candidate_layout=_layout("medSR Revenue & Contribution Margin"),
    )

    assert result.status == SlideQcStatus.FINDINGS
    assert result.alignment_confidence == 0.91
    assert result.summary == "Visible discrepancies found"
    assert result.comment_bullets == [
        "Contribution margin label is missing.",
        'Chart label includes an extra "$" before 3%.',
    ]
    assert [finding.finding_type for finding in result.findings] == [
        QcFindingType.MISSING_CONTENT,
        QcFindingType.WRONG_TEXT,
    ]
    assert len(fake_client.responses.calls) == 1


def test_openai_qc_prompt_includes_panel_and_pptx_text_only() -> None:
    payload = SimpleNamespace(
        status="ok",
        summary=None,
        bullets=[],
        note="",
        comparison_confidence=0.9,
        findings=[],
    )
    fake_client = _FakeClient(payload)
    evaluator = OpenAIQCEvaluator(api_key="test", client=fake_client)

    evaluator.compare_pages(
        slide_index=0,
        page_index=0,
        reference_page=_page(0),
        candidate_page=_page(0),
        candidate_layout=_layout("Federal Signal Q2 2023 Earnings Call"),
    )

    call = fake_client.responses.calls[0]
    system_prompt = call["input"][0]["content"]
    user_content = call["input"][1]["content"]
    text_blocks = [item["text"] for item in user_content if item["type"] == "input_text"]
    image_items = [item for item in user_content if item["type"] == "input_image"]

    assert len(image_items) == 3
    assert "Candidate PowerPoint text:" in text_blocks[1]
    assert "Reference PDF text:" not in text_blocks[1]
    assert "Do not rely on extracted PDF text." in system_prompt


def test_openai_qc_prompt_override_and_prompt_config_apply_to_single_pass() -> None:
    payload = SimpleNamespace(
        status="ok",
        summary=None,
        bullets=[],
        note="",
        comparison_confidence=0.9,
        findings=[],
    )
    fake_client = _FakeClient(payload)
    evaluator = OpenAIQCEvaluator(api_key="test", client=fake_client)

    evaluator.compare_pages(
        slide_index=0,
        page_index=0,
        reference_page=_page(0),
        candidate_page=_page(0),
        prompt_override="Focus especially on footnotes and presenter names.",
        prompt_config={
            "general_system_prompt": "Custom visual system",
            "general_user_prompt": "Custom visual user",
            "text_system_prompt": "Custom text support system",
            "text_user_prompt": "Custom text support user",
        },
    )

    call = fake_client.responses.calls[0]
    system_prompt = call["input"][0]["content"]
    first_user_text = call["input"][1]["content"][0]["text"]
    override_texts = [item["text"] for item in call["input"][1]["content"] if item["type"] == "input_text"]

    assert "Custom visual system" in system_prompt
    assert "Custom text support system" in system_prompt
    assert "Custom visual user" in first_user_text
    assert "Custom text support user" in first_user_text
    assert any("Focus especially on footnotes and presenter names." in item for item in override_texts)


def test_openai_qc_saves_panel_and_original_images_for_debug(tmp_path: Path) -> None:
    payload = SimpleNamespace(
        status="ok",
        summary=None,
        bullets=[],
        note="",
        comparison_confidence=0.9,
        findings=[],
    )
    fake_client = _FakeClient(payload)
    evaluator = OpenAIQCEvaluator(api_key="test", client=fake_client)
    debug_dir = tmp_path / "model-inputs" / "slide-001"

    evaluator.compare_pages(
        slide_index=0,
        page_index=0,
        reference_page=_page(0, width=360, height=280),
        candidate_page=_page(0, width=320, height=240),
        debug_output_dir=debug_dir,
    )

    assert (debug_dir / "01-comparison-panel.png").exists()
    assert (debug_dir / "02-candidate-slide.png").exists()
    assert (debug_dir / "03-reference-page.png").exists()
    metadata = (debug_dir / "00-upload-metadata.json").read_text(encoding="utf-8")
    assert '"left": "PPTX candidate"' in metadata
    assert '"right": "PDF reference"' in metadata


def test_openai_qc_suppresses_outer_edge_missing_and_extra_findings() -> None:
    payload = SimpleNamespace(
        status="findings",
        summary="Border discrepancies found",
        bullets=[
            "Missing white strip near the left border.",
            "Extra artifact appears along the bottom edge.",
        ],
        note="",
        comparison_confidence=0.86,
        findings=[
            SimpleNamespace(
                type="missing_content",
                severity="medium",
                message="Missing white strip near the left border.",
                confidence=0.88,
                bbox=[0.0, 0.30, 0.03, 0.70],
            ),
            SimpleNamespace(
                type="extra_content",
                severity="medium",
                message="Extra artifact appears along the bottom edge.",
                confidence=0.85,
                bbox=[0.25, 0.97, 0.75, 1.0],
            ),
        ],
    )
    fake_client = _FakeClient(payload)
    evaluator = OpenAIQCEvaluator(api_key="test", client=fake_client)

    result = evaluator.compare_pages(
        slide_index=0,
        page_index=0,
        reference_page=_page(0),
        candidate_page=_page(0),
    )

    assert result.status == SlideQcStatus.OK
    assert result.findings == []
    assert result.comment_bullets == []


def test_openai_qc_drops_line_break_and_position_findings() -> None:
    payload = SimpleNamespace(
        status="findings",
        summary="Layout differences found",
        bullets=[
            "Title wraps across lines.",
            "Chart is slightly shifted.",
        ],
        note="",
        comparison_confidence=0.83,
        findings=[
            SimpleNamespace(
                type="line_break_issue",
                severity="medium",
                message="Title wraps across lines.",
                confidence=0.9,
                bbox=[0.12, 0.08, 0.55, 0.21],
            ),
            SimpleNamespace(
                type="size_position_issue",
                severity="medium",
                message="Chart is slightly shifted.",
                confidence=0.89,
                bbox=[0.18, 0.22, 0.82, 0.79],
            ),
        ],
    )
    fake_client = _FakeClient(payload)
    evaluator = OpenAIQCEvaluator(api_key="test", client=fake_client)

    result = evaluator.compare_pages(
        slide_index=0,
        page_index=0,
        reference_page=_page(0),
        candidate_page=_page(0),
    )

    assert result.status == SlideQcStatus.OK
    assert result.findings == []


def test_openai_qc_keeps_only_high_confidence_text_color_findings_on_text() -> None:
    payload = SimpleNamespace(
        status="findings",
        summary="Text color issue found",
        bullets=["Title font color differs from the reference."],
        note="",
        comparison_confidence=0.9,
        findings=[
            SimpleNamespace(
                type="wrong_color",
                severity="medium",
                message="Title font color differs from the reference.",
                confidence=0.93,
                bbox=[0.12, 0.10, 0.58, 0.23],
            ),
        ],
    )
    fake_client = _FakeClient(payload)
    evaluator = OpenAIQCEvaluator(api_key="test", client=fake_client)

    result = evaluator.compare_pages(
        slide_index=0,
        page_index=0,
        reference_page=_page(0),
        candidate_page=_page(0),
        candidate_layout=_layout("Federal Signal Q2 2023 Earnings Call", bbox=(0.1, 0.1, 0.62, 0.24)),
    )

    assert result.status == SlideQcStatus.FINDINGS
    assert [finding.finding_type for finding in result.findings] == [QcFindingType.WRONG_COLOR]


def test_openai_qc_drops_non_text_or_low_confidence_color_findings() -> None:
    payload = SimpleNamespace(
        status="findings",
        summary="Color issue found",
        bullets=["Accent blue differs from the reference."],
        note="",
        comparison_confidence=0.9,
        findings=[
            SimpleNamespace(
                type="wrong_color",
                severity="medium",
                message="Accent blue differs from the reference.",
                confidence=0.86,
                bbox=[0.12, 0.10, 0.58, 0.23],
            ),
        ],
    )
    fake_client = _FakeClient(payload)
    evaluator = OpenAIQCEvaluator(api_key="test", client=fake_client)

    result = evaluator.compare_pages(
        slide_index=0,
        page_index=0,
        reference_page=_page(0),
        candidate_page=_page(0),
        candidate_layout=_layout("Federal Signal Q2 2023 Earnings Call", bbox=(0.70, 0.70, 0.90, 0.82)),
    )

    assert result.status == SlideQcStatus.OK
    assert result.findings == []


def test_openai_qc_preserves_original_sizes_when_uncapped_and_builds_panel() -> None:
    payload = SimpleNamespace(
        status="ok",
        summary=None,
        bullets=[],
        note="",
        comparison_confidence=0.9,
        findings=[],
    )
    evaluator = OpenAIQCEvaluator(api_key="test", client=_FakeClient(payload), max_image_dimension=0)

    candidate_image, reference_image = evaluator._prepare_comparison_images(
        candidate_page=_page(0, width=3000, height=2318),
        reference_page=_page(0, width=3200, height=2472),
    )
    panel = evaluator._build_side_by_side_comparison(
        candidate_image=candidate_image,
        reference_image=reference_image,
    )

    assert candidate_image.size == (3000, 2318)
    assert reference_image.size == (3200, 2472)
    assert panel.width > candidate_image.width + reference_image.width
    assert panel.height > max(candidate_image.height, reference_image.height)
