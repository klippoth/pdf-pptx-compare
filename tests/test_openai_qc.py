from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import numpy as np

from app.services.models import PageImage, ParagraphLayout, QcFindingType, SlideQcStatus, TextBox, TextLayout, TextSource
from app.services.openai_qc import OpenAIQCEvaluator


class _FakeResponses:
    def __init__(self, parsed_payload):
        self.parsed_payload = parsed_payload
        self.last_kwargs = None
        self.calls = []
        self._call_index = 0

    def parse(self, **kwargs):
        self.last_kwargs = kwargs
        self.calls.append(kwargs)
        if isinstance(self.parsed_payload, list):
            index = min(self._call_index, len(self.parsed_payload) - 1)
            payload = self.parsed_payload[index]
            self._call_index += 1
        else:
            payload = self.parsed_payload
        return SimpleNamespace(output_parsed=payload)


class _FakeClient:
    def __init__(self, parsed_payload):
        self.responses = _FakeResponses(parsed_payload)


def _page(page_index: int = 0) -> PageImage:
    image = np.full((240, 320, 3), 255, dtype=np.uint8)
    return PageImage(page_index=page_index, image=image, image_path=Path(f"/tmp/page-{page_index}.png"))


def _page_with_size(width: int, height: int, page_index: int = 0) -> PageImage:
    image = np.full((height, width, 3), 255, dtype=np.uint8)
    return PageImage(page_index=page_index, image=image, image_path=Path(f"/tmp/page-{page_index}-{width}x{height}.png"))


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


def test_openai_qc_evaluator_maps_structured_response_to_slide_qc_result() -> None:
    general_payload = SimpleNamespace(
        status="findings",
        summary="Visual differences found",
        bullets=[
            "Missing logo in the top-right corner",
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
                type="extra_content",
                severity="medium",
                message="Extra shape appears near the footer.",
                confidence=0.81,
                bbox=[0.62, 0.72, 0.94, 0.84],
            ),
        ],
    )
    text_payload = SimpleNamespace(
        status="findings",
        summary="Text discrepancy found",
        bullets=[
            "Headline text differs from the PDF reference",
        ],
        note="",
        comparison_confidence=0.89,
        findings=[
            SimpleNamespace(
                type="wrong_text",
                severity="medium",
                message="Headline text differs from the PDF reference.",
                confidence=0.88,
                bbox=[0.24, 0.18, 0.88, 0.32],
            ),
        ],
    )
    fake_client = _FakeClient([general_payload, text_payload])
    evaluator = OpenAIQCEvaluator(api_key="test", client=fake_client)

    result = evaluator.compare_pages(
        slide_index=2,
        page_index=2,
        reference_page=_page(2),
        candidate_page=_page(2),
        reference_layout=_layout("Federal Signal Q2 2023 Earnings Call", page_number=3),
        candidate_layout=_layout("Federal Signal O2 2023 Earnings Call", page_number=3),
    )

    assert result.status == SlideQcStatus.FINDINGS
    assert result.alignment_confidence == 0.9
    assert result.summary == "AI review findings"
    assert result.comment_bullets == [
        "Missing logo in the top-right corner",
        "Extra shape appears near the footer",
        "Headline text differs from the PDF reference",
    ]
    assert [finding.finding_type for finding in result.findings] == [
        QcFindingType.MISSING_CONTENT,
        QcFindingType.EXTRA_CONTENT,
        QcFindingType.WRONG_TEXT,
    ]
    assert fake_client.responses.calls[0]["model"] == "gpt-5.3-chat-latest"
    assert "temperature" not in fake_client.responses.calls[0]
    assert len(fake_client.responses.calls) == 2
    content = fake_client.responses.calls[0]["input"][1]["content"]
    image_items = [item for item in content if item["type"] == "input_image"]
    assert len(image_items) >= 2


def test_openai_qc_prompt_includes_extracted_candidate_and_reference_text() -> None:
    parsed_payload = [
        SimpleNamespace(
            status="ok",
            summary=None,
            bullets=[],
            note="",
            comparison_confidence=0.9,
            findings=[],
        ),
        SimpleNamespace(
            status="ok",
            summary=None,
            bullets=[],
            note="",
            comparison_confidence=0.9,
            findings=[],
        ),
    ]
    fake_client = _FakeClient(parsed_payload)
    evaluator = OpenAIQCEvaluator(api_key="test", client=fake_client)

    evaluator.compare_pages(
        slide_index=0,
        page_index=0,
        reference_page=_page(0),
        candidate_page=_page(0),
        reference_layout=_layout("Federal Signal Q2 2023 Earnings Call"),
        candidate_layout=_layout("Federal Signal O2 2023 Earnings Call"),
    )

    user_content = fake_client.responses.calls[1]["input"][1]["content"]
    text_blocks = [item["text"] for item in user_content if item["type"] == "input_text"]
    assert any("Candidate PowerPoint text:" in block for block in text_blocks)
    assert any("Reference PDF text:" in block for block in text_blocks)
    assert any("Potential text discrepancies from deterministic diff:" in block for block in text_blocks)
    assert any("Federal Signal O2 2023 Earnings Call" in block for block in text_blocks)
    assert any("Federal Signal Q2 2023 Earnings Call" in block for block in text_blocks)
    assert any('reference="Federal Signal Q2 2023 Earnings Call" | candidate="Federal Signal O2 2023 Earnings Call"' in block for block in text_blocks)
    assert any("Review suspicious text items line by line." in block for block in text_blocks)


def test_openai_qc_prompt_override_is_appended_to_both_model_passes() -> None:
    parsed_payload = [
        SimpleNamespace(status="ok", summary=None, bullets=[], note="", comparison_confidence=0.9, findings=[]),
        SimpleNamespace(status="ok", summary=None, bullets=[], note="", comparison_confidence=0.9, findings=[]),
    ]
    fake_client = _FakeClient(parsed_payload)
    evaluator = OpenAIQCEvaluator(api_key="test", client=fake_client)

    evaluator.compare_pages(
        slide_index=0,
        page_index=0,
        reference_page=_page(0),
        candidate_page=_page(0),
        prompt_override="Focus especially on footnotes, superscripts, and presenter names.",
    )

    general_system_prompt = fake_client.responses.calls[0]["input"][0]["content"]
    general_user_texts = [
        item["text"] for item in fake_client.responses.calls[0]["input"][1]["content"] if item["type"] == "input_text"
    ]
    text_system_prompt = fake_client.responses.calls[1]["input"][0]["content"]
    text_user_texts = [
        item["text"] for item in fake_client.responses.calls[1]["input"][1]["content"] if item["type"] == "input_text"
    ]
    assert "Additional run-specific user QC instructions" in general_system_prompt
    assert any("Focus especially on footnotes, superscripts, and presenter names." in item for item in general_user_texts)
    assert "Additional run-specific user QC instructions" in text_system_prompt
    assert any("Focus especially on footnotes, superscripts, and presenter names." in item for item in text_user_texts)


def test_openai_qc_full_prompt_config_can_be_overridden() -> None:
    parsed_payload = [
        SimpleNamespace(status="ok", summary=None, bullets=[], note="", comparison_confidence=0.9, findings=[]),
        SimpleNamespace(status="ok", summary=None, bullets=[], note="", comparison_confidence=0.9, findings=[]),
    ]
    fake_client = _FakeClient(parsed_payload)
    evaluator = OpenAIQCEvaluator(api_key="test", client=fake_client)

    evaluator.compare_pages(
        slide_index=0,
        page_index=0,
        reference_page=_page(0),
        candidate_page=_page(0),
        prompt_config={
            "general_system_prompt": "Custom visual system prompt",
            "general_user_prompt": "Custom visual user prompt",
            "text_system_prompt": "Custom text system prompt",
            "text_user_prompt": "Custom text user prompt",
        },
    )

    assert fake_client.responses.calls[0]["input"][0]["content"] == "Custom visual system prompt"
    assert fake_client.responses.calls[0]["input"][1]["content"][0]["text"] == "Custom visual user prompt"
    assert fake_client.responses.calls[1]["input"][0]["content"] == "Custom text system prompt"
    assert fake_client.responses.calls[1]["input"][1]["content"][0]["text"] == "Custom text user prompt"


def test_openai_qc_prompt_forbids_grammar_fixing_for_wrong_text() -> None:
    parsed_payload = [
        SimpleNamespace(status="ok", summary=None, bullets=[], note="", comparison_confidence=0.9, findings=[]),
        SimpleNamespace(status="ok", summary=None, bullets=[], note="", comparison_confidence=0.9, findings=[]),
    ]
    fake_client = _FakeClient(parsed_payload)
    evaluator = OpenAIQCEvaluator(api_key="test", client=fake_client)

    evaluator.compare_pages(slide_index=0, page_index=0, reference_page=_page(0), candidate_page=_page(0))

    system_prompt = fake_client.responses.calls[1]["input"][0]["content"]
    user_prompt = fake_client.responses.calls[1]["input"][1]["content"][0]["text"]
    assert "Do not proofread, rewrite, improve grammar, or suggest better phrasing." in system_prompt
    assert "Do not correct grammar" in user_prompt
    assert "normalize wording" in user_prompt
    assert "infer what the text probably meant" in user_prompt


def test_openai_qc_prompt_explicitly_checks_obvious_typos_and_misspellings() -> None:
    parsed_payload = [
        SimpleNamespace(status="ok", summary=None, bullets=[], note="", comparison_confidence=0.9, findings=[]),
        SimpleNamespace(status="ok", summary=None, bullets=[], note="", comparison_confidence=0.9, findings=[]),
    ]
    fake_client = _FakeClient(parsed_payload)
    evaluator = OpenAIQCEvaluator(api_key="test", client=fake_client)

    evaluator.compare_pages(slide_index=0, page_index=0, reference_page=_page(0), candidate_page=_page(0))

    system_prompt = fake_client.responses.calls[1]["input"][0]["content"]
    user_prompt = fake_client.responses.calls[1]["input"][1]["content"][0]["text"]
    assert "Obvious readable typos and misspellings should be treated as high-signal wrong_text findings" in system_prompt
    assert "This includes obvious typos" in user_prompt
    assert "Spell-check visible text carefully against the reference" in system_prompt
    assert "dropped letters, extra letters, swapped letters, and repeated letters" in user_prompt
    assert "review the slide in ordered passes" in system_prompt
    assert "check suspicious lines one by one" in system_prompt
    assert "verify every proposed finding a second time" in system_prompt
    assert "deterministic text-diff checklist" in system_prompt
    assert "Check these areas in order" in user_prompt
    assert "Work through those suspicious text items line by line" in user_prompt
    assert "If a region looks the same in both images on the second check, do not report a discrepancy there." in user_prompt
    assert "If a deterministic text-diff checklist is provided" in user_prompt


def test_openai_qc_evaluator_exposes_unavailability_without_key() -> None:
    evaluator = OpenAIQCEvaluator(api_key=None)
    assert evaluator.is_available() is False


def test_openai_qc_evaluator_keeps_obvious_line_break_findings_and_comments() -> None:
    parsed_payload = [
        SimpleNamespace(
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
        ),
        SimpleNamespace(
            status="ok",
            summary=None,
            bullets=[],
            note="",
            comparison_confidence=0.82,
            findings=[],
        ),
    ]
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


def test_openai_qc_evaluator_drops_chart_size_position_findings() -> None:
    parsed_payload = [
        SimpleNamespace(
            status="findings",
            summary="Chart discrepancy found",
            bullets=[
                "The bar chart is clearly shifted too far right and appears narrower than the reference.",
            ],
            note="",
            comparison_confidence=0.86,
            findings=[
                SimpleNamespace(
                    type="size_position_issue",
                    severity="medium",
                    message="Bar chart is clearly the wrong size and position compared to the reference.",
                    confidence=0.87,
                    bbox=[0.18, 0.22, 0.84, 0.78],
                ),
            ],
        ),
        SimpleNamespace(
            status="ok",
            summary=None,
            bullets=[],
            note="",
            comparison_confidence=0.86,
            findings=[],
        ),
    ]
    fake_client = _FakeClient(parsed_payload)
    evaluator = OpenAIQCEvaluator(api_key="test", client=fake_client)

    result = evaluator.compare_pages(slide_index=1, page_index=1, reference_page=_page(1), candidate_page=_page(1))

    assert result.status == SlideQcStatus.OK
    assert result.findings == []
    assert result.summary is None
    assert result.comment_bullets == []


def test_openai_qc_evaluator_saves_exact_model_input_images(tmp_path: Path) -> None:
    parsed_payload = [
        SimpleNamespace(status="ok", summary=None, bullets=[], note="", comparison_confidence=0.9, findings=[]),
        SimpleNamespace(status="ok", summary=None, bullets=[], note="", comparison_confidence=0.9, findings=[]),
    ]
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
    assert (debug_dir / "00-upload-metadata.json").exists()
    assert (debug_dir / "03-reference-page.png").exists() is False
    assert list(debug_dir.glob("10-hotspot-*.png")) == []
    assert list(debug_dir.glob("11-hotspot-*.png")) == []
    metadata = (debug_dir / "00-upload-metadata.json").read_text(encoding="utf-8")
    assert '"note": "These PNG files are the exact bytes uploaded to OpenAI by this app.' in metadata
    assert '"prompt_override": ""' in metadata


def test_openai_qc_general_pass_receives_only_full_slide_images() -> None:
    general_payload = SimpleNamespace(status="ok", summary=None, bullets=[], note="", comparison_confidence=0.9, findings=[])
    text_payload = SimpleNamespace(status="ok", summary=None, bullets=[], note="", comparison_confidence=0.9, findings=[])
    fake_client = _FakeClient([general_payload, text_payload])
    evaluator = OpenAIQCEvaluator(api_key="test", client=fake_client)

    evaluator.compare_pages(
        slide_index=0,
        page_index=0,
        candidate_page=_page(0),
        reference_page=_page(0),
    )

    first_call_content = fake_client.responses.calls[0]["input"][1]["content"]
    image_items = [item for item in first_call_content if item["type"] == "input_image"]
    text_items = [item["text"] for item in first_call_content if item["type"] == "input_text"]
    assert len(image_items) == 2
    assert not any("Hotspot" in item for item in text_items)


def test_openai_qc_text_pass_receives_line_level_discrepancy_crops() -> None:
    general_payload = SimpleNamespace(status="ok", summary=None, bullets=[], note="", comparison_confidence=0.9, findings=[])
    text_payload = SimpleNamespace(status="ok", summary=None, bullets=[], note="", comparison_confidence=0.9, findings=[])
    fake_client = _FakeClient([general_payload, text_payload])
    evaluator = OpenAIQCEvaluator(api_key="test", client=fake_client)

    evaluator.compare_pages(
        slide_index=0,
        page_index=0,
        reference_page=_page(0),
        candidate_page=_page(0),
        reference_layout=_layout("Reported results include 10K."),
        candidate_layout=_layout("Reported results include 10R."),
    )

    second_call_content = fake_client.responses.calls[1]["input"][1]["content"]
    image_items = [item for item in second_call_content if item["type"] == "input_image"]
    text_items = [item["text"] for item in second_call_content if item["type"] == "input_text"]
    assert len(image_items) > 2
    assert any("Suspicious text diff 1" in item for item in text_items)
    assert any('Reference reads "Reported results include 10K."' in item for item in text_items)
    assert any('Candidate reads "Reported results include 10R."' in item for item in text_items)


def test_openai_qc_filters_out_color_only_findings_and_comments() -> None:
    general_payload = SimpleNamespace(
        status="findings",
        summary="Color differences found",
        bullets=["Accent color differs from the reference."],
        note="",
        comparison_confidence=0.9,
        findings=[
            SimpleNamespace(
                type="wrong_color",
                severity="medium",
                message="Accent color differs from the reference.",
                confidence=0.8,
                bbox=[0.1, 0.1, 0.3, 0.2],
            ),
        ],
    )
    text_payload = SimpleNamespace(status="ok", summary=None, bullets=[], note="", comparison_confidence=0.9, findings=[])
    fake_client = _FakeClient([general_payload, text_payload])
    evaluator = OpenAIQCEvaluator(api_key="test", client=fake_client)

    result = evaluator.compare_pages(slide_index=0, page_index=0, reference_page=_page(0), candidate_page=_page(0))

    assert result.status == SlideQcStatus.OK
    assert result.findings == []
    assert result.summary is None
    assert result.comment_bullets == []


def test_openai_qc_keeps_font_color_findings_when_they_overlap_text() -> None:
    general_payload = SimpleNamespace(
        status="findings",
        summary="Font color mismatch found",
        bullets=["Title font color differs from the reference."],
        note="",
        comparison_confidence=0.9,
        findings=[
            SimpleNamespace(
                type="wrong_color",
                severity="medium",
                message="Title font color differs from the reference.",
                confidence=0.93,
                bbox=[0.12, 0.1, 0.62, 0.22],
            ),
        ],
    )
    text_payload = SimpleNamespace(status="ok", summary=None, bullets=[], note="", comparison_confidence=0.9, findings=[])
    fake_client = _FakeClient([general_payload, text_payload])
    evaluator = OpenAIQCEvaluator(api_key="test", client=fake_client)

    result = evaluator.compare_pages(
        slide_index=0,
        page_index=0,
        reference_page=_page(0),
        candidate_page=_page(0),
        candidate_layout=_layout("Federal Signal Q2 2023 Earnings Call", bbox=(0.1, 0.1, 0.65, 0.24)),
    )

    assert result.status == SlideQcStatus.FINDINGS
    assert [finding.finding_type for finding in result.findings] == [QcFindingType.WRONG_COLOR]
    assert result.comment_bullets == ["Title font color differs from the reference."]


def test_openai_qc_drops_low_confidence_font_color_findings() -> None:
    general_payload = SimpleNamespace(
        status="findings",
        summary="Font color mismatch found",
        bullets=["Title font color differs from the reference."],
        note="",
        comparison_confidence=0.9,
        findings=[
            SimpleNamespace(
                type="wrong_color",
                severity="medium",
                message="Title font color differs from the reference.",
                confidence=0.86,
                bbox=[0.12, 0.1, 0.62, 0.22],
            ),
        ],
    )
    text_payload = SimpleNamespace(status="ok", summary=None, bullets=[], note="", comparison_confidence=0.9, findings=[])
    fake_client = _FakeClient([general_payload, text_payload])
    evaluator = OpenAIQCEvaluator(api_key="test", client=fake_client)

    result = evaluator.compare_pages(
        slide_index=0,
        page_index=0,
        reference_page=_page(0),
        candidate_page=_page(0),
        candidate_layout=_layout("Federal Signal Q2 2023 Earnings Call", bbox=(0.1, 0.1, 0.65, 0.24)),
    )

    assert result.status == SlideQcStatus.OK
    assert result.findings == []
    assert result.summary is None
    assert result.comment_bullets == []


def test_openai_qc_drops_text_findings_when_deterministic_text_matches() -> None:
    general_payload = SimpleNamespace(status="ok", summary=None, bullets=[], note="", comparison_confidence=0.9, findings=[])
    text_payload = SimpleNamespace(
        status="findings",
        summary="Text discrepancy found",
        bullets=["Presenter name spelling differs from the reference."],
        note="",
        comparison_confidence=0.9,
        findings=[
            SimpleNamespace(
                type="wrong_text",
                severity="medium",
                message="Presenter name spelling differs from the reference.",
                confidence=0.88,
                bbox=[0.55, 0.78, 0.92, 0.9],
            ),
        ],
    )
    fake_client = _FakeClient([general_payload, text_payload])
    evaluator = OpenAIQCEvaluator(api_key="test", client=fake_client)

    result = evaluator.compare_pages(
        slide_index=0,
        page_index=0,
        reference_page=_page(0),
        candidate_page=_page(0),
        reference_layout=_layout("Jennifer Sherman"),
        candidate_layout=_layout("Jennifer Sherman"),
    )

    assert result.status == SlideQcStatus.OK
    assert result.findings == []
    assert result.summary is None
    assert result.comment_bullets == []


def test_openai_qc_evaluator_normalizes_both_images_to_same_dimensions() -> None:
    parsed_payload = [
        SimpleNamespace(status="ok", summary=None, bullets=[], note="", comparison_confidence=0.9, findings=[]),
        SimpleNamespace(status="ok", summary=None, bullets=[], note="", comparison_confidence=0.9, findings=[]),
    ]
    fake_client = _FakeClient(parsed_payload)
    evaluator = OpenAIQCEvaluator(api_key="test", client=fake_client, max_image_dimension=1400)

    candidate_image, reference_image, _ = evaluator._prepare_comparison_images(
        candidate_page=_page_with_size(960, 540, page_index=0),
        reference_page=_page_with_size(1400, 788, page_index=0),
    )

    assert candidate_image.size == reference_image.size
    assert candidate_image.size == (960, 540)


def test_openai_qc_evaluator_preserves_original_candidate_resolution_when_uncapped() -> None:
    parsed_payload = [
        SimpleNamespace(status="ok", summary=None, bullets=[], note="", comparison_confidence=0.9, findings=[]),
        SimpleNamespace(status="ok", summary=None, bullets=[], note="", comparison_confidence=0.9, findings=[]),
    ]
    fake_client = _FakeClient(parsed_payload)
    evaluator = OpenAIQCEvaluator(api_key="test", client=fake_client, max_image_dimension=0)

    candidate_image, reference_image, _ = evaluator._prepare_comparison_images(
        candidate_page=_page_with_size(3000, 2318, page_index=0),
        reference_page=_page_with_size(3200, 2472, page_index=0),
    )

    assert candidate_image.size == (3000, 2318)
    assert reference_image.size == (3000, 2318)
