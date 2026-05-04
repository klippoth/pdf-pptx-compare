from __future__ import annotations

from pathlib import Path

from PIL import Image
from pptx import Presentation

from app.services.deck_writer import DeckWriter
from app.services.models import PagePlacementResult, PlacementBundle, PlacementStatus


def _make_png(path: Path, size: tuple[int, int] = (1200, 700), color: tuple[int, int, int] = (255, 255, 255)) -> Path:
    image = Image.new("RGB", size, color)
    image.save(path)
    return path


def test_build_output_preserves_slides_inserts_pdf_reference_slides_and_appends_unmatched_pdf_pages(tmp_path: Path) -> None:
    source_pptx = tmp_path / "source.pptx"
    presentation = Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    slide.shapes.add_textbox(0, 0, presentation.slide_width, 600000).text_frame.text = "Original slide 1"
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    slide.shapes.add_textbox(0, 0, presentation.slide_width, 600000).text_frame.text = "Original slide 2"
    presentation.save(source_pptx)

    background_image = _make_png(tmp_path / "background-1.png", size=(1600, 900), color=(250, 248, 244))
    second_background_image = _make_png(tmp_path / "background-2.png", size=(1600, 900), color=(245, 242, 236))
    extra_image = _make_png(tmp_path / "extra.png", size=(1600, 900), color=(230, 230, 230))
    writer = DeckWriter()
    bundle = PlacementBundle(
        slide_results=[
            PagePlacementResult(
                candidate_slide_index=0,
                reference_page_index=0,
                status=PlacementStatus.PLACED,
                background_image_path=background_image,
                message="Parked PDF page off the top-right of the slide as a movable reference image",
            ),
            PagePlacementResult(
                candidate_slide_index=1,
                reference_page_index=1,
                status=PlacementStatus.PLACED,
                background_image_path=second_background_image,
                message="Parked PDF page off the top-right of the slide as a movable reference image",
            )
        ],
        extra_reference_results=[
            PagePlacementResult(
                candidate_slide_index=None,
                reference_page_index=2,
                status=PlacementStatus.EXTRA_PDF_PAGE,
                background_image_path=extra_image,
                message="Appended unmatched PDF page as a new reference slide",
            )
        ],
    )

    output_pptx = writer.build_output(source_pptx, bundle, tmp_path / "output.pptx")

    updated = Presentation(output_pptx)
    assert updated.slide_width == presentation.slide_width
    assert len(updated.slides) == 5
    assert len(updated.slides[0].shapes) == 2
    parked_picture = updated.slides[0].shapes[-1]
    assert parked_picture.left < updated.slide_width
    assert parked_picture.left + parked_picture.width > updated.slide_width
    assert parked_picture.name == "PDF_ORIGINAL"
    assert len(updated.slides[1].shapes) == 1
    first_reference_slide = updated.slides[1].shapes[0]
    assert first_reference_slide.left == 0
    assert first_reference_slide.top == 0
    assert first_reference_slide.width == updated.slide_width
    assert first_reference_slide.height == updated.slide_height
    assert first_reference_slide.name == "PDF_ORIGINAL"
    assert updated.slides[1]._element.cSld.get("name") == "PDF_ORIGINAL"

    assert len(updated.slides[2].shapes) == 2
    second_parked_picture = updated.slides[2].shapes[-1]
    assert second_parked_picture.left < updated.slide_width
    assert second_parked_picture.left + second_parked_picture.width > updated.slide_width
    assert second_parked_picture.name == "PDF_ORIGINAL"

    assert len(updated.slides[3].shapes) == 1
    second_reference_slide = updated.slides[3].shapes[0]
    assert second_reference_slide.left == 0
    assert second_reference_slide.top == 0
    assert second_reference_slide.width == updated.slide_width
    assert second_reference_slide.height == updated.slide_height
    assert second_reference_slide.name == "PDF_ORIGINAL"
    assert updated.slides[3]._element.cSld.get("name") == "PDF_ORIGINAL"

    assert len(updated.slides[4].shapes) == 1
    full_slide_picture = updated.slides[4].shapes[0]
    assert full_slide_picture.left == 0
    assert full_slide_picture.top == 0
    assert full_slide_picture.width == updated.slide_width
    assert full_slide_picture.height == updated.slide_height
    assert full_slide_picture.name == "PDF_ORIGINAL"
    assert updated.slides[4]._element.cSld.get("name") == "PDF_ORIGINAL"
