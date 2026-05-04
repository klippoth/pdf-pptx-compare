from __future__ import annotations

from pathlib import Path

from PIL import Image
from pptx import Presentation
from zipfile import ZipFile

from app.services.deck_writer import DeckWriter
from app.services.models import PagePlacementResult, PlacementBundle, PlacementStatus


def _make_png(path: Path, size: tuple[int, int] = (1200, 700), color: tuple[int, int, int] = (255, 255, 255)) -> Path:
    image = Image.new("RGB", size, color)
    image.save(path)
    return path


def test_build_output_preserves_slides_and_appends_extra_pdf_page(tmp_path: Path) -> None:
    source_pptx = tmp_path / "source.pptx"
    presentation = Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    slide.shapes.add_textbox(0, 0, presentation.slide_width, 600000).text_frame.text = "Original slide"
    presentation.save(source_pptx)

    background_image = _make_png(tmp_path / "background.png", size=(1600, 900), color=(250, 248, 244))
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
            )
        ],
        extra_reference_results=[
            PagePlacementResult(
                candidate_slide_index=None,
                reference_page_index=1,
                status=PlacementStatus.EXTRA_PDF_PAGE,
                background_image_path=extra_image,
                message="Appended unmatched PDF page as a new reference slide",
            )
        ],
    )

    output_pptx = writer.build_output(source_pptx, bundle, tmp_path / "output.pptx")

    updated = Presentation(output_pptx)
    assert updated.slide_width == presentation.slide_width
    assert len(updated.slides) == 2
    assert len(updated.slides[0].shapes) == 2
    parked_picture = updated.slides[0].shapes[-1]
    assert parked_picture.left < updated.slide_width
    assert parked_picture.left + parked_picture.width > updated.slide_width
    assert parked_picture.name == "Sticky"

    assert len(updated.slides[1].shapes) == 1
    full_slide_picture = updated.slides[1].shapes[-1]
    assert full_slide_picture.left == 0
    assert full_slide_picture.top == 0
    assert full_slide_picture.width == updated.slide_width
    assert full_slide_picture.height == updated.slide_height
    assert full_slide_picture.name == "Sticky"

    with ZipFile(output_pptx) as archive:
        slide_xml = archive.read("ppt/slides/slide1.xml").decode("utf-8")
        extra_slide_xml = archive.read("ppt/slides/slide2.xml").decode("utf-8")
    assert "alphaModFix" not in slide_xml
    assert 'name="Sticky"' in slide_xml
    assert 'name="Sticky"' in extra_slide_xml
