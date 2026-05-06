from __future__ import annotations

from pathlib import Path

import fitz

from app.services.models import PageImage, TextLayout, TextSource
from app.services.rasterizer import Rasterizer
from app.services.text_layout_extractor import TextLayoutExtractor


class FakeOcrProvider:
    def is_available(self) -> bool:
        return True

    def extract_page(self, page_image: PageImage) -> TextLayout:
        return TextLayout(
            page_number=page_image.page_index + 1,
            page_size=(float(page_image.width), float(page_image.height)),
            source=TextSource.OCR,
            total_characters=18,
            average_confidence=0.92,
            extracted_with_ocr=True,
        )


def test_text_layout_extractor_uses_native_pdf_text_when_available(tmp_path: Path) -> None:
    pdf_path = tmp_path / "native.pdf"
    document = fitz.open()
    page = document.new_page(width=400, height=220)
    page.insert_text((72, 72), "Alpha beta gamma")
    page.insert_text((72, 98), "Delta epsilon")
    document.save(pdf_path)
    document.close()

    page_images = Rasterizer(dpi=72).render_pdf(pdf_path, tmp_path / "renders")
    extractor = TextLayoutExtractor(sparse_text_threshold=5)

    layouts = extractor.extract_document(pdf_path, page_images)

    assert len(layouts) == 1
    assert layouts[0].source == TextSource.NATIVE
    assert layouts[0].total_characters > 10
    assert layouts[0].paragraphs
    assert layouts[0].lines


def test_text_layout_extractor_falls_back_to_ocr_when_native_text_is_sparse(tmp_path: Path) -> None:
    pdf_path = tmp_path / "image-only.pdf"
    document = fitz.open()
    document.new_page(width=400, height=220)
    document.save(pdf_path)
    document.close()

    page_images = Rasterizer(dpi=72).render_pdf(pdf_path, tmp_path / "renders")
    extractor = TextLayoutExtractor(ocr_provider=FakeOcrProvider(), sparse_text_threshold=5)

    layouts = extractor.extract_document(pdf_path, page_images)

    assert len(layouts) == 1
    assert layouts[0].source == TextSource.OCR
    assert layouts[0].extracted_with_ocr is True
    assert layouts[0].total_characters == 18
