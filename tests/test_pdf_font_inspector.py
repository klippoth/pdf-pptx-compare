from __future__ import annotations

from pathlib import Path

import fitz

from app.services.pdf_font_inspector import PDFFontInspector


def test_pdf_font_inspector_reports_detected_fonts_and_embedding_status(tmp_path: Path) -> None:
    pdf_path = tmp_path / "font-sample.pdf"
    document = fitz.open()
    page = document.new_page()
    page.insert_text((72, 72), "Hello world")
    document.save(str(pdf_path))
    document.close()

    inspector = PDFFontInspector()
    result = inspector.inspect(pdf_path)
    report_path = inspector.write_report(result, tmp_path / "pdf_fonts.json")

    assert result.page_count == 1
    assert result.page_character_totals == {1: 10}
    assert result.fonts
    assert any(font.name == "Helvetica" for font in result.fonts)
    assert any(font.embedded is False for font in result.fonts)
    helvetica = next(font for font in result.fonts if font.name == "Helvetica")
    assert helvetica.page_character_counts == {1: 10}
    assert report_path.exists()
