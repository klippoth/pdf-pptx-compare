from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import fitz
from PIL import Image

from app.services.rasterizer import Rasterizer


def test_rasterizer_uses_poppler_when_available(monkeypatch, tmp_path: Path) -> None:
    rasterizer = Rasterizer(dpi=144)
    monkeypatch.setattr(
        rasterizer,
        "_resolve_poppler_binary",
        lambda name: Path("/usr/bin/pdftocairo") if name == "pdftocairo" else None,
    )

    def fake_run(command, capture_output, text, timeout):
        assert str(command[0]).endswith("pdftocairo")
        assert "-png" in command
        assert "-r" in command
        Image.new("RGB", (64, 32), (255, 255, 255)).save(tmp_path / "renders" / "page-1.png")
        return SimpleNamespace(returncode=0, stdout="", stderr="")

    monkeypatch.setattr("app.services.rasterizer.subprocess.run", fake_run)

    pages = rasterizer.render_pdf(tmp_path / "candidate.pdf", tmp_path / "renders", engine="auto")

    assert rasterizer.last_used_engine == "poppler"
    assert len(pages) == 1
    assert pages[0].image_path.name == "page-1.png"


def test_rasterizer_falls_back_to_fitz_when_poppler_is_unavailable(monkeypatch, tmp_path: Path) -> None:
    pdf_path = tmp_path / "sample.pdf"
    document = fitz.open()
    page = document.new_page(width=160, height=90)
    page.insert_text((20, 40), "QC")
    document.save(pdf_path)
    document.close()

    rasterizer = Rasterizer(dpi=72)
    monkeypatch.setattr(rasterizer, "_resolve_poppler_binary", lambda name: None)
    pages = rasterizer.render_pdf(pdf_path, tmp_path / "renders", engine="auto")

    assert rasterizer.last_used_engine == "fitz"
    assert len(pages) == 1
    assert pages[0].image_path.exists()
