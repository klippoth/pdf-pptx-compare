from __future__ import annotations

from pathlib import Path

import pytest

from app.services.renderer import LibreOfficeRenderer, PowerPointRenderer, Renderer, RendererError


class FailingRenderer:
    def export_pptx_to_pdf(self, pptx_path: Path, output_path: Path) -> Path:
        raise RendererError("primary failed")


class SuccessfulRenderer:
    def export_pptx_to_pdf(self, pptx_path: Path, output_path: Path) -> Path:
        output_path.write_bytes(b"%PDF-1.4")
        return output_path


class RecordingRenderer:
    def __init__(self, should_fail: bool = False):
        self.called = False
        self.should_fail = should_fail

    def export_pptx_to_pdf(self, pptx_path: Path, output_path: Path) -> Path:
        self.called = True
        if self.should_fail:
            raise RendererError("renderer failed")
        output_path.write_bytes(b"%PDF-1.4")
        return output_path


def test_renderer_falls_back_to_secondary(tmp_path: Path) -> None:
    renderer = Renderer(
        powerpoint=FailingRenderer(),
        libreoffice=SuccessfulRenderer(),
        enable_powerpoint_fallback=True,
    )
    output_path = tmp_path / "candidate.pdf"

    exported = renderer.export_pptx_to_pdf(tmp_path / "deck.pptx", output_path)

    assert exported == output_path
    assert renderer.last_used_renderer == "libreoffice"


def test_renderer_does_not_call_powerpoint_when_fallback_disabled(tmp_path: Path) -> None:
    powerpoint = RecordingRenderer()
    libreoffice = RecordingRenderer()
    renderer = Renderer(
        powerpoint=powerpoint,
        libreoffice=libreoffice,
        enable_powerpoint_fallback=False,
    )

    exported = renderer.export_pptx_to_pdf(tmp_path / "deck.pptx", tmp_path / "candidate.pdf")

    assert exported.exists()
    assert libreoffice.called is True
    assert powerpoint.called is False
    assert renderer.last_used_renderer == "libreoffice"


def test_powerpoint_renderer_uses_documented_windows_pdf_export_flow() -> None:
    renderer = PowerPointRenderer(platform_name="windows")

    script = renderer._build_powershell_script(Path(r"C:\deck.pptx"), Path(r"C:\deck.pdf"))

    assert "Presentations.Open($inputPath, 0, 0, 0)" in script
    assert "SaveAs($outputPath, 32)" in script


def test_powerpoint_renderer_rejects_unsupported_platform(tmp_path: Path) -> None:
    renderer = PowerPointRenderer(platform_name="linux")

    with pytest.raises(RendererError, match="not supported"):
        renderer.export_pptx_to_pdf(tmp_path / "deck.pptx", tmp_path / "deck.pdf")


def test_libreoffice_renderer_uses_path_lookup_when_no_explicit_binary(monkeypatch, tmp_path: Path) -> None:
    expected = tmp_path / "soffice.exe"
    expected.write_text("", encoding="utf-8")
    monkeypatch.setattr("app.services.renderer.shutil.which", lambda name: str(expected) if name == "soffice.exe" else None)

    renderer = LibreOfficeRenderer(None)

    assert renderer._resolve_binary() == expected
