from __future__ import annotations

from pathlib import Path

import pytest
from pptx import Presentation

from app.services.renderer import LibreOfficeRenderer, PowerPointRenderer, Renderer, RendererError


class FailingRenderer:
    def is_available(self) -> bool:
        return True

    def export_pptx_to_pdf(self, pptx_path: Path, output_path: Path) -> Path:
        raise RendererError("primary failed")

    def export_pptx_to_slide_images(self, pptx_path: Path, output_dir: Path, *, dpi: int = 180) -> list[Path]:
        raise RendererError("primary failed")


class SuccessfulRenderer:
    def is_available(self) -> bool:
        return True

    def export_pptx_to_pdf(self, pptx_path: Path, output_path: Path) -> Path:
        output_path.write_bytes(b"%PDF-1.4")
        return output_path

    def export_pptx_to_slide_images(self, pptx_path: Path, output_dir: Path, *, dpi: int = 180) -> list[Path]:
        output_dir.mkdir(parents=True, exist_ok=True)
        image_path = output_dir / "Slide1.PNG"
        image_path.write_bytes(b"png")
        return [image_path]


class RecordingRenderer:
    def __init__(self, should_fail: bool = False, available: bool = True):
        self.called = False
        self.slide_image_called = False
        self.should_fail = should_fail
        self.available = available
        self.mac_slide_export_macro_name = None

    def is_available(self) -> bool:
        return self.available

    def export_pptx_to_pdf(self, pptx_path: Path, output_path: Path) -> Path:
        self.called = True
        if self.should_fail:
            raise RendererError("renderer failed")
        output_path.write_bytes(b"%PDF-1.4")
        return output_path

    def export_pptx_to_slide_images(self, pptx_path: Path, output_dir: Path, *, dpi: int = 180) -> list[Path]:
        self.slide_image_called = True
        if self.should_fail:
            raise RendererError("renderer failed")
        output_dir.mkdir(parents=True, exist_ok=True)
        image_path = output_dir / "Slide1.PNG"
        image_path.write_bytes(b"png")
        return [image_path]


def test_renderer_falls_back_to_secondary(tmp_path: Path) -> None:
    renderer = Renderer(
        platform_name="darwin",
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
        platform_name="darwin",
        powerpoint=powerpoint,
        libreoffice=libreoffice,
        enable_powerpoint_fallback=False,
    )

    exported = renderer.export_pptx_to_pdf(tmp_path / "deck.pptx", tmp_path / "candidate.pdf")

    assert exported.exists()
    assert libreoffice.called is True
    assert powerpoint.called is False
    assert renderer.last_used_renderer == "libreoffice"


def test_renderer_uses_powerpoint_when_libreoffice_is_missing(tmp_path: Path) -> None:
    powerpoint = RecordingRenderer(available=True)
    libreoffice = RecordingRenderer(available=False)
    renderer = Renderer(
        platform_name="darwin",
        powerpoint=powerpoint,
        libreoffice=libreoffice,
        enable_powerpoint_fallback=False,
    )

    exported = renderer.export_pptx_to_pdf(tmp_path / "deck.pptx", tmp_path / "candidate.pdf")

    assert exported.exists()
    assert libreoffice.called is False
    assert powerpoint.called is True
    assert renderer.last_used_renderer == "powerpoint"
    assert renderer.availability().can_export_slide_images is False


def test_renderer_can_prefer_powerpoint_without_fallback(tmp_path: Path) -> None:
    powerpoint = RecordingRenderer(available=True)
    libreoffice = RecordingRenderer(available=True)
    renderer = Renderer(
        platform_name="darwin",
        powerpoint=powerpoint,
        libreoffice=libreoffice,
        enable_powerpoint_fallback=False,
    )

    exported = renderer.export_pptx_to_pdf(
        tmp_path / "deck.pptx",
        tmp_path / "candidate.pdf",
        preferred_renderer="powerpoint",
        allow_fallback=False,
    )

    assert exported.exists()
    assert powerpoint.called is True
    assert libreoffice.called is False
    assert renderer.last_used_renderer == "powerpoint"


def test_renderer_rejects_unavailable_preferred_renderer(tmp_path: Path) -> None:
    renderer = Renderer(
        platform_name="darwin",
        powerpoint=RecordingRenderer(available=False),
        libreoffice=RecordingRenderer(available=True),
        enable_powerpoint_fallback=False,
    )

    with pytest.raises(RendererError, match="requested renderer `powerpoint`"):
        renderer.export_pptx_to_pdf(
            tmp_path / "deck.pptx",
            tmp_path / "candidate.pdf",
            preferred_renderer="powerpoint",
            allow_fallback=False,
        )


def test_renderer_reports_install_guidance_when_no_converter_is_available(tmp_path: Path) -> None:
    renderer = Renderer(
        platform_name="darwin",
        powerpoint=RecordingRenderer(available=False),
        libreoffice=RecordingRenderer(available=False),
        enable_powerpoint_fallback=False,
    )

    availability = renderer.availability()

    assert availability.can_convert is False
    assert availability.preferred_renderer == "none"
    assert availability.can_export_slide_images is False
    assert "Install LibreOffice or Microsoft PowerPoint" in availability.message

    with pytest.raises(RendererError, match="Install LibreOffice or Microsoft PowerPoint"):
        renderer.export_pptx_to_pdf(tmp_path / "deck.pptx", tmp_path / "candidate.pdf")


def test_powerpoint_renderer_uses_documented_windows_pdf_export_flow() -> None:
    renderer = PowerPointRenderer(platform_name="windows")

    script = renderer._build_powershell_script(Path(r"C:\deck.pptx"), Path(r"C:\deck.pdf"))

    assert "Presentations.Open($inputPath, 0, 0, 0)" in script
    assert "SaveAs($outputPath, 32)" in script


def test_powerpoint_renderer_uses_documented_windows_slide_image_export_flow() -> None:
    renderer = PowerPointRenderer(platform_name="windows")

    script = renderer._build_powershell_slide_export_script(Path(r"C:\deck.pptx"), Path(r"C:\slides"), dpi=180)

    assert "Presentations.Open($inputPath, 0, 0, 0)" in script
    assert "$presentation.Export($outputDir, 'PNG', $scaleWidth, $scaleHeight)" in script
    assert "$presentation.PageSetup.SlideWidth" in script
    assert "$presentation.PageSetup.SlideHeight" in script


def test_powerpoint_renderer_targets_staged_presentation_for_macro_slide_export(tmp_path: Path) -> None:
    renderer = PowerPointRenderer(
        platform_name="darwin",
        mac_slide_export_macro_name="ExportSlidesToFolder",
    )

    script = renderer._build_applescript_macro_slide_export(
        tmp_path / "codex-123-candidate.pptx",
        tmp_path / "slides",
    )

    assert 'set expectedPresentationName to "codex-123-candidate.pptx"' in script
    assert "set currentPresentation to active presentation" in script
    assert 'if name of currentPresentation is not expectedPresentationName then' in script


def test_powerpoint_renderer_natural_path_sort_handles_double_digits() -> None:
    paths = [
        Path("Slide_1.png"),
        Path("Slide_10.png"),
        Path("Slide_11.png"),
        Path("Slide_2.png"),
    ]

    ordered = sorted(paths, key=PowerPointRenderer._natural_path_sort_key)

    assert [path.name for path in ordered] == [
        "Slide_1.png",
        "Slide_2.png",
        "Slide_10.png",
        "Slide_11.png",
    ]


def test_powerpoint_renderer_validates_exported_slide_count(tmp_path: Path) -> None:
    presentation = Presentation()
    presentation.slides.add_slide(presentation.slide_layouts[6])
    presentation.slides.add_slide(presentation.slide_layouts[6])
    pptx_path = tmp_path / "deck.pptx"
    presentation.save(pptx_path)

    renderer = PowerPointRenderer(platform_name="darwin")
    exported = [tmp_path / "Slide_1.png"]
    exported[0].write_bytes(b"png")

    with pytest.raises(RendererError, match="wrong number of PNG files"):
        renderer._validate_exported_slide_images(pptx_path, exported)


def test_renderer_exports_slide_images_via_powerpoint_only(tmp_path: Path) -> None:
    powerpoint = RecordingRenderer(available=True)
    libreoffice = RecordingRenderer(available=True)
    renderer = Renderer(
        platform_name="windows",
        powerpoint=powerpoint,
        libreoffice=libreoffice,
        enable_powerpoint_fallback=False,
    )

    exported = renderer.export_pptx_to_slide_images(tmp_path / "deck.pptx", tmp_path / "slide-images")

    assert [path.name for path in exported] == ["Slide1.PNG"]
    assert powerpoint.slide_image_called is True
    assert libreoffice.slide_image_called is False


def test_renderer_reports_windows_slide_image_export_availability() -> None:
    renderer = Renderer(
        platform_name="windows",
        powerpoint=RecordingRenderer(available=True),
        libreoffice=RecordingRenderer(available=True),
        enable_powerpoint_fallback=True,
    )

    availability = renderer.availability()

    assert availability.can_export_slide_images is True
    assert availability.slide_image_export_renderer == "powerpoint"
    assert "PowerPoint automation" in availability.slide_image_export_message


def test_renderer_reports_macos_slide_image_export_fallback_message() -> None:
    renderer = Renderer(
        platform_name="darwin",
        powerpoint=RecordingRenderer(available=True),
        libreoffice=RecordingRenderer(available=True),
        enable_powerpoint_fallback=False,
    )

    availability = renderer.availability()

    assert availability.can_export_slide_images is False
    assert "requires a loaded VBA macro add-in" in availability.slide_image_export_message


def test_renderer_reports_macos_macro_slide_image_export_availability() -> None:
    powerpoint = RecordingRenderer(available=True)
    powerpoint.mac_slide_export_macro_name = "ExportSlidesToFolder"
    renderer = Renderer(
        platform_name="darwin",
        powerpoint=powerpoint,
        libreoffice=RecordingRenderer(available=True),
        enable_powerpoint_fallback=False,
    )

    availability = renderer.availability()

    assert availability.can_export_slide_images is True
    assert availability.slide_image_export_renderer == "powerpoint"
    assert "ExportSlidesToFolder" in availability.slide_image_export_message


def test_powerpoint_renderer_builds_macos_macro_slide_export_script(tmp_path: Path) -> None:
    renderer = PowerPointRenderer(
        platform_name="darwin",
        mac_slide_export_macro_name="ExportSlidesToFolder",
    )

    script = renderer._build_applescript_macro_slide_export(tmp_path / "deck.pptx", tmp_path / "slides")

    assert 'run VB macro macro name "ExportSlidesToFolder" list of parameters {outputFolder}' in script
    assert 'set outputFolder to "' in script
    assert 'mkdir -p ' in script


def test_powerpoint_renderer_finalizes_macos_slide_exports_into_job_folder(tmp_path: Path) -> None:
    staging_dir = tmp_path / "staging"
    staging_dir.mkdir(parents=True)
    exported = []
    for name in ("Slide_1.png", "Slide_2.png"):
        path = staging_dir / name
        path.write_bytes(b"png")
        exported.append(path)

    finalized = PowerPointRenderer._finalize_macos_slide_exports(exported, tmp_path / "job-output")

    assert [path.name for path in finalized] == ["Slide_1.png", "Slide_2.png"]
    assert all(path.exists() for path in finalized)
    assert finalized[0].parent.name == "slides"


def test_powerpoint_renderer_stages_macos_input_inside_powerpoint_container(tmp_path: Path, monkeypatch) -> None:
    sandbox_dir = tmp_path / "sandbox-documents"
    monkeypatch.setattr(
        PowerPointRenderer,
        "_macos_sandbox_documents_dir",
        staticmethod(lambda: sandbox_dir),
    )
    pptx_path = tmp_path / "deck.pptx"
    pptx_path.write_bytes(b"pptx")
    renderer = PowerPointRenderer(platform_name="darwin")

    staged = renderer._stage_macos_input(pptx_path)

    assert staged.parent == sandbox_dir
    assert staged.read_bytes() == b"pptx"
    assert staged.name.endswith("-deck.pptx")

    renderer._cleanup_staged_macos_input(staged)
    assert staged.exists() is False


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
