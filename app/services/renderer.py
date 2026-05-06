from __future__ import annotations

import shutil
import subprocess
import time
from dataclasses import dataclass
from pathlib import Path
import re
import secrets
from typing import Optional, Protocol

from pptx import Presentation as PptxPresentation

from app.config import Settings


class RendererError(RuntimeError):
    """Raised when a renderer cannot produce a PDF export."""


class PDFExporter(Protocol):
    def export_pptx_to_pdf(self, pptx_path: Path, output_path: Path) -> Path:
        """Export a PowerPoint deck to PDF."""


class SlideImageExporter(Protocol):
    def export_pptx_to_slide_images(self, pptx_path: Path, output_dir: Path, *, dpi: int = 180) -> list[Path]:
        """Export each PowerPoint slide to an image file."""


@dataclass(frozen=True)
class RendererAvailability:
    can_convert: bool
    preferred_renderer: str
    libreoffice_available: bool
    powerpoint_available: bool
    message: str
    can_export_slide_images: bool = False
    slide_image_export_renderer: str = "none"
    slide_image_export_message: str = ""


class PowerPointRenderer:
    def __init__(
        self,
        platform_name: str,
        app_path: Optional[Path] = None,
        powershell_bin: str = "powershell.exe",
        mac_slide_export_macro_name: Optional[str] = None,
        mac_slide_export_staging_dir: Optional[Path] = None,
    ):
        self.platform_name = platform_name.lower()
        self.app_path = Path(app_path) if app_path is not None else None
        self.powershell_bin = powershell_bin
        self.mac_slide_export_macro_name = mac_slide_export_macro_name.strip() if mac_slide_export_macro_name else None
        self.mac_slide_export_staging_dir = (
            Path(mac_slide_export_staging_dir) if mac_slide_export_staging_dir is not None else None
        )

    def is_available(self) -> bool:
        if self.platform_name == "darwin":
            return self.app_path is not None and self.app_path.exists()
        if self.platform_name == "windows":
            return self._is_windows_powerpoint_registered()
        return False

    def export_pptx_to_pdf(self, pptx_path: Path, output_path: Path) -> Path:
        if self.platform_name == "darwin":
            return self._export_on_macos(pptx_path, output_path)
        if self.platform_name == "windows":
            return self._export_on_windows(pptx_path, output_path)
        raise RendererError(f"PowerPoint export is not supported on platform `{self.platform_name}`.")

    def export_pptx_to_slide_images(self, pptx_path: Path, output_dir: Path, *, dpi: int = 180) -> list[Path]:
        if self.platform_name == "darwin":
            return self._export_slide_images_on_macos(pptx_path, output_dir)
        if self.platform_name == "windows":
            return self._export_slide_images_on_windows(pptx_path, output_dir, dpi=dpi)
        raise RendererError(f"PowerPoint slide-image export is not supported on platform `{self.platform_name}`.")

    def _export_on_macos(self, pptx_path: Path, output_path: Path) -> Path:
        if self.app_path is None or not self.app_path.exists():
            raise RendererError("Microsoft PowerPoint is not installed at the expected macOS path.")

        output_path.parent.mkdir(parents=True, exist_ok=True)
        if output_path.exists():
            output_path.unlink()

        staged_input_path = self._stage_macos_input(pptx_path)

        scripts = [
            self._build_applescript(staged_input_path, output_path, use_posix_file=True),
        ]

        errors: list[str] = []
        try:
            for script in scripts:
                try:
                    completed = subprocess.run(
                        ["osascript", "-"],
                        input=script,
                        capture_output=True,
                        text=True,
                        timeout=60,
                    )
                except FileNotFoundError as exc:
                    raise RendererError("`osascript` is not available on this machine.") from exc
                except subprocess.TimeoutExpired:
                    errors.append("PowerPoint timed out while exporting the deck to PDF.")
                    continue
                if completed.returncode != 0:
                    stderr = completed.stderr.strip() or completed.stdout.strip() or "Unknown AppleScript error."
                    errors.append(stderr)
                    continue

                exported = self._wait_for_output(output_path)
                if exported is not None:
                    return exported
        finally:
            self._cleanup_staged_macos_input(staged_input_path)

        raise RendererError(
            "PowerPoint export did not produce a PDF."
            + (f" Last error: {errors[-1]}" if errors else "")
        )

    def _export_on_windows(self, pptx_path: Path, output_path: Path) -> Path:
        powershell = shutil.which(self.powershell_bin) or self.powershell_bin
        output_path.parent.mkdir(parents=True, exist_ok=True)
        if output_path.exists():
            output_path.unlink()

        try:
            completed = subprocess.run(
                [
                    powershell,
                    "-NoLogo",
                    "-NoProfile",
                    "-NonInteractive",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-Command",
                    self._build_powershell_script(pptx_path, output_path),
                ],
                capture_output=True,
                text=True,
                timeout=240,
            )
        except FileNotFoundError as exc:
            raise RendererError(f"Windows PowerShell executable `{self.powershell_bin}` was not found.") from exc
        if completed.returncode != 0:
            stderr = completed.stderr.strip() or completed.stdout.strip() or "Unknown PowerShell error."
            raise RendererError(f"Windows PowerPoint export failed: {stderr}")

        exported = self._wait_for_output(output_path)
        if exported is None:
            raise RendererError("Windows PowerPoint export did not produce a PDF.")
        return exported

    def _export_slide_images_on_macos(self, pptx_path: Path, output_dir: Path) -> list[Path]:
        if self.app_path is None or not self.app_path.exists():
            raise RendererError("Microsoft PowerPoint is not installed at the expected macOS path.")
        if not self.mac_slide_export_macro_name:
            raise RendererError(
                "No PowerPoint VBA slide-export macro is configured. "
                "Set `PDF_PPTX_POWERPOINT_SLIDE_EXPORT_MACRO_NAME` and keep the add-in loaded in PowerPoint."
            )

        export_root = self._prepare_macos_slide_export_root(output_dir)
        staged_input_path = self._stage_macos_input(pptx_path)

        scripts = [
            self._build_applescript_macro_slide_export(staged_input_path, export_root),
        ]

        errors: list[str] = []
        try:
            for script in scripts:
                try:
                    completed = subprocess.run(
                        ["osascript", "-"],
                        input=script,
                        capture_output=True,
                        text=True,
                        timeout=60,
                    )
                except FileNotFoundError as exc:
                    raise RendererError("`osascript` is not available on this machine.") from exc
                except subprocess.TimeoutExpired:
                    errors.append("PowerPoint timed out while exporting slide images.")
                    continue
                if completed.returncode != 0:
                    stderr = completed.stderr.strip() or completed.stdout.strip() or "Unknown AppleScript error."
                    errors.append(stderr)
                    continue

                exported = self._wait_for_slide_images(export_root)
                if exported:
                    finalized = self._finalize_macos_slide_exports(exported, output_dir)
                    return self._validate_exported_slide_images(pptx_path, finalized)
        finally:
            self._cleanup_staged_macos_input(staged_input_path)

        raise RendererError(
            "PowerPoint slide-image export did not produce PNG files."
            + (f" Last error: {errors[-1]}" if errors else "")
        )

    def _export_slide_images_on_windows(self, pptx_path: Path, output_dir: Path, *, dpi: int) -> list[Path]:
        powershell = shutil.which(self.powershell_bin) or self.powershell_bin
        shutil.rmtree(output_dir, ignore_errors=True)
        output_dir.mkdir(parents=True, exist_ok=True)
        export_root = output_dir / "slides"
        export_root.mkdir(parents=True, exist_ok=True)

        try:
            completed = subprocess.run(
                [
                    powershell,
                    "-NoLogo",
                    "-NoProfile",
                    "-NonInteractive",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-Command",
                    self._build_powershell_slide_export_script(pptx_path, export_root, dpi=dpi),
                ],
                capture_output=True,
                text=True,
                timeout=240,
            )
        except FileNotFoundError as exc:
            raise RendererError(f"Windows PowerShell executable `{self.powershell_bin}` was not found.") from exc
        if completed.returncode != 0:
            stderr = completed.stderr.strip() or completed.stdout.strip() or "Unknown PowerShell error."
            raise RendererError(f"Windows PowerPoint slide-image export failed: {stderr}")

        exported = self._wait_for_slide_images(output_dir)
        if not exported:
            raise RendererError("Windows PowerPoint slide-image export did not produce PNG files.")
        return self._validate_exported_slide_images(pptx_path, exported)

    def _build_applescript(self, pptx_path: Path, output_path: Path, use_posix_file: bool) -> str:
        input_path = self._escape_applescript_string(str(pptx_path))
        output_value = self._escape_applescript_string(str(output_path))
        if use_posix_file:
            output_target = f'POSIX file "{output_value}"'
        else:
            output_target = f'"{output_value}"'

        return f"""
set inputPath to POSIX file "{input_path}"
tell application "Microsoft PowerPoint"
  open inputPath
  set currentPresentation to active presentation
  save currentPresentation in {output_target} as save as PDF file
  delay 2
  close currentPresentation saving no
end tell
"""

    def _build_applescript_slide_export(self, pptx_path: Path, output_dir: Path, use_posix_file: bool) -> str:
        input_path = self._escape_applescript_string(str(pptx_path))
        output_value = self._escape_applescript_string(str(output_dir))
        if use_posix_file:
            output_target = f'POSIX file "{output_value}"'
        else:
            output_target = f'"{output_value}"'

        return f"""
set inputPath to POSIX file "{input_path}"
tell application "Microsoft PowerPoint"
  open inputPath
  set currentPresentation to active presentation
  save currentPresentation in {output_target} as save as PNG
  delay 2
  close currentPresentation saving no
end tell
"""

    def _build_applescript_macro_slide_export(self, pptx_path: Path, output_dir: Path) -> str:
        if not self.mac_slide_export_macro_name:
            raise RendererError("No PowerPoint VBA slide-export macro is configured.")

        input_path = self._escape_applescript_string(str(pptx_path))
        output_folder = self._escape_applescript_string(f"{output_dir.as_posix().rstrip('/')}/")
        macro_name = self._escape_applescript_string(self.mac_slide_export_macro_name)
        staged_name = self._escape_applescript_string(pptx_path.name)

        return f"""
set inputPath to POSIX file "{input_path}"
set outputFolder to "{output_folder}"
set expectedPresentationName to "{staged_name}"
do shell script "mkdir -p " & quoted form of outputFolder

tell application "Microsoft PowerPoint"
  activate
  open inputPath
  set currentPresentation to active presentation
  if name of currentPresentation is not expectedPresentationName then
    error "Active presentation after opening does not match the staged file."
  end if
  try
    run VB macro macro name "{macro_name}" list of parameters {{outputFolder}}
    delay 2
    close currentPresentation saving no
  on error errMsg number errNum
    try
      close currentPresentation saving no
    end try
    error "PowerPoint macro slide export failed (" & errNum & "): " & errMsg
  end try
end tell
"""

    def _build_powershell_script(self, pptx_path: Path, output_path: Path) -> str:
        input_path = self._escape_powershell_string(str(pptx_path))
        output_value = self._escape_powershell_string(str(output_path))
        return f"""
$ErrorActionPreference = 'Stop'
$inputPath = '{input_path}'
$outputPath = '{output_value}'
$powerPoint = $null
$presentation = $null
try {{
  $powerPoint = New-Object -ComObject PowerPoint.Application
  $powerPoint.Visible = 0
  $presentation = $powerPoint.Presentations.Open($inputPath, 0, 0, 0)
  $presentation.SaveAs($outputPath, 32)
}}
finally {{
  if ($presentation -ne $null) {{
    $presentation.Close()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation)
  }}
  if ($powerPoint -ne $null) {{
    $powerPoint.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint)
  }}
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}}
"""

    def _build_powershell_slide_export_script(self, pptx_path: Path, output_dir: Path, *, dpi: int) -> str:
        input_path = self._escape_powershell_string(str(pptx_path))
        output_value = self._escape_powershell_string(str(output_dir))
        safe_dpi = max(72, int(dpi))
        return f"""
$ErrorActionPreference = 'Stop'
$inputPath = '{input_path}'
$outputDir = '{output_value}'
$dpi = {safe_dpi}
$powerPoint = $null
$presentation = $null
try {{
  if (-not (Test-Path -LiteralPath $outputDir)) {{
    New-Item -ItemType Directory -Path $outputDir | Out-Null
  }}
  $powerPoint = New-Object -ComObject PowerPoint.Application
  $powerPoint.Visible = 0
  $presentation = $powerPoint.Presentations.Open($inputPath, 0, 0, 0)
  $scaleWidth = [Math]::Max(1, [int][Math]::Round(($presentation.PageSetup.SlideWidth / 72.0) * $dpi))
  $scaleHeight = [Math]::Max(1, [int][Math]::Round(($presentation.PageSetup.SlideHeight / 72.0) * $dpi))
  $presentation.Export($outputDir, 'PNG', $scaleWidth, $scaleHeight)
}}
finally {{
  if ($presentation -ne $null) {{
    $presentation.Close()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation)
  }}
  if ($powerPoint -ne $null) {{
    $powerPoint.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint)
  }}
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}}
"""

    @staticmethod
    def _escape_applescript_string(value: str) -> str:
        return value.replace("\\", "\\\\").replace('"', '\\"')

    @staticmethod
    def _escape_powershell_string(value: str) -> str:
        return value.replace("'", "''")

    @staticmethod
    def _wait_for_output(expected_output: Path) -> Optional[Path]:
        candidates = [expected_output, expected_output.parent / expected_output.name]
        for _ in range(16):
            for candidate in candidates:
                if candidate.exists() and candidate.stat().st_size > 0:
                    return candidate
            time.sleep(1)
        return None

    def _stage_macos_input(self, pptx_path: Path) -> Path:
        if self.platform_name != "darwin":
            return pptx_path
        sandbox_dir = self._macos_sandbox_documents_dir()
        sandbox_dir.mkdir(parents=True, exist_ok=True)
        staged_path = sandbox_dir / f"codex-{secrets.token_hex(8)}-{pptx_path.name}"
        shutil.copy2(pptx_path, staged_path)
        return staged_path

    def _cleanup_staged_macos_input(self, staged_path: Path) -> None:
        if self.platform_name != "darwin":
            return
        sandbox_dir = self._macos_sandbox_documents_dir()
        try:
            if staged_path.parent == sandbox_dir and staged_path.exists():
                staged_path.unlink()
        except OSError:
            pass

    @staticmethod
    def _macos_sandbox_documents_dir() -> Path:
        return Path.home() / "Library" / "Containers" / "com.microsoft.Powerpoint" / "Data" / "Documents"

    @staticmethod
    def _wait_for_slide_images(search_root: Path) -> list[Path]:
        for _ in range(20):
            images = PowerPointRenderer._collect_slide_images(search_root)
            if images:
                return images
            time.sleep(1)
        return []

    @staticmethod
    def _collect_slide_images(search_root: Path) -> list[Path]:
        image_paths = [
            path
            for path in search_root.rglob("*")
            if path.is_file() and path.suffix.lower() == ".png" and path.stat().st_size > 0
        ]
        return sorted(image_paths, key=PowerPointRenderer._natural_path_sort_key)

    def _prepare_macos_slide_export_root(self, output_dir: Path) -> Path:
        output_dir = Path(output_dir)
        shutil.rmtree(output_dir, ignore_errors=True)
        output_dir.mkdir(parents=True, exist_ok=True)

        if self.mac_slide_export_staging_dir is None:
            export_root = output_dir / "slides"
            export_root.mkdir(parents=True, exist_ok=True)
            return export_root

        staging_dir = Path(self.mac_slide_export_staging_dir)
        shutil.rmtree(staging_dir, ignore_errors=True)
        staging_dir.mkdir(parents=True, exist_ok=True)
        return staging_dir

    @staticmethod
    def _finalize_macos_slide_exports(exported_paths: list[Path], output_dir: Path) -> list[Path]:
        destination_root = Path(output_dir) / "slides"
        shutil.rmtree(destination_root, ignore_errors=True)
        destination_root.mkdir(parents=True, exist_ok=True)

        finalized: list[Path] = []
        for exported_path in exported_paths:
            destination = destination_root / exported_path.name
            shutil.copy2(exported_path, destination)
            finalized.append(destination)
        return finalized

    @staticmethod
    def _expected_slide_count(pptx_path: Path) -> int:
        presentation = PptxPresentation(str(pptx_path))
        return len(presentation.slides)

    def _validate_exported_slide_images(self, pptx_path: Path, exported_paths: list[Path]) -> list[Path]:
        expected_count = self._expected_slide_count(pptx_path)
        actual_count = len(exported_paths)
        if actual_count != expected_count:
            raise RendererError(
                "PowerPoint slide-image export produced the wrong number of PNG files "
                f"({actual_count} exported for {expected_count} slides)."
            )
        return exported_paths

    @staticmethod
    def _natural_path_sort_key(path: Path) -> list[object]:
        return [int(part) if part.isdigit() else part.lower() for part in re.split(r"(\d+)", path.name)]

    def _is_windows_powerpoint_registered(self) -> bool:
        if self.platform_name != "windows":
            return False
        try:
            import winreg  # type: ignore[attr-defined]
        except ImportError:
            return False

        registry_paths = (
            r"PowerPoint.Application\CurVer",
            r"PowerPoint.Application\CLSID",
        )
        for registry_path in registry_paths:
            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, registry_path):
                    return True
            except OSError:
                continue
        return False


class LibreOfficeRenderer:
    def __init__(self, binary_path: Optional[Path]):
        self.binary_path = Path(binary_path) if binary_path is not None else None

    def is_available(self) -> bool:
        return self._resolve_binary() is not None

    def export_pptx_to_pdf(self, pptx_path: Path, output_path: Path) -> Path:
        binary = self._resolve_binary()
        if binary is None:
            raise RendererError("LibreOffice is not installed or `soffice` is not available on PATH.")

        output_path.parent.mkdir(parents=True, exist_ok=True)
        completed = subprocess.run(
            [
                str(binary),
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(output_path.parent),
                str(pptx_path),
            ],
            capture_output=True,
            text=True,
            timeout=240,
        )
        if completed.returncode != 0:
            stderr = completed.stderr.strip() or completed.stdout.strip() or "Unknown LibreOffice error."
            raise RendererError(f"LibreOffice export failed: {stderr}")

        generated = output_path.parent / f"{pptx_path.stem}.pdf"
        if not generated.exists():
            raise RendererError("LibreOffice reported success but did not write a PDF.")
        if generated != output_path:
            if output_path.exists():
                output_path.unlink()
            shutil.move(str(generated), str(output_path))
        return output_path

    def _resolve_binary(self) -> Optional[Path]:
        configured = self.binary_path
        if configured is not None and configured.exists():
            return configured

        for candidate_name in ("soffice", "soffice.exe", "soffice.com"):
            discovered = shutil.which(candidate_name)
            if discovered:
                return Path(discovered)
        return configured if configured is not None and configured.exists() else None


class Renderer:
    def __init__(
        self,
        platform_name: str,
        powerpoint: PDFExporter,
        libreoffice: PDFExporter,
        enable_powerpoint_fallback: bool = False,
    ):
        self.platform_name = platform_name.lower()
        self.powerpoint = powerpoint
        self.libreoffice = libreoffice
        self.enable_powerpoint_fallback = enable_powerpoint_fallback
        self.last_used_renderer = "unknown"

    def availability(self) -> RendererAvailability:
        libreoffice_available = getattr(self.libreoffice, "is_available", lambda: True)()
        powerpoint_available = getattr(self.powerpoint, "is_available", lambda: False)()
        macro_slide_export_available = bool(
            self.platform_name == "darwin"
            and getattr(self.powerpoint, "mac_slide_export_macro_name", None)
        )
        can_export_slide_images = powerpoint_available and (
            self.platform_name == "windows" or macro_slide_export_available
        )
        slide_image_export_renderer = "powerpoint" if powerpoint_available else "none"
        if can_export_slide_images:
            if macro_slide_export_available:
                macro_name = getattr(self.powerpoint, "mac_slide_export_macro_name", "ExportSlidesToFolder")
                slide_image_export_message = (
                    f"Microsoft PowerPoint will export slide images through the VBA macro `{macro_name}`. "
                    "Keep the add-in loaded in PowerPoint."
                )
            elif self.platform_name == "windows":
                slide_image_export_message = (
                    "Microsoft PowerPoint is ready for direct slide-image export on Windows via PowerPoint automation."
                )
            else:
                slide_image_export_message = "Microsoft PowerPoint is ready for direct slide-image export."
        elif powerpoint_available and self.platform_name == "darwin":
            slide_image_export_message = (
                "Direct PowerPoint slide-image export on macOS requires a loaded VBA macro add-in."
            )
        else:
            slide_image_export_message = "Direct slide-image export requires Microsoft PowerPoint on this machine."

        if libreoffice_available:
            message = "LibreOffice is ready for conversion."
            if powerpoint_available:
                message += " AI comparison prefers PowerPoint-rendered PDF pages when PowerPoint is available."
                if self.enable_powerpoint_fallback:
                    message += " Microsoft PowerPoint is also available as a fallback."
            return RendererAvailability(
                can_convert=True,
                preferred_renderer="libreoffice",
                libreoffice_available=libreoffice_available,
                powerpoint_available=powerpoint_available,
                message=message,
                can_export_slide_images=can_export_slide_images,
                slide_image_export_renderer=slide_image_export_renderer,
                slide_image_export_message=slide_image_export_message,
            )

        if powerpoint_available:
            return RendererAvailability(
                can_convert=True,
                preferred_renderer="powerpoint",
                libreoffice_available=libreoffice_available,
                powerpoint_available=powerpoint_available,
                message=(
                    "LibreOffice is not installed. Microsoft PowerPoint will be used for conversion on this machine."
                ),
                can_export_slide_images=can_export_slide_images,
                slide_image_export_renderer=slide_image_export_renderer,
                slide_image_export_message=slide_image_export_message,
            )

        install_message = (
            "LibreOffice is not installed and Microsoft PowerPoint was not found."
            " Install LibreOffice or Microsoft PowerPoint, then reopen the app."
        )
        if self.platform_name == "windows":
            install_message = (
                "LibreOffice is not installed and Microsoft PowerPoint was not found."
                " Install LibreOffice or Microsoft PowerPoint on Windows, then reopen the app."
            )
        return RendererAvailability(
            can_convert=False,
            preferred_renderer="none",
            libreoffice_available=libreoffice_available,
            powerpoint_available=powerpoint_available,
            message=install_message,
            can_export_slide_images=can_export_slide_images,
            slide_image_export_renderer=slide_image_export_renderer,
            slide_image_export_message=slide_image_export_message,
        )

    def export_pptx_to_pdf(
        self,
        pptx_path: Path,
        output_path: Path,
        *,
        preferred_renderer: str | None = None,
        allow_fallback: bool = True,
    ) -> Path:
        availability = self.availability()
        if not availability.can_convert:
            raise RendererError(availability.message)

        errors: list[str] = []
        renderers = self._pdf_renderers_for_request(
            availability=availability,
            preferred_renderer=preferred_renderer,
            allow_fallback=allow_fallback,
        )
        if not renderers:
            if preferred_renderer:
                raise RendererError(f"The requested renderer `{preferred_renderer}` is not available on this machine.")
            raise RendererError(availability.message)

        for name, renderer in renderers:
            try:
                exported = renderer.export_pptx_to_pdf(pptx_path, output_path)
                self.last_used_renderer = name
                return exported
            except RendererError as exc:
                errors.append(f"{name}: {exc}")

        raise RendererError(" ; ".join(errors))

    def can_export_slide_images(self) -> bool:
        return self.availability().can_export_slide_images

    def export_pptx_to_slide_images(self, pptx_path: Path, output_dir: Path, *, dpi: int = 180) -> list[Path]:
        availability = self.availability()
        if not availability.can_export_slide_images:
            raise RendererError(availability.slide_image_export_message)
        exporter = self.powerpoint
        if not hasattr(exporter, "export_pptx_to_slide_images"):
            raise RendererError("The configured PowerPoint renderer does not support slide-image export.")
        return exporter.export_pptx_to_slide_images(pptx_path, output_dir, dpi=dpi)  # type: ignore[attr-defined]

    def is_powerpoint_available(self) -> bool:
        return self.availability().powerpoint_available

    def _pdf_renderers_for_request(
        self,
        *,
        availability: RendererAvailability,
        preferred_renderer: str | None,
        allow_fallback: bool,
    ) -> list[tuple[str, PDFExporter]]:
        candidates: list[tuple[str, PDFExporter]] = []

        def append_renderer(name: str) -> None:
            if name == "powerpoint" and availability.powerpoint_available:
                candidates.append(("powerpoint", self.powerpoint))
            elif name == "libreoffice" and availability.libreoffice_available:
                candidates.append(("libreoffice", self.libreoffice))

        normalized_preference = (preferred_renderer or "").strip().lower() or None
        if normalized_preference is not None:
            append_renderer(normalized_preference)
            if allow_fallback:
                fallback_name = "libreoffice" if normalized_preference == "powerpoint" else "powerpoint"
                append_renderer(fallback_name)
        else:
            if availability.libreoffice_available:
                append_renderer("libreoffice")
            if availability.powerpoint_available and (
                self.enable_powerpoint_fallback or not availability.libreoffice_available
            ):
                append_renderer("powerpoint")

        seen: set[str] = set()
        ordered: list[tuple[str, PDFExporter]] = []
        for name, exporter in candidates:
            if name in seen:
                continue
            seen.add(name)
            ordered.append((name, exporter))
        return ordered


def build_renderer(settings: Settings) -> Renderer:
    return Renderer(
        platform_name=settings.platform_name,
        powerpoint=PowerPointRenderer(
            platform_name=settings.platform_name,
            app_path=settings.powerpoint_app_path,
            powershell_bin=settings.powershell_bin,
            mac_slide_export_macro_name=settings.powerpoint_slide_export_macro_name,
            mac_slide_export_staging_dir=settings.powerpoint_slide_export_staging_dir,
        ),
        libreoffice=LibreOfficeRenderer(settings.libreoffice_bin),
        enable_powerpoint_fallback=settings.enable_powerpoint_fallback,
    )
