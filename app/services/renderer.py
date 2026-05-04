from __future__ import annotations

import shutil
import subprocess
import time
from pathlib import Path
from typing import Optional, Protocol

from app.config import Settings


class RendererError(RuntimeError):
    """Raised when a renderer cannot produce a PDF export."""


class PDFExporter(Protocol):
    def export_pptx_to_pdf(self, pptx_path: Path, output_path: Path) -> Path:
        """Export a PowerPoint deck to PDF."""


class PowerPointRenderer:
    def __init__(
        self,
        platform_name: str,
        app_path: Optional[Path] = None,
        powershell_bin: str = "powershell.exe",
    ):
        self.platform_name = platform_name.lower()
        self.app_path = Path(app_path) if app_path is not None else None
        self.powershell_bin = powershell_bin

    def export_pptx_to_pdf(self, pptx_path: Path, output_path: Path) -> Path:
        if self.platform_name == "darwin":
            return self._export_on_macos(pptx_path, output_path)
        if self.platform_name == "windows":
            return self._export_on_windows(pptx_path, output_path)
        raise RendererError(f"PowerPoint export is not supported on platform `{self.platform_name}`.")

    def _export_on_macos(self, pptx_path: Path, output_path: Path) -> Path:
        if self.app_path is None or not self.app_path.exists():
            raise RendererError("Microsoft PowerPoint is not installed at the expected macOS path.")

        output_path.parent.mkdir(parents=True, exist_ok=True)
        if output_path.exists():
            output_path.unlink()

        scripts = [
            self._build_applescript(pptx_path, output_path, use_posix_file=False),
            self._build_applescript(pptx_path, output_path, use_posix_file=True),
        ]

        errors: list[str] = []
        for script in scripts:
            try:
                completed = subprocess.run(
                    ["osascript", "-"],
                    input=script,
                    capture_output=True,
                    text=True,
                    timeout=240,
                )
            except FileNotFoundError as exc:
                raise RendererError("`osascript` is not available on this machine.") from exc
            if completed.returncode != 0:
                stderr = completed.stderr.strip() or completed.stdout.strip() or "Unknown AppleScript error."
                errors.append(stderr)
                continue

            exported = self._wait_for_output(output_path)
            if exported is not None:
                return exported

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
  activate
  open inputPath
  set currentPresentation to active presentation
  save currentPresentation in {output_target} as save as PDF file
  delay 2
  close currentPresentation saving no
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


class LibreOfficeRenderer:
    def __init__(self, binary_path: Optional[Path]):
        self.binary_path = Path(binary_path) if binary_path is not None else None

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
        powerpoint: PDFExporter,
        libreoffice: PDFExporter,
        enable_powerpoint_fallback: bool = False,
    ):
        self.powerpoint = powerpoint
        self.libreoffice = libreoffice
        self.enable_powerpoint_fallback = enable_powerpoint_fallback
        self.last_used_renderer = "unknown"

    def export_pptx_to_pdf(self, pptx_path: Path, output_path: Path) -> Path:
        errors: list[str] = []
        renderers: list[tuple[str, PDFExporter]] = [("libreoffice", self.libreoffice)]
        if self.enable_powerpoint_fallback:
            renderers.append(("powerpoint", self.powerpoint))

        for name, renderer in renderers:
            try:
                exported = renderer.export_pptx_to_pdf(pptx_path, output_path)
                self.last_used_renderer = name
                return exported
            except RendererError as exc:
                errors.append(f"{name}: {exc}")

        raise RendererError(" ; ".join(errors))


def build_renderer(settings: Settings) -> Renderer:
    return Renderer(
        powerpoint=PowerPointRenderer(
            platform_name=settings.platform_name,
            app_path=settings.powerpoint_app_path,
            powershell_bin=settings.powershell_bin,
        ),
        libreoffice=LibreOfficeRenderer(settings.libreoffice_bin),
        enable_powerpoint_fallback=settings.enable_powerpoint_fallback,
    )
