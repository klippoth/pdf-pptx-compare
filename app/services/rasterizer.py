from __future__ import annotations

import shutil
import subprocess
from pathlib import Path
import re

import fitz
import numpy as np
from PIL import Image

from app.services.models import PageImage


class RasterizerError(RuntimeError):
    """Raised when a PDF cannot be rasterized."""


class Rasterizer:
    def __init__(self, dpi: int = 180, poppler_bin_dir: Path | None = None):
        self.dpi = dpi
        self.poppler_bin_dir = Path(poppler_bin_dir) if poppler_bin_dir is not None else None
        self.last_used_engine = "fitz"

    def render_pdf(self, pdf_path: Path, output_dir: Path, *, engine: str = "auto") -> list[PageImage]:
        normalized_engine = engine.lower()
        if normalized_engine not in {"auto", "fitz", "poppler"}:
            raise ValueError(f"Unsupported rasterizer engine `{engine}`.")

        output_dir = Path(output_dir)
        shutil.rmtree(output_dir, ignore_errors=True)
        output_dir.mkdir(parents=True, exist_ok=True)

        if normalized_engine == "fitz":
            pages = self._render_with_fitz(pdf_path, output_dir)
            self.last_used_engine = "fitz"
            return pages

        if normalized_engine == "poppler":
            pages = self._render_with_poppler(pdf_path, output_dir)
            self.last_used_engine = "poppler"
            return pages

        if self.can_render_with_poppler():
            try:
                pages = self._render_with_poppler(pdf_path, output_dir)
                self.last_used_engine = "poppler"
                return pages
            except RasterizerError:
                # Fall back to PyMuPDF when Poppler is unavailable or fails at runtime.
                pass

        pages = self._render_with_fitz(pdf_path, output_dir)
        self.last_used_engine = "fitz"
        return pages

    def load_images(self, image_paths: list[Path]) -> list[PageImage]:
        pages: list[PageImage] = []
        for index, image_path in enumerate(image_paths):
            image = np.array(Image.open(image_path).convert("RGB"))
            pages.append(PageImage(page_index=index, image=image, image_path=image_path))
        return pages

    def can_render_with_poppler(self) -> bool:
        return self._resolve_poppler_binary("pdftocairo") is not None or self._resolve_poppler_binary("pdftoppm") is not None

    def _render_with_fitz(self, pdf_path: Path, output_dir: Path) -> list[PageImage]:
        document = fitz.open(str(pdf_path))
        scale = self.dpi / 72.0
        matrix = fitz.Matrix(scale, scale)
        pages: list[PageImage] = []

        for index, page in enumerate(document):
            pixmap = page.get_pixmap(matrix=matrix, alpha=False)
            image = np.frombuffer(pixmap.samples, dtype=np.uint8).reshape(
                pixmap.height,
                pixmap.width,
                pixmap.n,
            )[:, :, :3].copy()

            image_path = output_dir / f"page-{index + 1:03d}.png"
            Image.fromarray(image).save(image_path)
            pages.append(PageImage(page_index=index, image=image, image_path=image_path))

        return pages

    def _render_with_poppler(self, pdf_path: Path, output_dir: Path) -> list[PageImage]:
        pdftocairo = self._resolve_poppler_binary("pdftocairo")
        pdftoppm = self._resolve_poppler_binary("pdftoppm")

        prefix = output_dir / "page"
        if pdftocairo is not None:
            command = [
                str(pdftocairo),
                "-png",
                "-r",
                str(self.dpi),
                str(pdf_path),
                str(prefix),
            ]
        elif pdftoppm is not None:
            command = [
                str(pdftoppm),
                "-png",
                "-r",
                str(self.dpi),
                str(pdf_path),
                str(prefix),
            ]
        else:
            raise RasterizerError("Poppler is not installed or `pdftocairo` / `pdftoppm` is not available.")

        try:
            completed = subprocess.run(
                command,
                capture_output=True,
                text=True,
                timeout=240,
            )
        except FileNotFoundError as exc:
            raise RasterizerError("The configured Poppler binary was not found.") from exc

        if completed.returncode != 0:
            stderr = completed.stderr.strip() or completed.stdout.strip() or "Unknown Poppler error."
            raise RasterizerError(f"Poppler rasterization failed: {stderr}")

        image_paths = self._collect_rendered_images(output_dir)
        if not image_paths:
            raise RasterizerError("Poppler reported success but did not write any PNG files.")
        return self.load_images(image_paths)

    def _resolve_poppler_binary(self, binary_name: str) -> Path | None:
        if self.poppler_bin_dir is not None:
            direct = self.poppler_bin_dir / binary_name
            if direct.exists():
                return direct
            exe_candidate = self.poppler_bin_dir / f"{binary_name}.exe"
            if exe_candidate.exists():
                return exe_candidate

        discovered = shutil.which(binary_name)
        if discovered:
            return Path(discovered)
        return None

    @staticmethod
    def _collect_rendered_images(output_dir: Path) -> list[Path]:
        image_paths = [
            path
            for path in output_dir.iterdir()
            if path.is_file() and path.suffix.lower() == ".png" and path.stat().st_size > 0
        ]
        return sorted(image_paths, key=Rasterizer._natural_path_sort_key)

    @staticmethod
    def _natural_path_sort_key(path: Path) -> list[object]:
        return [int(part) if part.isdigit() else part.lower() for part in re.split(r"(\d+)", path.name)]
