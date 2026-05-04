from __future__ import annotations

from pathlib import Path

import fitz
import numpy as np
from PIL import Image

from app.services.models import PageImage


class Rasterizer:
    def __init__(self, dpi: int = 180):
        self.dpi = dpi

    def render_pdf(self, pdf_path: Path, output_dir: Path) -> list[PageImage]:
        output_dir.mkdir(parents=True, exist_ok=True)
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
