from __future__ import annotations

from typing import Optional

import fitz

from app.services.models import NormalizedBBox, PageImage, ParagraphLayout, TextBox, TextLayout, TextSource
from app.services.ocr_provider import NoopOcrProvider, OcrProvider


class NativeTextLayoutExtractor:
    def extract_document(self, pdf_path) -> list[TextLayout]:
        document = fitz.open(str(pdf_path))
        layouts = [self._extract_page(page, page_index + 1) for page_index, page in enumerate(document)]
        document.close()
        return layouts

    def _extract_page(self, page: fitz.Page, page_number: int) -> TextLayout:
        page_rect = page.rect
        page_width = float(page_rect.width or 1.0)
        page_height = float(page_rect.height or 1.0)
        text_data = page.get_text("dict")

        lines: list[TextBox] = []
        paragraphs: list[ParagraphLayout] = []
        for block in text_data.get("blocks", []):
            if block.get("type") != 0:
                continue

            paragraph_lines: list[TextBox] = []
            for line in block.get("lines", []):
                spans = line.get("spans", [])
                line_text = "".join(str(span.get("text") or "") for span in spans).strip()
                if not line_text:
                    continue

                line_box = TextBox(
                    text=line_text,
                    bbox=self._normalize_bbox(line.get("bbox"), page_width, page_height),
                    page_number=page_number,
                    confidence=1.0,
                    source=TextSource.NATIVE,
                )
                lines.append(line_box)
                paragraph_lines.append(line_box)

            if paragraph_lines:
                paragraph_text = " ".join(line.text for line in paragraph_lines if line.text.strip()).strip()
                if paragraph_text:
                    paragraphs.append(
                        ParagraphLayout(
                            text=paragraph_text,
                            lines=paragraph_lines,
                            bbox=self._union_bbox([line.bbox for line in paragraph_lines]),
                            page_number=page_number,
                            confidence=1.0,
                            source=TextSource.NATIVE,
                        )
                    )

        total_characters = sum(1 for paragraph in paragraphs for character in paragraph.text if not character.isspace())
        return TextLayout(
            page_number=page_number,
            page_size=(page_width, page_height),
            paragraphs=paragraphs,
            lines=lines,
            source=TextSource.NATIVE,
            total_characters=total_characters,
            average_confidence=1.0 if paragraphs else 0.0,
            extracted_with_ocr=False,
        )

    @staticmethod
    def _normalize_bbox(raw_bbox, page_width: float, page_height: float) -> NormalizedBBox:
        if not raw_bbox:
            return (0.0, 0.0, 1.0, 1.0)
        x0, y0, x1, y1 = raw_bbox
        return (
            NativeTextLayoutExtractor._clamp(float(x0) / page_width),
            NativeTextLayoutExtractor._clamp(float(y0) / page_height),
            NativeTextLayoutExtractor._clamp(float(x1) / page_width),
            NativeTextLayoutExtractor._clamp(float(y1) / page_height),
        )

    @staticmethod
    def _union_bbox(boxes: list[NormalizedBBox]) -> NormalizedBBox:
        x0 = min(box[0] for box in boxes)
        y0 = min(box[1] for box in boxes)
        x1 = max(box[2] for box in boxes)
        y1 = max(box[3] for box in boxes)
        return (x0, y0, x1, y1)

    @staticmethod
    def _clamp(value: float) -> float:
        return max(0.0, min(1.0, value))


class TextLayoutExtractor:
    def __init__(
        self,
        native_extractor: Optional[NativeTextLayoutExtractor] = None,
        ocr_provider: Optional[OcrProvider] = None,
        sparse_text_threshold: int = 24,
    ):
        self.native_extractor = native_extractor or NativeTextLayoutExtractor()
        self.ocr_provider = ocr_provider or NoopOcrProvider()
        self.sparse_text_threshold = sparse_text_threshold

    def extract_document(self, pdf_path, page_images: list[PageImage]) -> list[TextLayout]:
        native_layouts = self.native_extractor.extract_document(pdf_path)
        layouts: list[TextLayout] = []

        for page_index, page_image in enumerate(page_images):
            native_layout = native_layouts[page_index] if page_index < len(native_layouts) else self._empty_layout(page_image)
            if self._needs_ocr(native_layout):
                ocr_layout = self.ocr_provider.extract_page(page_image) if self.ocr_provider.is_available() else None
                if ocr_layout is not None and ocr_layout.total_characters >= native_layout.total_characters:
                    layouts.append(ocr_layout)
                    continue
            layouts.append(native_layout)

        return layouts

    def _needs_ocr(self, layout: TextLayout) -> bool:
        return layout.total_characters < self.sparse_text_threshold or not layout.paragraphs

    @staticmethod
    def _empty_layout(page_image: PageImage) -> TextLayout:
        return TextLayout(
            page_number=page_image.page_index + 1,
            page_size=(float(page_image.width), float(page_image.height)),
            source=TextSource.NATIVE,
            extracted_with_ocr=False,
        )
