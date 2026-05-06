from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Optional, Protocol

from app.services.models import NormalizedBBox, PageImage, ParagraphLayout, TextBox, TextLayout, TextSource


class OcrProvider(Protocol):
    def is_available(self) -> bool:
        """Return whether the OCR provider is configured and importable."""

    def extract_page(self, page_image: PageImage) -> Optional[TextLayout]:
        """Extract structured text from a page image."""


class NoopOcrProvider:
    def is_available(self) -> bool:
        return False

    def extract_page(self, page_image: PageImage) -> Optional[TextLayout]:
        return None


@dataclass
class GoogleDocumentAiOcrProvider:
    project_id: Optional[str]
    location: Optional[str]
    processor_id: Optional[str]
    processor_version: Optional[str] = None

    def __post_init__(self) -> None:
        self._client: Any = None
        self._documentai: Any = None
        self._client_options_cls: Any = None

    def is_available(self) -> bool:
        return bool(self.project_id and self.location and self.processor_id and self._load_sdk())

    def extract_page(self, page_image: PageImage) -> Optional[TextLayout]:
        if not self.is_available():
            return None

        documentai = self._documentai
        client = self._get_client()
        processor_name = self._processor_name(client)
        raw_document = documentai.RawDocument(
            content=page_image.image_path.read_bytes(),
            mime_type="image/png",
        )
        request = documentai.ProcessRequest(
            name=processor_name,
            raw_document=raw_document,
        )
        result = client.process_document(request=request)
        return self._build_text_layout(result.document, page_image)

    def _load_sdk(self) -> bool:
        if self._documentai is not None and self._client_options_cls is not None:
            return True

        try:
            from google.api_core.client_options import ClientOptions
            from google.cloud import documentai
        except ImportError:
            return False

        self._documentai = documentai
        self._client_options_cls = ClientOptions
        return True

    def _get_client(self) -> Any:
        if self._client is None:
            endpoint = f"{self.location}-documentai.googleapis.com"
            self._client = self._documentai.DocumentProcessorServiceClient(
                client_options=self._client_options_cls(api_endpoint=endpoint)
            )
        return self._client

    def _processor_name(self, client: Any) -> str:
        if self.processor_version:
            return client.processor_version_path(
                self.project_id,
                self.location,
                self.processor_id,
                self.processor_version,
            )
        return client.processor_path(
            self.project_id,
            self.location,
            self.processor_id,
        )

    def _build_text_layout(self, document: Any, page_image: PageImage) -> TextLayout:
        if not getattr(document, "pages", None):
            return TextLayout(
                page_number=page_image.page_index + 1,
                page_size=(float(page_image.width), float(page_image.height)),
                source=TextSource.OCR,
                extracted_with_ocr=True,
            )

        page = document.pages[0]
        document_text = getattr(document, "text", "")
        page_number = page_image.page_index + 1
        page_width = float(getattr(getattr(page, "dimension", None), "width", page_image.width) or page_image.width)
        page_height = float(getattr(getattr(page, "dimension", None), "height", page_image.height) or page_image.height)

        line_boxes = [
            self._build_text_box(line.layout, document_text, page_number, page_width, page_height)
            for line in getattr(page, "lines", [])
        ]
        line_boxes = [line for line in line_boxes if line and line.text.strip()]

        paragraph_layouts = [
            self._build_paragraph(paragraph.layout, document_text, page_number, page_width, page_height, line_boxes)
            for paragraph in getattr(page, "paragraphs", [])
        ]
        paragraph_layouts = [paragraph for paragraph in paragraph_layouts if paragraph and paragraph.text.strip()]

        if not paragraph_layouts:
            paragraph_layouts = [
                ParagraphLayout(
                    text=line.text,
                    lines=[line],
                    bbox=line.bbox,
                    page_number=page_number,
                    confidence=line.confidence,
                    source=TextSource.OCR,
                )
                for line in line_boxes
            ]

        total_characters = sum(1 for paragraph in paragraph_layouts for character in paragraph.text if not character.isspace())
        confidence_values = [paragraph.confidence for paragraph in paragraph_layouts] or [0.0]
        return TextLayout(
            page_number=page_number,
            page_size=(page_width, page_height),
            paragraphs=paragraph_layouts,
            lines=line_boxes,
            source=TextSource.OCR,
            total_characters=total_characters,
            average_confidence=sum(confidence_values) / len(confidence_values),
            extracted_with_ocr=True,
        )

    def _build_text_box(
        self,
        layout: Any,
        document_text: str,
        page_number: int,
        page_width: float,
        page_height: float,
    ) -> Optional[TextBox]:
        text = self._anchor_text(document_text, getattr(layout, "text_anchor", None)).strip()
        if not text:
            return None
        return TextBox(
            text=text,
            bbox=self._normalized_bbox(getattr(layout, "bounding_poly", None), page_width, page_height),
            page_number=page_number,
            confidence=float(getattr(layout, "confidence", 0.85) or 0.85),
            source=TextSource.OCR,
        )

    def _build_paragraph(
        self,
        layout: Any,
        document_text: str,
        page_number: int,
        page_width: float,
        page_height: float,
        line_boxes: list[TextBox],
    ) -> Optional[ParagraphLayout]:
        text = self._anchor_text(document_text, getattr(layout, "text_anchor", None)).strip()
        if not text:
            return None
        bbox = self._normalized_bbox(getattr(layout, "bounding_poly", None), page_width, page_height)
        lines = [line for line in line_boxes if self._bbox_overlap_ratio(bbox, line.bbox) >= 0.55]
        return ParagraphLayout(
            text=text,
            lines=lines,
            bbox=bbox,
            page_number=page_number,
            confidence=float(getattr(layout, "confidence", 0.85) or 0.85),
            source=TextSource.OCR,
        )

    @staticmethod
    def _anchor_text(document_text: str, text_anchor: Any) -> str:
        if not text_anchor:
            return ""
        segments: list[str] = []
        for segment in getattr(text_anchor, "text_segments", []):
            start_index = int(getattr(segment, "start_index", 0) or 0)
            end_index = int(getattr(segment, "end_index", 0) or 0)
            segments.append(document_text[start_index:end_index])
        return "".join(segments)

    @staticmethod
    def _normalized_bbox(bounding_poly: Any, page_width: float, page_height: float) -> NormalizedBBox:
        if not bounding_poly:
            return (0.0, 0.0, 1.0, 1.0)

        normalized_vertices = list(getattr(bounding_poly, "normalized_vertices", []) or [])
        if normalized_vertices:
            xs = [float(vertex.x) for vertex in normalized_vertices]
            ys = [float(vertex.y) for vertex in normalized_vertices]
            return (
                GoogleDocumentAiOcrProvider._clamp(min(xs)),
                GoogleDocumentAiOcrProvider._clamp(min(ys)),
                GoogleDocumentAiOcrProvider._clamp(max(xs)),
                GoogleDocumentAiOcrProvider._clamp(max(ys)),
            )

        vertices = list(getattr(bounding_poly, "vertices", []) or [])
        if not vertices:
            return (0.0, 0.0, 1.0, 1.0)

        xs = [float(vertex.x) / max(page_width, 1.0) for vertex in vertices]
        ys = [float(vertex.y) / max(page_height, 1.0) for vertex in vertices]
        return (
            GoogleDocumentAiOcrProvider._clamp(min(xs)),
            GoogleDocumentAiOcrProvider._clamp(min(ys)),
            GoogleDocumentAiOcrProvider._clamp(max(xs)),
            GoogleDocumentAiOcrProvider._clamp(max(ys)),
        )

    @staticmethod
    def _bbox_overlap_ratio(left: NormalizedBBox, right: NormalizedBBox) -> float:
        left_x0, left_y0, left_x1, left_y1 = left
        right_x0, right_y0, right_x1, right_y1 = right
        overlap_x0 = max(left_x0, right_x0)
        overlap_y0 = max(left_y0, right_y0)
        overlap_x1 = min(left_x1, right_x1)
        overlap_y1 = min(left_y1, right_y1)
        overlap_width = max(0.0, overlap_x1 - overlap_x0)
        overlap_height = max(0.0, overlap_y1 - overlap_y0)
        overlap_area = overlap_width * overlap_height
        left_area = max(0.0, left_x1 - left_x0) * max(0.0, left_y1 - left_y0)
        if left_area <= 0:
            return 0.0
        return overlap_area / left_area

    @staticmethod
    def _clamp(value: float) -> float:
        return max(0.0, min(1.0, value))
