from __future__ import annotations

from pathlib import Path

from pptx import Presentation

from app.services.annotation_writer import AnnotationWriter
from app.services.models import PagePlacementResult, PlacementBundle, PlacementStatus, QcReport


class DeckWriter:
    def __init__(self):
        self.pdf_reference_shape_name = "PDF_ORIGINAL"
        self.pdf_reference_slide_name = "PDF_ORIGINAL"
        self.annotation_writer = AnnotationWriter()

    def build_output(
        self,
        source_pptx: Path,
        placement_bundle: PlacementBundle,
        output_path: Path,
        qc_report: QcReport | None = None,
    ) -> Path:
        presentation = Presentation(str(source_pptx))
        original_slides = list(presentation.slides)
        reference_insertions: list[tuple[int, PagePlacementResult]] = []

        for slide_index, slide in enumerate(original_slides):
            if slide_index >= len(placement_bundle.slide_results):
                break
            result = placement_bundle.slide_results[slide_index]
            if qc_report is not None and slide_index < len(qc_report.slide_results):
                self.annotation_writer.apply(
                    slide=slide,
                    slide_qc_result=qc_report.slide_results[slide_index],
                    slide_width=presentation.slide_width,
                    slide_height=presentation.slide_height,
                )
            if result.background_image_path and result.status == PlacementStatus.PLACED:
                reference_insertions.append((slide_index, result))

        for slide_index, result in reversed(reference_insertions):
            self._insert_pdf_reference_slide_after(
                presentation=presentation,
                insert_after_index=slide_index,
                result=result,
            )

        for result in placement_bundle.extra_reference_results:
            self._append_pdf_only_slide(
                presentation=presentation,
                result=result,
            )

        output_path.parent.mkdir(parents=True, exist_ok=True)
        presentation.save(str(output_path))
        return output_path

    def _append_pdf_only_slide(
        self,
        presentation: Presentation,
        result: PagePlacementResult,
    ) -> None:
        slide = self._append_blank_slide(presentation)
        self._set_slide_name(slide, self.pdf_reference_slide_name)
        if result.background_image_path:
            self._add_full_slide_picture(
                slide=slide,
                image_path=result.background_image_path,
                slide_width=presentation.slide_width,
                slide_height=presentation.slide_height,
            )

    def _insert_pdf_reference_slide_after(
        self,
        presentation: Presentation,
        insert_after_index: int,
        result: PagePlacementResult,
    ) -> None:
        slide = self._append_blank_slide(presentation)
        self._set_slide_name(slide, self.pdf_reference_slide_name)
        if result.background_image_path:
            self._add_full_slide_picture(
                slide=slide,
                image_path=result.background_image_path,
                slide_width=presentation.slide_width,
                slide_height=presentation.slide_height,
            )
        self._move_last_slide_to_index(presentation, insert_after_index + 1)

    def _append_blank_slide(self, presentation: Presentation):
        layout = presentation.slide_layouts[6] if len(presentation.slide_layouts) > 6 else presentation.slide_layouts[-1]
        return presentation.slides.add_slide(layout)

    def _move_last_slide_to_index(self, presentation: Presentation, insert_index: int) -> None:
        slide_id_list = presentation.slides._sldIdLst
        slide_id = slide_id_list[-1]
        slide_id_list.remove(slide_id)
        slide_id_list.insert(insert_index, slide_id)

    def _add_full_slide_picture(self, slide, image_path: Path, slide_width, slide_height) -> None:
        picture = slide.shapes.add_picture(str(image_path), 0, 0, width=slide_width, height=slide_height)
        picture.name = self.pdf_reference_shape_name

    def _set_slide_name(self, slide, name: str) -> None:
        slide._element.cSld.set("name", name)
