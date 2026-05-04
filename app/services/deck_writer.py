from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.util import Emu

from app.services.models import PagePlacementResult, PlacementBundle, PlacementStatus


class DeckWriter:
    def __init__(
        self,
        parked_width_ratio: float = 0.42,
        visible_peek_ratio: float = 0.18,
        top_margin_ratio: float = 0.02,
    ):
        self.parked_width_ratio = parked_width_ratio
        self.visible_peek_ratio = visible_peek_ratio
        self.top_margin_ratio = top_margin_ratio
        self.sticky_shape_name = "Sticky"

    def build_output(self, source_pptx: Path, placement_bundle: PlacementBundle, output_path: Path) -> Path:
        presentation = Presentation(str(source_pptx))

        for slide_index, slide in enumerate(presentation.slides):
            if slide_index >= len(placement_bundle.slide_results):
                break
            result = placement_bundle.slide_results[slide_index]
            if result.background_image_path and result.status == PlacementStatus.PLACED:
                self._add_parked_reference_picture(
                    slide=slide,
                    image_path=result.background_image_path,
                    slide_width=presentation.slide_width,
                    slide_height=presentation.slide_height,
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
        layout = presentation.slide_layouts[6] if len(presentation.slide_layouts) > 6 else presentation.slide_layouts[-1]
        slide = presentation.slides.add_slide(layout)
        if result.background_image_path:
            self._add_full_slide_picture(
                slide=slide,
                image_path=result.background_image_path,
                slide_width=presentation.slide_width,
                slide_height=presentation.slide_height,
            )

    def _add_parked_reference_picture(self, slide, image_path: Path, slide_width: Emu, slide_height: Emu) -> None:
        parked_width = max(1, int(round(slide_width * self.parked_width_ratio)))
        parked_height = max(1, int(round(slide_height * self.parked_width_ratio)))
        visible_peek = max(1, int(round(slide_width * self.visible_peek_ratio)))
        left = int(slide_width - visible_peek)
        top = int(round(slide_height * self.top_margin_ratio))
        picture = slide.shapes.add_picture(str(image_path), left, top, width=parked_width, height=parked_height)
        picture.name = self.sticky_shape_name

    def _add_full_slide_picture(self, slide, image_path: Path, slide_width: Emu, slide_height: Emu) -> None:
        picture = slide.shapes.add_picture(str(image_path), 0, 0, width=slide_width, height=slide_height)
        picture.name = self.sticky_shape_name
