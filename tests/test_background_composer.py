from __future__ import annotations

from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw

from app.services.background_composer import BackgroundComposer
from app.services.models import PageImage, PlacementStatus


def _make_slide_art(size: tuple[int, int] = (1200, 700)) -> np.ndarray:
    width, height = size
    image = Image.new("RGB", size, "white")
    draw = ImageDraw.Draw(image)
    draw.rectangle((80, 60, width - 80, 180), fill=(210, 210, 210))
    draw.rectangle((130, 270, width - 140, 330), fill=(40, 40, 40))
    draw.rectangle((150, 380, width - 360, 430), fill=(40, 40, 40))
    draw.rectangle((90, 266, 118, 294), fill=(220, 40, 40))
    draw.rectangle((90, 376, 118, 404), fill=(220, 40, 40))
    return np.array(image)


def _page(path: Path, image: np.ndarray, index: int) -> PageImage:
    Image.fromarray(image).save(path)
    return PageImage(page_index=index, image=image, image_path=path)


def test_background_composer_rotates_reference_to_match_slide(tmp_path: Path) -> None:
    composer = BackgroundComposer()
    candidate = _make_slide_art()
    reference = np.rot90(candidate, 1).copy()

    result = composer.prepare_background(
        reference_page=_page(tmp_path / "reference.png", reference, 0),
        candidate_page=_page(tmp_path / "candidate.png", candidate, 0),
        output_path=tmp_path / "background.png",
    )

    assert result.status == PlacementStatus.PLACED
    assert result.rotation_degrees == 270
    assert result.background_image_path.exists()
    with Image.open(result.background_image_path) as image:
        assert image.size == (candidate.shape[1], candidate.shape[0])


def test_background_composer_excludes_upside_down_rotation_and_keeps_reference_image_opaque(tmp_path: Path) -> None:
    composer = BackgroundComposer()
    candidate = _make_slide_art()
    reference = candidate.copy()

    result = composer.prepare_background(
        reference_page=_page(tmp_path / "reference.png", reference, 0),
        candidate_page=_page(tmp_path / "candidate.png", candidate, 0),
        output_path=tmp_path / "background.png",
    )

    assert 180 not in composer.rotations
    with Image.open(result.background_image_path) as image:
        assert image.mode == "RGB"
        output = np.array(image)

    top_left_region = output[:40, :40]
    red_pixels = (top_left_region[:, :, 0] > 200) & (top_left_region[:, :, 1] < 90) & (top_left_region[:, :, 2] < 90)
    assert red_pixels.sum() > 0
