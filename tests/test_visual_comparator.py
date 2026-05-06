from __future__ import annotations

from pathlib import Path

import cv2
import numpy as np

from app.services.models import PageImage
from app.services.visual_comparator import VisualComparator


def _page(image: np.ndarray, page_index: int = 0) -> PageImage:
    return PageImage(page_index=page_index, image=image, image_path=Path(f"/tmp/page-{page_index}.png"))


def test_visual_comparator_detects_missing_regions() -> None:
    reference = np.full((300, 400, 3), 255, dtype=np.uint8)
    candidate = reference.copy()
    cv2.rectangle(reference, (40, 40), (180, 120), (0, 0, 0), thickness=-1)
    cv2.rectangle(reference, (220, 160), (340, 240), (0, 0, 0), thickness=-1)
    cv2.rectangle(candidate, (40, 40), (180, 120), (0, 0, 0), thickness=-1)

    result = VisualComparator().compare(_page(reference), _page(candidate))

    assert result.missing_regions


def test_visual_comparator_detects_geometry_differences() -> None:
    reference = np.full((300, 400, 3), 255, dtype=np.uint8)
    candidate = reference.copy()
    cv2.putText(reference, "Alpha Beta", (40, 120), cv2.FONT_HERSHEY_SIMPLEX, 1.1, (0, 0, 0), 2, cv2.LINE_AA)
    cv2.putText(candidate, "Alpha", (40, 120), cv2.FONT_HERSHEY_SIMPLEX, 1.1, (0, 0, 0), 2, cv2.LINE_AA)
    cv2.putText(candidate, "Beta", (40, 165), cv2.FONT_HERSHEY_SIMPLEX, 1.1, (0, 0, 0), 2, cv2.LINE_AA)

    result = VisualComparator().compare(_page(reference), _page(candidate))

    assert result.geometry_regions


def test_visual_comparator_detects_color_differences() -> None:
    reference = np.full((300, 400, 3), 255, dtype=np.uint8)
    candidate = reference.copy()
    cv2.rectangle(reference, (80, 80), (220, 180), (20, 70, 230), thickness=-1)
    cv2.rectangle(candidate, (80, 80), (220, 180), (20, 180, 40), thickness=-1)

    result = VisualComparator().compare(_page(reference), _page(candidate))

    assert result.color_regions
