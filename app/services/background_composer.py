from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import cv2
import numpy as np
from PIL import Image, ImageDraw

from app.services.models import PageImage, PagePlacementResult, PlacementStatus


@dataclass
class _PlacementCandidate:
    angle: int
    canvas: np.ndarray
    similarity_score: float


class BackgroundComposer:
    def __init__(self, rotations: tuple[int, ...] = (0, 90, 270)):
        self.rotations = rotations

    def prepare_background(
        self,
        reference_page: PageImage,
        candidate_page: PageImage,
        output_path: Path,
    ) -> PagePlacementResult:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        best = self._best_candidate(reference_page.image, candidate_page.image)
        self._add_reference_border(Image.fromarray(best.canvas)).save(output_path)

        return PagePlacementResult(
            candidate_slide_index=candidate_page.page_index,
            reference_page_index=reference_page.page_index,
            status=PlacementStatus.PLACED,
            background_image_path=output_path,
            message="Parked PDF page off the top-right of the slide as a movable reference image",
            rotation_degrees=best.angle,
            similarity_score=best.similarity_score,
        )

    def prepare_extra_reference_page(
        self,
        reference_page: PageImage,
        target_size: tuple[int, int],
        output_path: Path,
    ) -> PagePlacementResult:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        best = self._best_candidate(reference_page.image, None, target_size=target_size)
        Image.fromarray(best.canvas).save(output_path)

        return PagePlacementResult(
            candidate_slide_index=None,
            reference_page_index=reference_page.page_index,
            status=PlacementStatus.EXTRA_PDF_PAGE,
            background_image_path=output_path,
            message="Appended unmatched PDF page as a new reference slide",
            rotation_degrees=best.angle,
            similarity_score=best.similarity_score,
        )

    def no_matching_pdf_page(self, candidate_slide_index: int) -> PagePlacementResult:
        return PagePlacementResult(
            candidate_slide_index=candidate_slide_index,
            reference_page_index=None,
            status=PlacementStatus.NO_MATCHING_PDF_PAGE,
            background_image_path=None,
            message="No matching PDF page for this slide",
        )

    def _best_candidate(
        self,
        reference_image: np.ndarray,
        candidate_image: Optional[np.ndarray],
        target_size: Optional[tuple[int, int]] = None,
    ) -> _PlacementCandidate:
        if candidate_image is None and target_size is None:
            raise ValueError("Either a candidate image or an explicit target size is required.")

        if candidate_image is not None:
            height, width = candidate_image.shape[:2]
            target_size = (width, height)
        else:
            width, height = target_size

        best: Optional[_PlacementCandidate] = None
        candidate_bbox = self._content_bbox(candidate_image) if candidate_image is not None else None

        for angle in self.rotations:
            rotated = self._rotate_image(reference_image, angle)
            canvas = self._fit_reference_to_canvas(rotated, target_size, candidate_bbox)
            similarity = self._similarity(canvas, candidate_image) if candidate_image is not None else self._aspect_score(rotated, target_size)
            placement = _PlacementCandidate(angle=angle, canvas=canvas, similarity_score=similarity)
            if best is None or placement.similarity_score > best.similarity_score:
                best = placement

        if best is None:
            raise RuntimeError("Failed to generate a PDF background image.")
        return best

    def _fit_reference_to_canvas(
        self,
        reference_image: np.ndarray,
        target_size: tuple[int, int],
        candidate_bbox: Optional[tuple[int, int, int, int]],
    ) -> np.ndarray:
        target_width, target_height = target_size
        ref_height, ref_width = reference_image.shape[:2]
        scale = min(target_width / ref_width, target_height / ref_height)
        resized_width = max(1, int(round(ref_width * scale)))
        resized_height = max(1, int(round(ref_height * scale)))
        interpolation = cv2.INTER_CUBIC if scale > 1 else cv2.INTER_AREA
        resized = cv2.resize(reference_image, (resized_width, resized_height), interpolation=interpolation)

        canvas = np.full((target_height, target_width, 3), 255, dtype=np.uint8)
        x = (target_width - resized_width) // 2
        y = (target_height - resized_height) // 2

        if candidate_bbox is not None:
            ref_bbox = self._content_bbox(resized)
            candidate_center = self._bbox_center(candidate_bbox)
            reference_center = self._bbox_center(ref_bbox)
            x = int(round(candidate_center[0] - reference_center[0]))
            y = int(round(candidate_center[1] - reference_center[1]))
            x = max(0, min(x, target_width - resized_width))
            y = max(0, min(y, target_height - resized_height))

        canvas[y : y + resized_height, x : x + resized_width] = resized
        return canvas

    def _similarity(self, reference_canvas: np.ndarray, candidate_image: np.ndarray) -> float:
        reference_crop = self._crop_to_content(reference_canvas)
        candidate_crop = self._crop_to_content(candidate_image)

        common_width = 720
        common_height = 405
        reference_gray = cv2.cvtColor(
            cv2.resize(reference_crop, (common_width, common_height), interpolation=cv2.INTER_AREA),
            cv2.COLOR_RGB2GRAY,
        ).astype(np.float32)
        candidate_gray = cv2.cvtColor(
            cv2.resize(candidate_crop, (common_width, common_height), interpolation=cv2.INTER_AREA),
            cv2.COLOR_RGB2GRAY,
        ).astype(np.float32)

        reference_gray -= reference_gray.mean()
        candidate_gray -= candidate_gray.mean()
        denominator = np.linalg.norm(reference_gray) * np.linalg.norm(candidate_gray)
        if denominator == 0:
            return -1.0

        correlation = float(np.sum(reference_gray * candidate_gray) / denominator)
        aspect_bonus = self._aspect_score(reference_crop, (candidate_crop.shape[1], candidate_crop.shape[0]))
        return correlation + (0.05 * aspect_bonus)

    def _aspect_score(self, image: np.ndarray, target_size: tuple[int, int]) -> float:
        target_width, target_height = target_size
        target_ratio = target_width / max(target_height, 1)
        image_ratio = image.shape[1] / max(image.shape[0], 1)
        return -abs(target_ratio - image_ratio)

    def _crop_to_content(self, image: np.ndarray) -> np.ndarray:
        gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
        mask = (cv2.GaussianBlur(gray, (5, 5), 0) < 248).astype(np.uint8) * 255
        coords = cv2.findNonZero(mask)
        if coords is None:
            return image
        x, y, w, h = cv2.boundingRect(coords)
        padding = 8
        x0 = max(0, x - padding)
        y0 = max(0, y - padding)
        x1 = min(image.shape[1], x + w + padding)
        y1 = min(image.shape[0], y + h + padding)
        return image[y0:y1, x0:x1]

    @staticmethod
    def _add_reference_border(image: Image.Image) -> Image.Image:
        bordered = image.copy()
        draw = ImageDraw.Draw(bordered)
        width, height = bordered.size
        thickness = max(6, int(round(min(width, height) * 0.012)))
        inset = max(1, thickness // 2)
        draw.rectangle(
            (inset, inset, width - inset - 1, height - inset - 1),
            outline=(220, 40, 40),
            width=thickness,
        )
        return bordered

    def _content_bbox(self, image: Optional[np.ndarray]) -> tuple[int, int, int, int]:
        if image is None:
            return (0, 0, 0, 0)
        gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
        mask = (cv2.GaussianBlur(gray, (5, 5), 0) < 248).astype(np.uint8) * 255
        coords = cv2.findNonZero(mask)
        if coords is None:
            return (0, 0, image.shape[1], image.shape[0])
        return cv2.boundingRect(coords)

    @staticmethod
    def _bbox_center(bbox: tuple[int, int, int, int]) -> tuple[float, float]:
        x, y, w, h = bbox
        return (x + (w / 2.0), y + (h / 2.0))

    @staticmethod
    def _rotate_image(image: np.ndarray, angle: int) -> np.ndarray:
        angle = angle % 360
        if angle == 0:
            return image.copy()
        if angle == 90:
            return np.rot90(image, 1).copy()
        if angle == 180:
            return np.rot90(image, 2).copy()
        if angle == 270:
            return np.rot90(image, 3).copy()
        raise ValueError(f"Unsupported rotation angle: {angle}")
