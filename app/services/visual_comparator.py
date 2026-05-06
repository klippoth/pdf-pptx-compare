from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import cv2
import numpy as np

from app.services.models import PageAlignment, PageImage, VisualDiff, VisualDiffRegion


@dataclass
class _AlignmentCandidate:
    angle: int
    canvas: np.ndarray
    similarity_score: float
    offset: tuple[int, int]
    scaled_size: tuple[int, int]
    target_size: tuple[int, int]


class VisualComparator:
    def __init__(
        self,
        rotations: tuple[int, ...] = (0, 90, 270),
        content_threshold: int = 248,
        min_missing_region_area: int = 700,
        min_missing_area_ratio: float = 0.00085,
        min_geometry_region_area: int = 650,
        min_geometry_area_ratio: float = 0.00075,
        color_distance_threshold: float = 26.0,
        min_color_region_area: int = 500,
        min_color_area_ratio: float = 0.00055,
    ):
        self.rotations = rotations
        self.content_threshold = content_threshold
        self.min_missing_region_area = min_missing_region_area
        self.min_missing_area_ratio = min_missing_area_ratio
        self.min_geometry_region_area = min_geometry_region_area
        self.min_geometry_area_ratio = min_geometry_area_ratio
        self.color_distance_threshold = color_distance_threshold
        self.min_color_region_area = min_color_region_area
        self.min_color_area_ratio = min_color_area_ratio

    def compare(
        self,
        reference_page: PageImage,
        candidate_page: PageImage,
        missing_mask_path: Optional[Path] = None,
        geometry_mask_path: Optional[Path] = None,
        color_mask_path: Optional[Path] = None,
    ) -> VisualDiff:
        best = self._best_alignment(reference_page.image, candidate_page.image)
        alignment = PageAlignment(
            page_index=candidate_page.page_index,
            angle=best.angle,
            aligned_reference_image=best.canvas,
            similarity_score=best.similarity_score,
            target_size=best.target_size,
            offset=best.offset,
            scaled_size=best.scaled_size,
        )

        missing_mask = self._missing_mask(best.canvas, candidate_page.image)
        geometry_mask = self._geometry_mask(best.canvas, candidate_page.image)
        color_mask = self._color_mask(best.canvas, candidate_page.image)
        regions = self._extract_regions(
            missing_mask,
            min_region_area=self.min_missing_region_area,
            min_area_ratio=self.min_missing_area_ratio,
        )
        geometry_regions = self._extract_regions(
            geometry_mask,
            min_region_area=self.min_geometry_region_area,
            min_area_ratio=self.min_geometry_area_ratio,
        )
        color_regions = self._extract_regions(
            color_mask,
            min_region_area=self.min_color_region_area,
            min_area_ratio=self.min_color_area_ratio,
        )

        if missing_mask_path is not None:
            missing_mask_path.parent.mkdir(parents=True, exist_ok=True)
            cv2.imwrite(str(missing_mask_path), missing_mask)
        if geometry_mask_path is not None:
            geometry_mask_path.parent.mkdir(parents=True, exist_ok=True)
            cv2.imwrite(str(geometry_mask_path), geometry_mask)
        if color_mask_path is not None:
            color_mask_path.parent.mkdir(parents=True, exist_ok=True)
            cv2.imwrite(str(color_mask_path), color_mask)

        return VisualDiff(
            page_index=candidate_page.page_index,
            alignment=alignment,
            missing_regions=regions,
            geometry_regions=geometry_regions,
            color_regions=color_regions,
            missing_mask_path=missing_mask_path,
            geometry_mask_path=geometry_mask_path,
            color_mask_path=color_mask_path,
        )

    def _best_alignment(self, reference_image: np.ndarray, candidate_image: np.ndarray) -> _AlignmentCandidate:
        target_size = (candidate_image.shape[1], candidate_image.shape[0])
        candidate_bbox = self._content_bbox(candidate_image)
        best: Optional[_AlignmentCandidate] = None

        for angle in self.rotations:
            rotated = self._rotate_image(reference_image, angle)
            canvas, offset, scaled_size = self._fit_reference_to_canvas(rotated, target_size, candidate_bbox)
            similarity = self._similarity(canvas, candidate_image)
            alignment = _AlignmentCandidate(
                angle=angle,
                canvas=canvas,
                similarity_score=similarity,
                offset=offset,
                scaled_size=scaled_size,
                target_size=target_size,
            )
            if best is None or alignment.similarity_score > best.similarity_score:
                best = alignment

        if best is None:
            raise RuntimeError("Failed to align pages for QC.")
        return best

    def _fit_reference_to_canvas(
        self,
        reference_image: np.ndarray,
        target_size: tuple[int, int],
        candidate_bbox: tuple[int, int, int, int],
    ) -> tuple[np.ndarray, tuple[int, int], tuple[int, int]]:
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

        ref_bbox = self._content_bbox(resized)
        candidate_center = self._bbox_center(candidate_bbox)
        reference_center = self._bbox_center(ref_bbox)
        x = int(round(candidate_center[0] - reference_center[0]))
        y = int(round(candidate_center[1] - reference_center[1]))
        x = max(0, min(x, target_width - resized_width))
        y = max(0, min(y, target_height - resized_height))

        canvas[y : y + resized_height, x : x + resized_width] = resized
        return canvas, (x, y), (resized_width, resized_height)

    def _missing_mask(self, aligned_reference: np.ndarray, candidate_image: np.ndarray) -> np.ndarray:
        reference_gray = cv2.cvtColor(aligned_reference, cv2.COLOR_RGB2GRAY)
        candidate_gray = cv2.cvtColor(candidate_image, cv2.COLOR_RGB2GRAY)

        reference_mask, candidate_mask = self._content_masks(reference_gray, candidate_gray)

        missing = cv2.bitwise_and(reference_mask, cv2.bitwise_not(candidate_mask))
        kernel = np.ones((3, 3), dtype=np.uint8)
        missing = cv2.morphologyEx(missing, cv2.MORPH_OPEN, kernel)
        missing = cv2.morphologyEx(missing, cv2.MORPH_CLOSE, kernel)
        return missing

    def _geometry_mask(self, aligned_reference: np.ndarray, candidate_image: np.ndarray) -> np.ndarray:
        reference_gray = cv2.cvtColor(aligned_reference, cv2.COLOR_RGB2GRAY)
        candidate_gray = cv2.cvtColor(candidate_image, cv2.COLOR_RGB2GRAY)

        reference_mask, candidate_mask = self._content_masks(reference_gray, candidate_gray)
        diff_mask = cv2.bitwise_xor(reference_mask, candidate_mask)
        kernel = np.ones((3, 3), dtype=np.uint8)
        diff_mask = cv2.morphologyEx(diff_mask, cv2.MORPH_OPEN, kernel)
        diff_mask = cv2.morphologyEx(diff_mask, cv2.MORPH_CLOSE, kernel)
        diff_mask = cv2.dilate(diff_mask, kernel, iterations=1)
        return diff_mask

    def _color_mask(self, aligned_reference: np.ndarray, candidate_image: np.ndarray) -> np.ndarray:
        reference_gray = cv2.cvtColor(aligned_reference, cv2.COLOR_RGB2GRAY)
        candidate_gray = cv2.cvtColor(candidate_image, cv2.COLOR_RGB2GRAY)
        reference_mask, candidate_mask = self._content_masks(reference_gray, candidate_gray)
        overlap_mask = cv2.bitwise_and(reference_mask, candidate_mask)
        kernel = np.ones((3, 3), dtype=np.uint8)
        overlap_mask = cv2.erode(overlap_mask, kernel, iterations=1)

        reference_blur = cv2.GaussianBlur(aligned_reference, (5, 5), 0)
        candidate_blur = cv2.GaussianBlur(candidate_image, (5, 5), 0)
        reference_lab = cv2.cvtColor(reference_blur, cv2.COLOR_RGB2LAB).astype(np.float32)
        candidate_lab = cv2.cvtColor(candidate_blur, cv2.COLOR_RGB2LAB).astype(np.float32)
        color_distance = np.linalg.norm(reference_lab - candidate_lab, axis=2)
        diff_mask = ((color_distance >= self.color_distance_threshold).astype(np.uint8)) * 255
        diff_mask = cv2.bitwise_and(diff_mask, overlap_mask)
        diff_mask = cv2.morphologyEx(diff_mask, cv2.MORPH_OPEN, kernel)
        diff_mask = cv2.morphologyEx(diff_mask, cv2.MORPH_CLOSE, kernel)
        diff_mask = cv2.dilate(diff_mask, kernel, iterations=1)
        return diff_mask

    def _extract_regions(
        self,
        diff_mask: np.ndarray,
        *,
        min_region_area: int,
        min_area_ratio: float,
    ) -> list[VisualDiffRegion]:
        total_area = float(diff_mask.shape[0] * diff_mask.shape[1])
        region_count, labels, stats, _ = cv2.connectedComponentsWithStats(diff_mask, connectivity=8)
        regions: list[VisualDiffRegion] = []

        for region_index in range(1, region_count):
            x, y, width, height, area = stats[region_index]
            area_ratio = float(area) / total_area if total_area else 0.0
            if area < min_region_area or area_ratio < min_area_ratio:
                continue

            regions.append(
                VisualDiffRegion(
                    bbox=(
                        x / diff_mask.shape[1],
                        y / diff_mask.shape[0],
                        (x + width) / diff_mask.shape[1],
                        (y + height) / diff_mask.shape[0],
                    ),
                    area_ratio=area_ratio,
                )
            )

        return regions

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
        return float(np.sum(reference_gray * candidate_gray) / denominator)

    def _content_masks(self, reference_gray: np.ndarray, candidate_gray: np.ndarray) -> tuple[np.ndarray, np.ndarray]:
        reference_blur = cv2.GaussianBlur(reference_gray, (5, 5), 0)
        candidate_blur = cv2.GaussianBlur(candidate_gray, (5, 5), 0)
        reference_mask = (reference_blur < self.content_threshold).astype(np.uint8) * 255
        candidate_mask = (candidate_blur < self.content_threshold).astype(np.uint8) * 255
        return reference_mask, candidate_mask

    def _crop_to_content(self, image: np.ndarray) -> np.ndarray:
        gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
        mask = (cv2.GaussianBlur(gray, (5, 5), 0) < self.content_threshold).astype(np.uint8) * 255
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

    def _content_bbox(self, image: np.ndarray) -> tuple[int, int, int, int]:
        gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
        mask = (cv2.GaussianBlur(gray, (5, 5), 0) < self.content_threshold).astype(np.uint8) * 255
        coords = cv2.findNonZero(mask)
        if coords is None:
            return (0, 0, image.shape[1], image.shape[0])
        return cv2.boundingRect(coords)

    @staticmethod
    def _bbox_center(bbox: tuple[int, int, int, int]) -> tuple[float, float]:
        x, y, width, height = bbox
        return (x + (width / 2.0), y + (height / 2.0))

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
