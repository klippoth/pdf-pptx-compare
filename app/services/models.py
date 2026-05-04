from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
from enum import Enum
from pathlib import Path
from typing import Optional

import numpy as np


class JobState(str, Enum):
    QUEUED = "queued"
    PROCESSING = "processing"
    COMPLETED = "completed"
    FAILED = "failed"


class PlacementStatus(str, Enum):
    PLACED = "placed"
    NO_MATCHING_PDF_PAGE = "no_matching_pdf_page"
    EXTRA_PDF_PAGE = "extra_pdf_page"


@dataclass
class PageImage:
    page_index: int
    image: np.ndarray
    image_path: Path

    @property
    def width(self) -> int:
        return int(self.image.shape[1])

    @property
    def height(self) -> int:
        return int(self.image.shape[0])


@dataclass
class PagePlacementResult:
    candidate_slide_index: Optional[int]
    reference_page_index: Optional[int]
    status: PlacementStatus
    background_image_path: Optional[Path]
    message: str
    rotation_degrees: int = 0
    similarity_score: float = 0.0


@dataclass
class PlacementBundle:
    slide_results: list[PagePlacementResult] = field(default_factory=list)
    extra_reference_results: list[PagePlacementResult] = field(default_factory=list)


@dataclass
class PDFFontInfo:
    name: str
    font_type: str
    embedded: bool
    subset: bool
    page_numbers: list[int] = field(default_factory=list)
    page_character_counts: dict[int, int] = field(default_factory=dict)


@dataclass
class PDFFontInspectionResult:
    fonts: list[PDFFontInfo] = field(default_factory=list)
    page_character_totals: dict[int, int] = field(default_factory=dict)
    page_count: int = 0


@dataclass
class JobRecord:
    job_id: str
    working_dir: Path
    input_pdf_path: Path
    input_pptx_path: Path
    original_pptx_name: str
    status: str = JobState.QUEUED.value
    step: str = "Queued for processing"
    slide_progress: int = 0
    slide_count: int = 0
    output_pptx_path: Optional[Path] = None
    pdf_font_report_path: Optional[Path] = None
    pdf_fonts: list[PDFFontInfo] = field(default_factory=list)
    pdf_page_character_totals: dict[int, int] = field(default_factory=dict)
    pdf_page_count: int = 0
    error: Optional[str] = None
    created_at: datetime = field(default_factory=lambda: datetime.now(timezone.utc))
