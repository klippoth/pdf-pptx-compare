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


class TextSource(str, Enum):
    NATIVE = "native"
    OCR = "ocr"
    MODEL = "model"


class SlideQcStatus(str, Enum):
    OK = "ok"
    FINDINGS = "findings"
    MANUAL_REVIEW = "manual_review"


class QcFindingType(str, Enum):
    MISSING_CONTENT = "missing_content"
    EXTRA_CONTENT = "extra_content"
    WRONG_TEXT = "wrong_text"
    LINE_BREAK_ISSUE = "line_break_issue"
    ALIGNMENT_DRIFT = "alignment_drift"
    SIZE_POSITION_ISSUE = "size_position_issue"
    WRONG_COLOR = "wrong_color"


class QcFindingSeverity(str, Enum):
    HIGH = "high"
    MEDIUM = "medium"
    LOW = "low"


NormalizedBBox = tuple[float, float, float, float]


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
class TextBox:
    text: str
    bbox: NormalizedBBox
    page_number: int
    confidence: float
    source: TextSource


@dataclass
class ParagraphLayout:
    text: str
    lines: list[TextBox]
    bbox: NormalizedBBox
    page_number: int
    confidence: float
    source: TextSource


@dataclass
class TextLayout:
    page_number: int
    page_size: tuple[float, float]
    paragraphs: list[ParagraphLayout] = field(default_factory=list)
    lines: list[TextBox] = field(default_factory=list)
    source: TextSource = TextSource.NATIVE
    total_characters: int = 0
    average_confidence: float = 0.0
    extracted_with_ocr: bool = False


@dataclass
class PageAlignment:
    page_index: int
    angle: int
    aligned_reference_image: np.ndarray
    similarity_score: float
    target_size: tuple[int, int]
    offset: tuple[int, int]
    scaled_size: tuple[int, int]


@dataclass
class VisualDiffRegion:
    bbox: NormalizedBBox
    area_ratio: float


@dataclass
class VisualDiff:
    page_index: int
    alignment: PageAlignment
    missing_regions: list[VisualDiffRegion] = field(default_factory=list)
    geometry_regions: list[VisualDiffRegion] = field(default_factory=list)
    color_regions: list[VisualDiffRegion] = field(default_factory=list)
    missing_mask_path: Optional[Path] = None
    geometry_mask_path: Optional[Path] = None
    color_mask_path: Optional[Path] = None


@dataclass
class SlideQcFinding:
    finding_id: int
    finding_type: QcFindingType
    severity: QcFindingSeverity
    bbox: NormalizedBBox
    message: str
    confidence: float


@dataclass
class SlideQcResult:
    slide_index: int
    page_index: int
    status: SlideQcStatus
    findings: list[SlideQcFinding] = field(default_factory=list)
    alignment_confidence: float = 0.0
    reference_source: TextSource = TextSource.NATIVE
    candidate_source: TextSource = TextSource.NATIVE
    summary: Optional[str] = None
    comment_bullets: list[str] = field(default_factory=list)
    note: Optional[str] = None


@dataclass
class QcReport:
    slide_results: list[SlideQcResult] = field(default_factory=list)
    counts_by_type: dict[str, int] = field(default_factory=dict)
    manual_review_count: int = 0


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
    enable_ai_qc: bool = True
    qc_prompt_override: Optional[str] = None
    qc_prompt_config: dict[str, str] = field(default_factory=dict)
    status: str = JobState.QUEUED.value
    step: str = "Queued for processing"
    slide_progress: int = 0
    slide_count: int = 0
    output_pptx_path: Optional[Path] = None
    pdf_page_count: int = 0
    qc_report_path: Optional[Path] = None
    qc_counts_by_type: dict[str, int] = field(default_factory=dict)
    qc_manual_review_count: int = 0
    error: Optional[str] = None
    created_at: datetime = field(default_factory=lambda: datetime.now(timezone.utc))
