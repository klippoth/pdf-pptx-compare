from __future__ import annotations

from typing import Optional

from pydantic import BaseModel, ConfigDict, Field


class PDFFontInfoResponse(BaseModel):
    name: str
    font_type: str = Field(alias="fontType")
    embedded: bool
    subset: bool
    page_numbers: list[int] = Field(alias="pageNumbers")
    page_character_counts: dict[int, int] = Field(default_factory=dict, alias="pageCharacterCounts")


class CompareResponse(BaseModel):
    model_config = ConfigDict(populate_by_name=True)

    job_id: str = Field(alias="jobId")
    status: str
    ai_qc_enabled: bool = Field(default=True, alias="aiQcEnabled")


class RendererStatusResponse(BaseModel):
    model_config = ConfigDict(populate_by_name=True)

    can_convert: bool = Field(alias="canConvert")
    preferred_renderer: str = Field(alias="preferredRenderer")
    libreoffice_available: bool = Field(alias="libreofficeAvailable")
    powerpoint_available: bool = Field(alias="powerpointAvailable")
    can_export_slide_images: bool = Field(default=False, alias="canExportSlideImages")
    slide_image_export_renderer: str = Field(default="none", alias="slideImageExportRenderer")
    slide_image_export_message: str = Field(default="", alias="slideImageExportMessage")
    message: str


class HealthResponse(BaseModel):
    model_config = ConfigDict(populate_by_name=True)

    status: str
    renderer: RendererStatusResponse


class JobStatusResponse(BaseModel):
    model_config = ConfigDict(populate_by_name=True)

    job_id: str = Field(alias="jobId")
    status: str
    step: str
    slide_progress: int = Field(alias="slideProgress")
    slide_count: int = Field(alias="slideCount")
    output_ready: bool = Field(alias="outputReady")
    ai_qc_enabled: bool = Field(default=True, alias="aiQcEnabled")
    pdf_page_count: int = Field(default=0, alias="pdfPageCount")
    pdf_page_character_totals: dict[int, int] = Field(default_factory=dict, alias="pdfPageCharacterTotals")
    pdf_fonts: list[PDFFontInfoResponse] = Field(default_factory=list, alias="pdfFonts")
    qc_counts_by_type: dict[str, int] = Field(default_factory=dict, alias="qcCountsByType")
    qc_manual_review_count: int = Field(default=0, alias="qcManualReviewCount")
    error: Optional[str] = None
