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


class JobStatusResponse(BaseModel):
    model_config = ConfigDict(populate_by_name=True)

    job_id: str = Field(alias="jobId")
    status: str
    step: str
    slide_progress: int = Field(alias="slideProgress")
    slide_count: int = Field(alias="slideCount")
    output_ready: bool = Field(alias="outputReady")
    pdf_page_count: int = Field(default=0, alias="pdfPageCount")
    pdf_page_character_totals: dict[int, int] = Field(default_factory=dict, alias="pdfPageCharacterTotals")
    pdf_fonts: list[PDFFontInfoResponse] = Field(default_factory=list, alias="pdfFonts")
    error: Optional[str] = None
