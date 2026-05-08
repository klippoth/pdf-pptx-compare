from __future__ import annotations

from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from app.config import get_settings
from app.schemas import (
    CompareResponse,
    HealthResponse,
    JobStatusResponse,
    PDFFontInfoResponse,
    QcPromptConfigResponse,
    RendererStatusResponse,
)
from app.services.jobs import JobManager
from app.services.openai_qc import OpenAIQCEvaluator


settings = get_settings()
MAX_QC_PROMPT_OVERRIDE_CHARS = 4000
MAX_QC_PROMPT_FIELD_CHARS = 20000


@asynccontextmanager
async def lifespan(app: FastAPI):
    job_manager = JobManager(settings=settings)
    job_manager.start()
    app.state.job_manager = job_manager
    try:
        yield
    finally:
        job_manager.stop()


app = FastAPI(title="PDF-to-PPTX Redline", lifespan=lifespan)
app.mount("/static", StaticFiles(directory=settings.static_dir), name="static")


@app.get("/", include_in_schema=False)
async def index() -> FileResponse:
    return FileResponse(settings.static_dir / "index.html")


@app.get("/api/health", response_model=HealthResponse)
async def health(request: Request) -> HealthResponse:
    availability = request.app.state.job_manager.renderer.availability()
    return HealthResponse(
        status="ok",
        renderer=RendererStatusResponse(
            canConvert=availability.can_convert,
            preferredRenderer=availability.preferred_renderer,
            libreofficeAvailable=availability.libreoffice_available,
            powerpointAvailable=availability.powerpoint_available,
            canExportSlideImages=availability.can_export_slide_images,
            slideImageExportRenderer=availability.slide_image_export_renderer,
            slideImageExportMessage=availability.slide_image_export_message,
            message=availability.message,
        ),
    )


@app.get("/api/qc-prompts", response_model=QcPromptConfigResponse)
async def get_qc_prompts() -> QcPromptConfigResponse:
    prompts = OpenAIQCEvaluator.get_default_prompt_config()
    return QcPromptConfigResponse(
        generalSystemPrompt=prompts["general_system_prompt"],
        generalUserPrompt=prompts["general_user_prompt"],
        textSystemPrompt=prompts["text_system_prompt"],
        textUserPrompt=prompts["text_user_prompt"],
    )


@app.post("/api/compare", response_model=CompareResponse, status_code=202)
async def compare(
    request: Request,
    pdf: UploadFile = File(...),
    pptx: UploadFile = File(...),
    enable_ai_qc: bool = Form(True),
    qc_prompt_override: str = Form(""),
    qc_general_system_prompt: str = Form(""),
    qc_general_user_prompt: str = Form(""),
    qc_text_system_prompt: str = Form(""),
    qc_text_user_prompt: str = Form(""),
) -> CompareResponse:
    if Path(pdf.filename or "").suffix.lower() != ".pdf":
        raise HTTPException(status_code=400, detail="The `pdf` upload must be a .pdf file.")
    if Path(pptx.filename or "").suffix.lower() != ".pptx":
        raise HTTPException(status_code=400, detail="The `pptx` upload must be a .pptx file.")

    prompt_override = qc_prompt_override.strip()
    if len(prompt_override) > MAX_QC_PROMPT_OVERRIDE_CHARS:
        raise HTTPException(
            status_code=400,
            detail=f"The AI QC prompt override must be {MAX_QC_PROMPT_OVERRIDE_CHARS} characters or fewer.",
        )

    prompt_config = {
        "general_system_prompt": qc_general_system_prompt.strip(),
        "general_user_prompt": qc_general_user_prompt.strip(),
        "text_system_prompt": qc_text_system_prompt.strip(),
        "text_user_prompt": qc_text_user_prompt.strip(),
    }
    for field_name, field_value in prompt_config.items():
        if len(field_value) > MAX_QC_PROMPT_FIELD_CHARS:
            raise HTTPException(
                status_code=400,
                detail=f"The AI QC prompt field `{field_name}` must be {MAX_QC_PROMPT_FIELD_CHARS} characters or fewer.",
            )

    job = request.app.state.job_manager.create_job(
        pdf_upload=pdf,
        pptx_upload=pptx,
        enable_ai_qc=enable_ai_qc,
        qc_prompt_override=prompt_override or None,
        qc_prompt_config=prompt_config,
    )
    return CompareResponse(jobId=job.job_id, status=job.status, aiQcEnabled=job.enable_ai_qc)


@app.get("/api/jobs/{job_id}", response_model=JobStatusResponse)
async def get_job(job_id: str, request: Request) -> JobStatusResponse:
    job = request.app.state.job_manager.get_job(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="Job not found.")

    return JobStatusResponse(
        jobId=job.job_id,
        status=job.status,
        step=job.step,
        slideProgress=job.slide_progress,
        slideCount=job.slide_count,
        outputReady=bool(job.output_pptx_path and job.output_pptx_path.exists()),
        aiQcEnabled=job.enable_ai_qc,
        pdfPageCount=job.pdf_page_count,
        pdfPageCharacterTotals=job.pdf_page_character_totals,
        pdfFonts=[
            PDFFontInfoResponse(
                name=font.name,
                fontType=font.font_type,
                embedded=font.embedded,
                subset=font.subset,
                pageNumbers=font.page_numbers,
                pageCharacterCounts=font.page_character_counts,
            )
            for font in job.pdf_fonts
        ],
        qcCountsByType=job.qc_counts_by_type,
        qcManualReviewCount=job.qc_manual_review_count,
        error=job.error,
    )


@app.get("/api/jobs/{job_id}/download")
async def download(job_id: str, request: Request) -> FileResponse:
    job = request.app.state.job_manager.get_job(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="Job not found.")
    if not job.output_pptx_path or not job.output_pptx_path.exists():
        raise HTTPException(status_code=409, detail="Job output is not ready yet.")

    return FileResponse(
        path=job.output_pptx_path,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=job.output_pptx_path.name,
    )
