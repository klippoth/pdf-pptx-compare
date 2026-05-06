from __future__ import annotations

from io import BytesIO
from pathlib import Path

from fastapi import UploadFile

from app.config import Settings
from app.services.jobs import JobManager


def test_create_job_persists_ai_qc_toggle(tmp_path: Path) -> None:
    settings = Settings(
        base_dir=tmp_path,
        static_dir=tmp_path,
        runs_dir=tmp_path / "runs",
        platform_name="linux",
    )
    manager = JobManager(settings=settings)

    pdf_upload = UploadFile(BytesIO(b"%PDF-1.4"), filename="reference.pdf")
    pptx_upload = UploadFile(BytesIO(b"pptx"), filename="candidate.pptx")

    job = manager.create_job(pdf_upload=pdf_upload, pptx_upload=pptx_upload, enable_ai_qc=False)

    assert job.enable_ai_qc is False
    assert job.input_pdf_path.read_bytes() == b"%PDF-1.4"
    assert job.input_pptx_path.read_bytes() == b"pptx"
