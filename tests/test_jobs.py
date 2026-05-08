from __future__ import annotations

from io import BytesIO
from pathlib import Path
import threading
import time

import numpy as np

from fastapi import UploadFile

from app.config import Settings
from app.services.jobs import JobManager
from app.services.models import PageImage, SlideQcResult, SlideQcStatus


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


def test_create_job_persists_prompt_override(tmp_path: Path) -> None:
    settings = Settings(
        base_dir=tmp_path,
        static_dir=tmp_path,
        runs_dir=tmp_path / "runs",
        platform_name="linux",
    )
    manager = JobManager(settings=settings)

    pdf_upload = UploadFile(BytesIO(b"%PDF-1.4"), filename="reference.pdf")
    pptx_upload = UploadFile(BytesIO(b"pptx"), filename="candidate.pptx")

    job = manager.create_job(
        pdf_upload=pdf_upload,
        pptx_upload=pptx_upload,
        qc_prompt_override="  Focus especially on superscripts and names.  ",
    )

    assert job.qc_prompt_override == "Focus especially on superscripts and names."


def test_create_job_persists_full_prompt_config(tmp_path: Path) -> None:
    settings = Settings(
        base_dir=tmp_path,
        static_dir=tmp_path,
        runs_dir=tmp_path / "runs",
        platform_name="linux",
    )
    manager = JobManager(settings=settings)

    pdf_upload = UploadFile(BytesIO(b"%PDF-1.4"), filename="reference.pdf")
    pptx_upload = UploadFile(BytesIO(b"pptx"), filename="candidate.pptx")

    job = manager.create_job(
        pdf_upload=pdf_upload,
        pptx_upload=pptx_upload,
        qc_prompt_config={
            "general_system_prompt": "visual system",
            "general_user_prompt": "visual user",
            "text_system_prompt": "text system",
            "text_user_prompt": "text user",
        },
    )

    assert job.qc_prompt_config == {
        "general_system_prompt": "visual system",
        "general_user_prompt": "visual user",
        "text_system_prompt": "text system",
        "text_user_prompt": "text user",
    }


def test_run_openai_qc_uses_fresh_evaluator_instances_in_parallel(tmp_path: Path) -> None:
    settings = Settings(
        base_dir=tmp_path,
        static_dir=tmp_path,
        runs_dir=tmp_path / "runs",
        platform_name="linux",
        openai_api_key="test-key",
        openai_qc_parallelism=4,
    )
    manager = JobManager(settings=settings)
    manager._update_job = lambda *args, **kwargs: None

    created_instance_ids: list[int] = []
    active_calls = 0
    max_active_calls = 0
    active_lock = threading.Lock()

    class FakeEvaluator:
        def __init__(self, instance_id: int):
            self.instance_id = instance_id

        def compare_pages(
            self,
            *,
            slide_index,
            page_index,
            reference_page,
            candidate_page,
            reference_layout=None,
            candidate_layout=None,
            debug_output_dir,
            prompt_override=None,
            prompt_config=None,
        ):
            nonlocal active_calls, max_active_calls
            with active_lock:
                active_calls += 1
                max_active_calls = max(max_active_calls, active_calls)
            time.sleep(0.05)
            with active_lock:
                active_calls -= 1
            return SlideQcResult(
                slide_index=slide_index,
                page_index=page_index,
                status=SlideQcStatus.OK,
                note=f"instance-{self.instance_id}",
            )

    def build_fake_evaluator():
        instance_id = len(created_instance_ids) + 1
        created_instance_ids.append(instance_id)
        return FakeEvaluator(instance_id)

    manager._build_openai_qc_evaluator = build_fake_evaluator

    def page(i: int) -> PageImage:
        image = np.full((10, 10, 3), 255, dtype=np.uint8)
        return PageImage(page_index=i, image=image, image_path=tmp_path / f"page-{i}.png")

    matched_pairs = [(i, page(i), page(i)) for i in range(4)]

    results = manager._run_openai_qc(
        "job-123",
        matched_pairs,
        reference_layouts=[],
        candidate_layouts=[],
        prompt_override="Focus on titles",
        prompt_config={"text_user_prompt": "Review lines one by one"},
    )

    assert len(created_instance_ids) == 4
    assert {result.note for result in results.values()} == {"instance-1", "instance-2", "instance-3", "instance-4"}
    assert max_active_calls > 1
