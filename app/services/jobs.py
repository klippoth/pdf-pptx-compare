from __future__ import annotations

import logging
import secrets
import shutil
import threading
from datetime import datetime, timedelta, timezone
from pathlib import Path
from queue import Queue
from typing import Optional

from fastapi import UploadFile

from app.config import Settings
from app.services.background_composer import BackgroundComposer
from app.services.deck_writer import DeckWriter
from app.services.models import JobRecord, JobState, PlacementBundle
from app.services.pdf_font_inspector import PDFFontInspector
from app.services.rasterizer import Rasterizer
from app.services.renderer import build_renderer


logger = logging.getLogger(__name__)


class JobManager:
    def __init__(self, settings: Settings):
        self.settings = settings
        self.settings.runs_dir.mkdir(parents=True, exist_ok=True)
        self.renderer = build_renderer(settings)
        self.rasterizer = Rasterizer(dpi=settings.render_dpi)
        self.background_composer = BackgroundComposer()
        self.pdf_font_inspector = PDFFontInspector()
        self.deck_writer = DeckWriter()

        self._jobs: dict[str, JobRecord] = {}
        self._queue: Queue[Optional[str]] = Queue()
        self._lock = threading.Lock()
        self._worker_thread: Optional[threading.Thread] = None
        self._stop_event = threading.Event()

    def start(self) -> None:
        self._cleanup_expired_runs()
        if self._worker_thread and self._worker_thread.is_alive():
            return
        self._worker_thread = threading.Thread(target=self._worker_loop, name="pdf-background-worker", daemon=True)
        self._worker_thread.start()

    def stop(self) -> None:
        self._stop_event.set()
        self._queue.put(None)
        if self._worker_thread and self._worker_thread.is_alive():
            self._worker_thread.join(timeout=2)

    def create_job(self, pdf_upload: UploadFile, pptx_upload: UploadFile) -> JobRecord:
        timestamp = datetime.now(timezone.utc).strftime("%Y%m%d%H%M%S")
        job_id = f"{timestamp}-{secrets.token_hex(4)}"
        working_dir = self.settings.runs_dir / job_id
        input_dir = working_dir / "input"
        render_dir = working_dir / "render"
        placed_dir = working_dir / "placed"
        output_dir = working_dir / "output"
        for directory in (input_dir, render_dir, placed_dir, output_dir):
            directory.mkdir(parents=True, exist_ok=True)

        pdf_path = input_dir / "reference.pdf"
        pptx_path = input_dir / "candidate.pptx"
        self._save_upload(pdf_upload, pdf_path)
        self._save_upload(pptx_upload, pptx_path)

        record = JobRecord(
            job_id=job_id,
            working_dir=working_dir,
            input_pdf_path=pdf_path,
            input_pptx_path=pptx_path,
            original_pptx_name=Path(pptx_upload.filename or "output.pptx").name,
        )
        with self._lock:
            self._jobs[job_id] = record

        self._queue.put(job_id)
        return record

    def get_job(self, job_id: str) -> Optional[JobRecord]:
        with self._lock:
            return self._jobs.get(job_id)

    def _worker_loop(self) -> None:
        while not self._stop_event.is_set():
            job_id = self._queue.get()
            if job_id is None:
                return
            try:
                self._process_job(job_id)
            except Exception:  # pragma: no cover - final guardrail
                logger.exception("Unhandled job-processing failure", extra={"job_id": job_id})
                self._update_job(
                    job_id,
                    status=JobState.FAILED.value,
                    step="Failed",
                    error="The job failed unexpectedly. Check the server log for details.",
                )

    def _process_job(self, job_id: str) -> None:
        job = self.get_job(job_id)
        if job is None:
            return

        render_dir = job.working_dir / "render"
        placed_dir = job.working_dir / "placed"
        output_dir = job.working_dir / "output"

        try:
            self._update_job(job_id, status=JobState.PROCESSING.value, step="Inspecting PDF fonts")
            font_inspection = self.pdf_font_inspector.inspect(job.input_pdf_path)
            pdf_font_report_path = self.pdf_font_inspector.write_report(font_inspection, output_dir / "pdf_fonts.json")
            self._update_job(
                job_id,
                pdf_fonts=font_inspection.fonts,
                pdf_page_character_totals=font_inspection.page_character_totals,
                pdf_page_count=font_inspection.page_count,
                pdf_font_report_path=pdf_font_report_path,
            )

            self._update_job(job_id, status=JobState.PROCESSING.value, step="Exporting PowerPoint to PDF")
            candidate_pdf = self.renderer.export_pptx_to_pdf(job.input_pptx_path, render_dir / "candidate.pdf")

            self._update_job(
                job_id,
                step=f"Rendering PowerPoint via {self.renderer.last_used_renderer}",
            )
            candidate_pages = self.rasterizer.render_pdf(candidate_pdf, render_dir / "candidate-pages")

            self._update_job(job_id, step="Rendering reference PDF")
            reference_pages = self.rasterizer.render_pdf(job.input_pdf_path, render_dir / "reference-pages")

            total_pages = max(len(reference_pages), len(candidate_pages))
            self._update_job(job_id, step="Preparing PDF page reference images", slide_count=total_pages, slide_progress=0)

            bundle = PlacementBundle()
            fallback_target_size = (
                (candidate_pages[0].width, candidate_pages[0].height)
                if candidate_pages
                else (reference_pages[0].width, reference_pages[0].height)
            )
            for index in range(total_pages):
                if index < len(reference_pages) and index < len(candidate_pages):
                    result = self.background_composer.prepare_background(
                        reference_page=reference_pages[index],
                        candidate_page=candidate_pages[index],
                        output_path=placed_dir / f"background-{index + 1:03d}.png",
                    )
                    bundle.slide_results.append(result)
                elif index < len(candidate_pages):
                    bundle.slide_results.append(self.background_composer.no_matching_pdf_page(candidate_slide_index=index))
                else:
                    bundle.extra_reference_results.append(
                        self.background_composer.prepare_extra_reference_page(
                            reference_page=reference_pages[index],
                            target_size=fallback_target_size,
                            output_path=placed_dir / f"extra-{index + 1:03d}.png",
                        )
                    )
                self._update_job(job_id, slide_progress=index + 1)

            self._update_job(job_id, step="Building PowerPoint with parked PDF references")
            output_name = f"{Path(job.original_pptx_name).stem}_with_pdf_pages.pptx"
            output_path = output_dir / output_name
            self.deck_writer.build_output(job.input_pptx_path, bundle, output_path)

            self._update_job(
                job_id,
                status=JobState.COMPLETED.value,
                step="Complete",
                output_pptx_path=output_path,
            )
        except Exception as exc:
            logger.exception("Job failed", extra={"job_id": job_id})
            self._update_job(
                job_id,
                status=JobState.FAILED.value,
                step="Failed",
                error=str(exc),
            )

    def _update_job(self, job_id: str, **changes) -> None:
        with self._lock:
            job = self._jobs[job_id]
            for field_name, value in changes.items():
                setattr(job, field_name, value)

    def _save_upload(self, upload: UploadFile, destination: Path) -> None:
        destination.parent.mkdir(parents=True, exist_ok=True)
        upload.file.seek(0)
        with destination.open("wb") as handle:
            shutil.copyfileobj(upload.file, handle)

    def _cleanup_expired_runs(self) -> None:
        cutoff = datetime.now(timezone.utc) - timedelta(hours=self.settings.cleanup_after_hours)
        for run_dir in self.settings.runs_dir.iterdir():
            if not run_dir.is_dir():
                continue
            modified_at = datetime.fromtimestamp(run_dir.stat().st_mtime, tz=timezone.utc)
            if modified_at < cutoff:
                shutil.rmtree(run_dir, ignore_errors=True)
