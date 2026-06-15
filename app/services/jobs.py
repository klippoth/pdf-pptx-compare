from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor, as_completed
import logging
import os
import secrets
import shutil
import subprocess
import threading
import zipfile
from datetime import datetime, timedelta, timezone
from pathlib import Path
from queue import Queue
from typing import Optional
import xml.etree.ElementTree as ET

from fastapi import UploadFile

from app.config import Settings
from app.services.background_composer import BackgroundComposer
from app.services.deck_writer import DeckWriter
from app.services.models import (
    PageImage,
    JobRecord,
    JobState,
    PlacementBundle,
    QcFindingSeverity,
    QcFindingType,
    QcReport,
    SlideQcFinding,
    SlideQcResult,
    SlideQcStatus,
)
from app.services.ocr_provider import GoogleDocumentAiOcrProvider
from app.services.openai_qc import OpenAIQCEvaluator
from app.services.pptx_text_extractor import PptxTextExtractor
from app.services.qc_detector import QcDetector
from app.services.qc_report_writer import QcReportWriter
from app.services.rasterizer import Rasterizer
from app.services.renderer import build_renderer
from app.services.text_layout_extractor import TextLayoutExtractor
from app.services.visual_comparator import VisualComparator


logger = logging.getLogger(__name__)
EMU_PER_INCH = 914400.0
PRESENTATION_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"


class JobManager:
    def __init__(self, settings: Settings):
        self.settings = settings
        self.settings.runs_dir.mkdir(parents=True, exist_ok=True)
        self.settings.downloads_dir.mkdir(parents=True, exist_ok=True)
        self.renderer = build_renderer(settings)
        self.rasterizer = Rasterizer(dpi=settings.render_dpi, poppler_bin_dir=settings.poppler_bin_dir)
        self.background_composer = BackgroundComposer()
        self.pptx_text_extractor = PptxTextExtractor()
        self.text_layout_extractor = TextLayoutExtractor(
            ocr_provider=GoogleDocumentAiOcrProvider(
                project_id=settings.google_document_ai_project_id,
                location=settings.google_document_ai_location,
                processor_id=settings.google_document_ai_processor_id,
                processor_version=settings.google_document_ai_processor_version,
            )
        )
        self.visual_comparator = VisualComparator()
        self.qc_detector = QcDetector()
        self.qc_report_writer = QcReportWriter()
        self.deck_writer = DeckWriter()
        self.openai_qc_evaluator = self._build_openai_qc_evaluator()

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

    def create_job(
        self,
        pdf_upload: UploadFile,
        pptx_upload: UploadFile,
        enable_ai_qc: bool = False,
        qc_prompt_override: Optional[str] = None,
        qc_prompt_config: Optional[dict[str, str]] = None,
    ) -> JobRecord:
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
            enable_ai_qc=enable_ai_qc,
            qc_prompt_override=(qc_prompt_override or "").strip() or None,
            qc_prompt_config=dict(qc_prompt_config or {}),
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
        qc_dir = job.working_dir / "qc"

        try:
            self._update_job(job_id, status=JobState.PROCESSING.value, step="Rendering reference PDF")

            ai_qc_enabled = job.enable_ai_qc
            use_model_qc = ai_qc_enabled and self.openai_qc_evaluator.is_available()
            use_reference_only_insertion = not ai_qc_enabled
            use_direct_slide_images = self.renderer.can_export_slide_images() and use_model_qc
            comparison_export_renderer = None
            comparison_export_fallback = True
            comparison_raster_engine = "fitz"
            if use_model_qc and not use_direct_slide_images:
                if self.settings.prefer_powerpoint_for_ai_qc and self.renderer.is_powerpoint_available():
                    comparison_export_renderer = "powerpoint"
                    comparison_export_fallback = False
                if self.settings.prefer_poppler_for_ai_qc:
                    comparison_raster_engine = "auto"

            reference_pages = self.rasterizer.render_pdf(
                job.input_pdf_path,
                render_dir / "reference-pages",
                engine=comparison_raster_engine,
            )
            self._update_job(
                job_id,
                step=f"Rendered reference pages via {self.rasterizer.last_used_engine}",
                pdf_page_count=len(reference_pages),
            )

            candidate_pdf: Optional[Path] = None
            candidate_pdf_pages: list[PageImage] = []
            comparison_candidate_pages: list[PageImage] = []
            reference_layouts = []
            candidate_layouts = []
            bundle = PlacementBundle()
            slide_qc_results_by_index: dict[int, SlideQcResult] = {}
            matched_pairs: list[tuple[int, PageImage, PageImage]] = []

            if use_reference_only_insertion:
                self._update_job(job_id, step="Preparing PDF reference slides without candidate rendering")
                bundle, slide_qc_results_by_index, total_pages = self._build_reference_only_placement_bundle(
                    job.input_pptx_path,
                    reference_pages,
                    placed_dir,
                )
                self._update_job(job_id, step="Preparing slide references", slide_count=total_pages, slide_progress=0)
            elif use_direct_slide_images:
                self._update_job(
                    job_id,
                    step=(
                        "Exporting PowerPoint slides as images for AI QC"
                        if ai_qc_enabled
                        else "Exporting PowerPoint slides as images for PDF insertion"
                    ),
                )
                candidate_slide_image_paths = self.renderer.export_pptx_to_slide_images(
                    job.input_pptx_path,
                    render_dir / "candidate-slide-images",
                    dpi=self.settings.render_dpi,
                )
                comparison_candidate_pages = self.rasterizer.load_images(candidate_slide_image_paths)
                self._update_job(
                    job_id,
                    step=(
                        "Prepared PowerPoint slide images for AI comparison"
                        if ai_qc_enabled
                        else "Prepared PowerPoint slide images for PDF insertion"
                    ),
                )
            else:
                export_step = "Exporting PowerPoint to PDF"
                if use_model_qc and comparison_export_renderer == "powerpoint":
                    export_step = "Exporting PowerPoint to PDF via PowerPoint"
                self._update_job(job_id, status=JobState.PROCESSING.value, step=export_step)
                candidate_pdf = self.renderer.export_pptx_to_pdf(
                    job.input_pptx_path,
                    render_dir / "candidate.pdf",
                    preferred_renderer=comparison_export_renderer,
                    allow_fallback=comparison_export_fallback,
                )

                candidate_pdf_pages = self.rasterizer.render_pdf(
                    candidate_pdf,
                    render_dir / "candidate-pages",
                    engine=comparison_raster_engine,
                )
                comparison_candidate_pages = candidate_pdf_pages
                self._update_job(
                    job_id,
                    step=f"Rendered candidate slides via {self.renderer.last_used_renderer} + {self.rasterizer.last_used_engine}",
                )

            if not use_reference_only_insertion:
                if use_model_qc:
                    self._update_job(job_id, step="Preparing candidate slide images for AI comparison")
                total_pages = max(len(reference_pages), len(comparison_candidate_pages))
                self._update_job(job_id, step="Preparing slide references", slide_count=total_pages, slide_progress=0)

                fallback_target_size = (
                    (comparison_candidate_pages[0].width, comparison_candidate_pages[0].height)
                    if comparison_candidate_pages
                    else (reference_pages[0].width, reference_pages[0].height)
                )
                for index in range(total_pages):
                    if index < len(reference_pages) and index < len(comparison_candidate_pages):
                        result = self.background_composer.prepare_background(
                            reference_page=reference_pages[index],
                            candidate_page=comparison_candidate_pages[index],
                            output_path=placed_dir / f"background-{index + 1:03d}.png",
                        )
                        bundle.slide_results.append(result)
                        matched_pairs.append((index, reference_pages[index], comparison_candidate_pages[index]))
                    elif index < len(comparison_candidate_pages):
                        bundle.slide_results.append(self.background_composer.no_matching_pdf_page(candidate_slide_index=index))
                        slide_qc_results_by_index[index] = self._missing_reference_qc_result(index)
                    else:
                        bundle.extra_reference_results.append(
                            self.background_composer.prepare_extra_reference_page(
                                reference_page=reference_pages[index],
                                target_size=fallback_target_size,
                                output_path=placed_dir / f"extra-{index + 1:03d}.png",
                            )
                        )

            qc_report: QcReport | None = None
            if use_model_qc:
                self._update_job(job_id, step="Extracting PowerPoint text for AI QC")
                candidate_layouts = self.pptx_text_extractor.extract_document(job.input_pptx_path)
                self._update_job(job_id, step="Running GPT slide QC", slide_progress=0)
                slide_qc_results_by_index.update(
                    self._run_openai_qc(
                        job_id,
                        matched_pairs,
                        candidate_layouts=candidate_layouts,
                        prompt_override=job.qc_prompt_override,
                        prompt_config=job.qc_prompt_config,
                    )
                )
            elif ai_qc_enabled:
                self._update_job(job_id, step="Extracting page text and layout")
                reference_layouts = self.text_layout_extractor.extract_document(job.input_pdf_path, reference_pages)
                candidate_layouts = self.text_layout_extractor.extract_document(candidate_pdf, candidate_pdf_pages)
                self._update_job(job_id, step="Running rendered-slide QC", slide_progress=0)
                slide_qc_results_by_index.update(
                    self._run_rule_based_qc(
                        job_id=job_id,
                        matched_pairs=matched_pairs,
                        reference_layouts=reference_layouts,
                        candidate_layouts=candidate_layouts,
                        qc_dir=qc_dir,
                    )
                )
            else:
                self._update_job(job_id, step="AI QC disabled; inserting PDF reference slides only", slide_progress=total_pages)

            self._update_job(job_id, slide_progress=total_pages)
            if ai_qc_enabled:
                slide_qc_results = [slide_qc_results_by_index[index] for index in sorted(slide_qc_results_by_index)]
                qc_report = self._build_qc_report(slide_qc_results)
                qc_report_path = self.qc_report_writer.write(qc_report, output_dir / "qc_report.json")
                self._update_job(
                    job_id,
                    qc_report_path=qc_report_path,
                    qc_counts_by_type=qc_report.counts_by_type,
                    qc_manual_review_count=qc_report.manual_review_count,
                )
            else:
                self._update_job(
                    job_id,
                    qc_report_path=None,
                    qc_counts_by_type={},
                    qc_manual_review_count=0,
                )

            self._update_job(
                job_id,
                step="Annotating PowerPoint with QC findings" if ai_qc_enabled else "Building output deck",
            )
            output_name = f"{Path(job.original_pptx_name).stem}_with_pdf_pages.pptx"
            output_path = output_dir / output_name
            try:
                if use_reference_only_insertion and self.deck_writer.can_build_reference_only_output_by_package_patch():
                    self.deck_writer.build_reference_only_output_by_package_patch(
                        job.input_pptx_path,
                        bundle,
                        output_path,
                    )
                else:
                    self.deck_writer.build_output(job.input_pptx_path, bundle, output_path, qc_report=qc_report)
            except Exception as exc:
                if self._should_retry_with_powerpoint_native_output(ai_qc_enabled=ai_qc_enabled, error=exc):
                    self._update_job(
                        job_id,
                        step="Corrupted embedded PPT media detected; retrying output assembly in PowerPoint",
                    )
                    self.renderer.build_output_with_reference_slides(job.input_pptx_path, bundle, output_path)
                else:
                    raise

            self._update_job(job_id, step="Publishing output deck to Downloads")
            published_output_path = self._publish_output_pptx(output_path, output_name)
            self._update_job(job_id, step="Opening output deck")
            self._open_output_pptx(published_output_path)

            self._update_job(
                job_id,
                status=JobState.COMPLETED.value,
                step="Complete",
                output_pptx_path=published_output_path,
            )
        except Exception as exc:
            logger.exception("Job failed", extra={"job_id": job_id})
            self._update_job(
                job_id,
                status=JobState.FAILED.value,
                step="Failed",
                error=str(exc),
            )

    def _build_reference_only_placement_bundle(
        self,
        source_pptx_path: Path,
        reference_pages: list[PageImage],
        placed_dir: Path,
    ) -> tuple[PlacementBundle, dict[int, SlideQcResult], int]:
        slide_count, target_size = self._read_pptx_slide_metrics(source_pptx_path)
        total_pages = max(len(reference_pages), slide_count)
        bundle = PlacementBundle()
        slide_qc_results_by_index: dict[int, SlideQcResult] = {}

        for index in range(total_pages):
            if index < len(reference_pages) and index < slide_count:
                bundle.slide_results.append(
                    self.background_composer.prepare_reference_page(
                        reference_page=reference_pages[index],
                        target_size=target_size,
                        output_path=placed_dir / f"background-{index + 1:03d}.png",
                    )
                )
            elif index < slide_count:
                bundle.slide_results.append(self.background_composer.no_matching_pdf_page(candidate_slide_index=index))
                slide_qc_results_by_index[index] = self._missing_reference_qc_result(index)
            else:
                bundle.extra_reference_results.append(
                    self.background_composer.prepare_extra_reference_page(
                        reference_page=reference_pages[index],
                        target_size=target_size,
                        output_path=placed_dir / f"extra-{index + 1:03d}.png",
                    )
                )

        return bundle, slide_qc_results_by_index, total_pages

    def _read_pptx_slide_metrics(self, pptx_path: Path) -> tuple[int, tuple[int, int]]:
        with zipfile.ZipFile(pptx_path, "r") as archive:
            presentation_root = ET.fromstring(archive.read("ppt/presentation.xml"))

        slide_id_list = presentation_root.find(f"{{{PRESENTATION_NS}}}sldIdLst")
        slide_count = len(list(slide_id_list)) if slide_id_list is not None else 0

        slide_size = presentation_root.find(f"{{{PRESENTATION_NS}}}sldSz")
        if slide_size is None:
            raise RuntimeError("The PowerPoint package is missing slide size metadata.")

        slide_width_emu = int(slide_size.get("cx", "0"))
        slide_height_emu = int(slide_size.get("cy", "0"))
        slide_width = max(1, int(round((float(slide_width_emu) / EMU_PER_INCH) * self.settings.render_dpi)))
        slide_height = max(1, int(round((float(slide_height_emu) / EMU_PER_INCH) * self.settings.render_dpi)))
        return slide_count, (slide_width, slide_height)

    def _run_openai_qc(
        self,
        job_id: str,
        matched_pairs: list[tuple[int, PageImage, PageImage]],
        *,
        candidate_layouts,
        prompt_override: Optional[str] = None,
        prompt_config: Optional[dict[str, str]] = None,
    ) -> dict[int, SlideQcResult]:
        results: dict[int, SlideQcResult] = {}
        processed = 0
        max_workers = min(max(1, self.settings.openai_qc_parallelism), max(1, len(matched_pairs)))
        job_qc_inputs_dir = self.settings.runs_dir / job_id / "qc" / "model-inputs"
        with ThreadPoolExecutor(max_workers=max_workers, thread_name_prefix="openai-slide-qc") as executor:
            future_map = {
                executor.submit(
                    self._compare_slide_with_fresh_openai_evaluator,
                    index=index,
                    reference_page=reference_page,
                    candidate_page=candidate_page,
                    candidate_layout=candidate_layouts[index] if index < len(candidate_layouts) else None,
                    debug_output_dir=job_qc_inputs_dir / f"slide-{index + 1:03d}",
                    prompt_override=prompt_override,
                    prompt_config=prompt_config,
                ): index
                for index, reference_page, candidate_page in matched_pairs
            }
            for future in as_completed(future_map):
                index = future_map[future]
                try:
                    results[index] = future.result()
                except Exception as exc:  # pragma: no cover - network/runtime guardrail
                    logger.exception("OpenAI slide QC failed", extra={"job_id": job_id, "slide_index": index})
                    results[index] = SlideQcResult(
                        slide_index=index,
                        page_index=index,
                        status=SlideQcStatus.MANUAL_REVIEW,
                        alignment_confidence=0.0,
                        note=f"GPT slide comparison failed for this page: {exc}",
                    )
                processed += 1
                self._update_job(job_id, slide_progress=processed)
        return results

    def _build_openai_qc_evaluator(self) -> OpenAIQCEvaluator:
        return OpenAIQCEvaluator(
            api_key=self.settings.openai_api_key,
            model=self.settings.openai_qc_model,
            timeout_seconds=self.settings.openai_qc_timeout_seconds,
            max_image_dimension=self.settings.openai_qc_max_image_dimension,
        )

    def _compare_slide_with_fresh_openai_evaluator(
        self,
        *,
        index: int,
        reference_page: PageImage,
        candidate_page: PageImage,
        candidate_layout=None,
        debug_output_dir: Path,
        prompt_override: Optional[str] = None,
        prompt_config: Optional[dict[str, str]] = None,
    ) -> SlideQcResult:
        evaluator = self._build_openai_qc_evaluator()
        return evaluator.compare_pages(
            slide_index=index,
            page_index=index,
            reference_page=reference_page,
            candidate_page=candidate_page,
            candidate_layout=candidate_layout,
            debug_output_dir=debug_output_dir,
            prompt_override=prompt_override,
            prompt_config=prompt_config,
        )

    def _run_rule_based_qc(
        self,
        *,
        job_id: str,
        matched_pairs: list[tuple[int, PageImage, PageImage]],
        reference_layouts,
        candidate_layouts,
        qc_dir: Path,
    ) -> dict[int, SlideQcResult]:
        results: dict[int, SlideQcResult] = {}
        processed = 0
        for index, reference_page, candidate_page in matched_pairs:
            visual_diff = self.visual_comparator.compare(
                reference_page=reference_page,
                candidate_page=candidate_page,
                missing_mask_path=qc_dir / f"missing-mask-{index + 1:03d}.png",
                geometry_mask_path=qc_dir / f"geometry-mask-{index + 1:03d}.png",
                color_mask_path=qc_dir / f"color-mask-{index + 1:03d}.png",
            )
            reference_layout = (
                reference_layouts[index]
                if index < len(reference_layouts)
                else self.text_layout_extractor._empty_layout(reference_page)
            )
            candidate_layout = (
                candidate_layouts[index]
                if index < len(candidate_layouts)
                else self.text_layout_extractor._empty_layout(candidate_page)
            )
            results[index] = self.qc_detector.detect(
                reference_layout=reference_layout,
                candidate_layout=candidate_layout,
                visual_diff=visual_diff,
            )
            processed += 1
            self._update_job(job_id, slide_progress=processed)
        return results

    def _build_qc_report(self, slide_qc_results: list[SlideQcResult]) -> QcReport:
        counts_by_type: dict[str, int] = {}
        manual_review_count = 0
        for slide_result in slide_qc_results:
            if slide_result.status == SlideQcStatus.MANUAL_REVIEW:
                manual_review_count += 1
            for finding in slide_result.findings:
                key = finding.finding_type.value
                counts_by_type[key] = counts_by_type.get(key, 0) + 1
        return QcReport(
            slide_results=slide_qc_results,
            counts_by_type=counts_by_type,
            manual_review_count=manual_review_count,
        )

    @staticmethod
    def _missing_reference_qc_result(slide_index: int) -> SlideQcResult:
        return SlideQcResult(
            slide_index=slide_index,
            page_index=slide_index,
            status=SlideQcStatus.MANUAL_REVIEW,
            findings=[
                SlideQcFinding(
                    finding_id=1,
                    finding_type=QcFindingType.MISSING_CONTENT,
                    severity=QcFindingSeverity.HIGH,
                    bbox=(0.03, 0.03, 0.97, 0.97),
                    message="No matching PDF page exists for this slide.",
                    confidence=1.0,
                )
            ],
            alignment_confidence=0.0,
            note="No matching PDF page exists for this slide.",
        )

    def _should_retry_with_powerpoint_native_output(self, *, ai_qc_enabled: bool, error: Exception) -> bool:
        if ai_qc_enabled:
            return False
        if not self.renderer.can_build_output_with_reference_slides():
            return False
        if isinstance(error, zipfile.BadZipFile):
            return True
        message = str(error)
        if "Bad CRC-32 for file 'ppt/media/" in message:
            return True
        if "bad CRC" in message and "ppt/media/" in message:
            return True
        return False

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

    def _publish_output_pptx(self, source_path: Path, output_name: str) -> Path:
        destination_path = self.settings.downloads_dir / output_name
        destination_path.parent.mkdir(parents=True, exist_ok=True)
        temp_destination = destination_path.with_name(f".{destination_path.name}.tmp")
        if temp_destination.exists():
            temp_destination.unlink()
        shutil.copy2(source_path, temp_destination)
        temp_destination.replace(destination_path)
        return destination_path

    def _open_output_pptx(self, output_path: Path) -> None:
        try:
            if self.settings.platform_name == "windows":
                os.startfile(str(output_path))
                return
            if self.settings.platform_name == "darwin":
                subprocess.Popen(
                    ["open", str(output_path)],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                )
                return
            subprocess.Popen(
                ["xdg-open", str(output_path)],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
        except Exception:
            logger.warning("Failed to open output deck automatically.", exc_info=True, extra={"output_path": str(output_path)})

    def _cleanup_expired_runs(self) -> None:
        cutoff = datetime.now(timezone.utc) - timedelta(hours=self.settings.cleanup_after_hours)
        for run_dir in self.settings.runs_dir.iterdir():
            if not run_dir.is_dir():
                continue
            modified_at = datetime.fromtimestamp(run_dir.stat().st_mtime, tz=timezone.utc)
            if modified_at < cutoff:
                shutil.rmtree(run_dir, ignore_errors=True)
