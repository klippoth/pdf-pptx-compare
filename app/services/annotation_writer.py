from __future__ import annotations

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.util import Pt

from app.services.models import QcFindingType, SlideQcResult, SlideQcStatus


class AnnotationWriter:
    def __init__(self):
        self._colors = {
            QcFindingType.MISSING_CONTENT: RGBColor(220, 40, 40),
            QcFindingType.EXTRA_CONTENT: RGBColor(176, 76, 18),
            QcFindingType.WRONG_TEXT: RGBColor(49, 109, 214),
            QcFindingType.LINE_BREAK_ISSUE: RGBColor(226, 145, 25),
            QcFindingType.ALIGNMENT_DRIFT: RGBColor(49, 109, 214),
            QcFindingType.SIZE_POSITION_ISSUE: RGBColor(49, 109, 214),
            QcFindingType.WRONG_COLOR: RGBColor(148, 56, 201),
        }

    def apply(self, slide, slide_qc_result: SlideQcResult, slide_width, slide_height) -> None:
        for finding in slide_qc_result.findings:
            color = self._colors[finding.finding_type]
            left = int(slide_width * finding.bbox[0])
            top = int(slide_height * finding.bbox[1])
            width = max(int(slide_width * (finding.bbox[2] - finding.bbox[0])), int(slide_width * 0.02))
            height = max(int(slide_height * (finding.bbox[3] - finding.bbox[1])), int(slide_height * 0.02))

            box = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, left, top, width, height)
            box.name = f"QC_{finding.finding_type.value}_{finding.finding_id}"
            box.fill.background()
            box.line.color.rgb = color
            box.line.width = Pt(2.0)

        if slide_qc_result.status == SlideQcStatus.FINDINGS and (
            slide_qc_result.summary or slide_qc_result.comment_bullets or slide_qc_result.note
        ):
            self._add_summary_note(
                slide,
                slide_qc_result.summary,
                slide_qc_result.comment_bullets,
                slide_qc_result.note,
                slide_width,
                slide_height,
            )
        if slide_qc_result.status == SlideQcStatus.MANUAL_REVIEW and slide_qc_result.note:
            self._add_manual_review_note(slide, slide_qc_result.note, slide_width, slide_height)

    def _add_summary_note(self, slide, summary: str | None, bullets: list[str], note: str | None, slide_width, slide_height) -> None:
        rendered_bullets = [bullet.strip() for bullet in bullets if bullet and bullet.strip()]
        rendered_summary = summary.strip() if summary and summary.strip() else None
        if rendered_summary is None and note:
            rendered_summary = note.strip().splitlines()[0].strip()
        bullet_count = max(1, len(rendered_bullets))
        note_width = int(slide_width * 0.42)
        note_height = int(slide_height * min(0.26, 0.10 + (bullet_count * 0.035)))
        left = int(slide_width * 0.03)
        top = int(slide_height * 0.03)
        box = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, left, top, note_width, note_height)
        box.name = "QC_SUMMARY"
        box.fill.solid()
        box.fill.fore_color.rgb = RGBColor(255, 249, 231)
        box.line.color.rgb = RGBColor(212, 196, 140)

        text_frame = box.text_frame
        text_frame.clear()
        text_frame.word_wrap = True

        if rendered_summary:
            paragraph = text_frame.paragraphs[0]
            run = paragraph.add_run()
            run.text = rendered_summary
            run.font.size = Pt(12)
            run.font.bold = True
            run.font.color.rgb = RGBColor(58, 58, 58)

        for bullet in rendered_bullets:
            paragraph = text_frame.add_paragraph()
            run = paragraph.add_run()
            run.text = f"• {bullet}"
            run.font.size = Pt(11)
            run.font.bold = False
            run.font.color.rgb = RGBColor(58, 58, 58)

        if not rendered_bullets and note:
            paragraph = text_frame.add_paragraph() if rendered_summary else text_frame.paragraphs[0]
            run = paragraph.add_run()
            run.text = note if not rendered_summary else f"• {note}"
            run.font.size = Pt(11 if rendered_summary else 12)
            run.font.bold = False
            run.font.color.rgb = RGBColor(58, 58, 58)

    def _add_manual_review_note(self, slide, note: str, slide_width, slide_height) -> None:
        note_width = int(slide_width * 0.38)
        note_height = int(slide_height * 0.08)
        left = int(slide_width * 0.03)
        top = int(slide_height * 0.03)
        pill = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, left, top, note_width, note_height)
        pill.name = "QC_MANUAL_REVIEW"
        pill.fill.solid()
        pill.fill.fore_color.rgb = RGBColor(88, 94, 104)
        pill.line.color.rgb = RGBColor(88, 94, 104)
        text_frame = pill.text_frame
        text_frame.clear()
        paragraph = text_frame.paragraphs[0]
        run = paragraph.add_run()
        run.text = note
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
