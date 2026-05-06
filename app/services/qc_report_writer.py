from __future__ import annotations

import json
from pathlib import Path

from app.services.models import QcReport


class QcReportWriter:
    def write(self, report: QcReport, output_path: Path) -> Path:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "slideCount": len(report.slide_results),
            "manualReviewCount": report.manual_review_count,
            "countsByType": report.counts_by_type,
            "slides": [
                {
                    "slideIndex": result.slide_index,
                    "pageIndex": result.page_index,
                    "status": result.status.value,
                    "alignmentConfidence": result.alignment_confidence,
                    "referenceSource": result.reference_source.value,
                    "candidateSource": result.candidate_source.value,
                    "summary": result.summary,
                    "bullets": result.comment_bullets,
                    "note": result.note,
                    "findings": [
                        {
                            "id": finding.finding_id,
                            "type": finding.finding_type.value,
                            "severity": finding.severity.value,
                            "bbox": list(finding.bbox),
                            "message": finding.message,
                            "confidence": finding.confidence,
                        }
                        for finding in result.findings
                    ],
                }
                for result in report.slide_results
            ],
        }
        output_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        return output_path
