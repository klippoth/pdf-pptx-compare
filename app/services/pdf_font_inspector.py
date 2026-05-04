from __future__ import annotations

import json
from pathlib import Path
from typing import Optional

import fitz

from app.services.models import PDFFontInfo, PDFFontInspectionResult


class PDFFontInspector:
    def inspect(self, pdf_path: Path) -> PDFFontInspectionResult:
        document = fitz.open(str(pdf_path))
        fonts_by_key: dict[tuple[str, str], PDFFontInfo] = {}
        page_character_totals: dict[int, int] = {}

        for page_index, page in enumerate(document):
            page_number = page_index + 1
            page_character_totals[page_number] = 0

            for font in page.get_fonts(full=True):
                xref = font[0]
                font_type = str(font[2])
                raw_name = str(font[3] or font[4] or f"Font-{xref}")
                normalized_name = self._normalize_font_name(raw_name)
                extracted_name, _, extracted_type, font_buffer = document.extract_font(xref)
                resolved_name = self._normalize_font_name(str(extracted_name or normalized_name))
                resolved_type = str(extracted_type or font_type)
                embedded = bool(font_buffer)
                subset = self._is_subset_font(raw_name) or self._is_subset_font(str(extracted_name or ""))

                key = (resolved_name, resolved_type)
                info = fonts_by_key.get(key)
                if info is None:
                    info = PDFFontInfo(
                        name=resolved_name,
                        font_type=resolved_type,
                        embedded=embedded,
                        subset=subset,
                        page_numbers=[],
                    )
                    fonts_by_key[key] = info
                else:
                    info.embedded = info.embedded or embedded
                    info.subset = info.subset or subset

                if page_number not in info.page_numbers:
                    info.page_numbers.append(page_number)

            for span_name, character_count in self._extract_page_font_character_counts(page).items():
                page_character_totals[page_number] += character_count
                key = self._resolve_font_key(span_name, fonts_by_key)
                if key is None:
                    key = (span_name, "Unknown")
                    fonts_by_key[key] = PDFFontInfo(
                        name=span_name,
                        font_type="Unknown",
                        embedded=False,
                        subset=self._is_subset_font(span_name),
                        page_numbers=[],
                        page_character_counts={},
                    )

                info = fonts_by_key[key]
                if page_number not in info.page_numbers:
                    info.page_numbers.append(page_number)
                info.page_character_counts[page_number] = info.page_character_counts.get(page_number, 0) + character_count

        fonts = sorted(fonts_by_key.values(), key=lambda item: (not item.embedded, item.name.lower()))
        for font in fonts:
            font.page_numbers.sort()

        result = PDFFontInspectionResult(
            fonts=fonts,
            page_character_totals=page_character_totals,
            page_count=len(document),
        )
        document.close()
        return result

    def write_report(self, result: PDFFontInspectionResult, output_path: Path) -> Path:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        report = {
            "fontCount": len(result.fonts),
            "embeddedFontCount": sum(1 for font in result.fonts if font.embedded),
            "pageCount": result.page_count,
            "pageCharacterTotals": result.page_character_totals,
            "fonts": [
                {
                    "name": font.name,
                    "fontType": font.font_type,
                    "embedded": font.embedded,
                    "subset": font.subset,
                    "pageNumbers": font.page_numbers,
                    "pageCharacterCounts": font.page_character_counts,
                }
                for font in result.fonts
            ],
        }
        output_path.write_text(json.dumps(report, indent=2), encoding="utf-8")
        return output_path

    def _extract_page_font_character_counts(self, page: fitz.Page) -> dict[str, int]:
        counts: dict[str, int] = {}
        text_data = page.get_text("dict")
        for block in text_data.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    text = str(span.get("text") or "")
                    character_count = self._count_extractable_characters(text)
                    if character_count <= 0:
                        continue
                    span_name = self._normalize_font_name(str(span.get("font") or "Unknown Font"))
                    counts[span_name] = counts.get(span_name, 0) + character_count
        return counts

    @staticmethod
    def _count_extractable_characters(text: str) -> int:
        return sum(1 for character in text if not character.isspace())

    @staticmethod
    def _resolve_font_key(
        font_name: str,
        fonts_by_key: dict[tuple[str, str], PDFFontInfo],
    ) -> Optional[tuple[str, str]]:
        matches = [key for key in fonts_by_key if key[0] == font_name]
        if not matches:
            return None
        return sorted(matches, key=lambda key: (key[1] == "Unknown", key[1]))[0]

    @staticmethod
    def _is_subset_font(name: str) -> bool:
        prefix, _, _ = name.partition("+")
        return len(prefix) == 6 and prefix.isupper()

    @staticmethod
    def _normalize_font_name(name: str) -> str:
        _, _, suffix = name.partition("+")
        cleaned = suffix or name
        return cleaned.strip() or "Unknown Font"
