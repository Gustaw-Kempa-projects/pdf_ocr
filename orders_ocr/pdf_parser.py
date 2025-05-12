from __future__ import annotations

import re
from pathlib import Path
from typing import Iterable

import pandas as pd
from PyPDF2 import PdfReader

from .utils import pop_word


def parse_pdf(file_path: Path | str) -> pd.DataFrame:
    """
    Extract order-line tables from *file_path* into a tidy DataFrame.
    """
    reader = PdfReader(str(file_path), strict=False)
    pages = len(reader.pages)
    intro = reader.pages[0].extract_text()

    # quick metadata pulls
    def _scan(pattern: str, group_id: int = 1) -> str:
        return re.search(pattern, intro).group(group_id)

    contract_no = _scan(r"Warunki płatności:\s*(?:AUKCJA|PWD) (\d+)", 1)
    auction_no = _scan(r"Warunki płatności:\s*(?:AUKCJA|PWD) \d+ (.*)", 1)
    order_date = _scan(r"Z dnia:\s*(\d{2}.\d{2}.\d{4})")
    order_nr = _scan(r"Z dnia:.*?Nr (.*?)Oświadczamy")
    period = re.search(r"od\s*(\d{2}.\d{2}.\d{4})\s*do\s*(\d{2}.\d{2}.\d{4})", intro)
    start_date, end_date = period.group(1), period.group(2)

    results = pd.DataFrame(
        columns=[
            "Lp.",
            "Zakład",
            "CPV",
            "Nazwa materiału",
            "Symbol",
            "Jm",
            "Ilość w Jm",
            "Cena",
            "Wartość",
        ]
    )

    y = 0
    start_marker = "Lp. CPV Nazwa materiału Symbol Jm Ilość w Jm Cena Wartość Zap. Zakład"
    for i in range(pages - 1):
        page_text = reader.pages[i].extract_text()
        segment = _extract_table_segment(page_text, start_marker)
        if segment is None:
            continue

        table_lines: list[str] = _merge_wrapped_rows(segment.splitlines(), y)
        for line in table_lines:
            if re.findall(r"Na podst\. zap\.:.*", line):
                continue
            row: dict[str, str] = {}
            line = _cleanup_numbers(line)
            # progressive right-to-left pops
            for col in (
                "Lp.",
                "Zakład",
                "CPV",
                "Wartość",
                "Cena",
                "Ilość w Jm",
                "Jm",
                "Symbol",
            ):
                idx = 0 if col == "Lp." else len(line.split()) - 1
                line = pop_word(line, idx, row, col)
            row["Nazwa materiału"] = line
            results = pd.concat([results, pd.DataFrame([row])], ignore_index=True)
        y += len(table_lines)

    # flat metadata per line
    for col, val in (
        ("Umowa", contract_no),
        ("Aukcja", auction_no),
        ("Data zamowienia", order_date),
        ("Nr zamowienia", order_nr),
        ("Realizacja od", start_date),
        ("Realizacja do", end_date),
    ):
        results[col] = val

    return results


# ----------------------------------------------------------------------------- #
# helpers                                                                       #
# ----------------------------------------------------------------------------- #
def _extract_table_segment(page_text: str, start_marker: str) -> str | None:
    end_marker1, end_marker2 = r"Strona ", r"Wartość słownie:"
    for marker in (end_marker2, end_marker1):
        m = re.search(
            rf"{re.escape(start_marker)}(.*?){re.escape(marker)}", page_text, re.DOTALL
        )
        if m:
            return m.group(1).strip()
    return None


def _merge_wrapped_rows(lines: list[str], y: int) -> list[str]:
    """
    Glue together rows that have wrapped on PDF extraction.
    """
    x = 0
    while x < len(lines):
        if lines[x].strip().startswith(f"{y + 1} "):
            x += 1
            y += 1
        else:
            if lines[x].startswith("Na podst. zap.:"):
                lines.pop(x)
            else:
                lines[x - 1] += lines[x]
                lines.pop(x)
    return lines


def _cleanup_numbers(line: str) -> str:
    """
    Fix various numeric quirks found in the extracted text.
    """
    replacements = {
        r"(\d) (\d{3},\d{2})": r"\1\2",
        r"(\d{2})x(\d{2})": r"\1x\2 ",
        r"(mm)(\d{10})": r"\1 \2",
        r"(mm)(\d)": r"\1 \2 ",
        r"([A-Za-z]{4})(\d{3})": r"\1 \2",
        r"\)(\S)": r") \1",
        r"(?<!\ )(\d{10})(?!\d)": r" \1",
    }
    for pat, repl in replacements.items():
        line = re.sub(pat, repl, line)
    return re.sub(r"\d{2,5}$", "", line)