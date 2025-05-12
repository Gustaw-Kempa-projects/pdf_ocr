from __future__ import annotations

import datetime as _dt
from pathlib import Path

import pandas as pd
from xlsxwriter.workbook import Workbook

from .utils import is_valid_float

_HEADERS = [
    "Lp.",
    "Umowa",
    "Aukcja",
    "Data zamowienia",
    "Nr zamowienia",
    "Nr pozycji",
    "Lokalizacja",
    "Nazwa",
    "Liczba sztuk",
    "Cena sprzedazy netto",
    "Wartosc sprzedazy netto",
    "Realizacja od",
    "Realizacja do",
]

_SUMMARY_HEADERS = ["Nr. zamowienia", "Suma"]


def dataframe_to_excel(df: pd.DataFrame) -> Path:
    """
    Dump *df* to an XLSX file and return the path.
    """
    ts = _dt.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    outfile = Path(f"processed_orders_{ts}.xlsx")
    workbook = Workbook(str(outfile))
    sheet = workbook.add_worksheet()

    # body
    for c, header in enumerate(_HEADERS):
        sheet.write(0, c, header)
    for r, (_, row) in enumerate(df.iterrows(), start=1):
        _write_body_row(sheet, r, row)

    # small totals table on the right
    sheet.write(0, 14, _SUMMARY_HEADERS[0])
    sheet.write(0, 15, _SUMMARY_HEADERS[1])
    for idx, (order_nr, group) in enumerate(df.groupby("Nr zamowienia"), start=1):
        valid_vals = [
            float(v.replace(",", "."))
            for v in group["Wartosc"]
            if is_valid_float(v.replace(",", "."))
        ]
        sheet.write(idx, 14, order_nr)
        sheet.write(idx, 15, sum(valid_vals))

    # Excel niceties
    sheet.add_table(0, 0, len(df), len(_HEADERS) - 1, {"columns": [{"header": h} for h in _HEADERS]})
    sheet.add_table(
        0,
        14,
        len(df.groupby("Nr zamowienia")),
        15,
        {"columns": [{"header": h} for h in _SUMMARY_HEADERS]},
    )
    sheet.autofit()
    workbook.close()
    return outfile


def _write_body_row(sheet, r: int, row):
    """One long ugly mapping kept isolated here."""
    sheet.write(r, 0, int(row["Lp."]))
    sheet.write(r, 1, int(row["Umowa"]))
    sheet.write(r, 2, row["Aukcja"])
    sheet.write(r, 3, row["Data zamowienia"])
    sheet.write(r, 4, row["Nr zamowienia"])
    sheet.write(r, 5, int(row["Lp."]))  # 'Nr pozycji' duplicates 'Lp.' for now
    sheet.write(r, 6, row["Zakład"])
    sheet.write(r, 7, row["Nazwa materiału"])
    sheet.write(r, 8, float(row["Ilość w Jm"].replace(",", ".")))
    sheet.write(r, 9, float(row["Cena"].replace(",", ".")))
    sheet.write(r, 10, float(row["Wartość"].replace(",", ".")))
    sheet.write(r, 11, row["Realizacja od"])
    sheet.write(r, 12, row["Realizacja do"])