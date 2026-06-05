"""Generate 3 templates x 2 months for manual inspection.
Run with: python -m pytest test/test_generate_outputs.py -v -s
"""
from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import load_workbook

from ..excel.cleaning_shortfall_v2 import _adjust_day_columns, _update_dates, _find_date_start

OUT_DIR = Path(r"D:\claude\claude_hk\backend\app\engine\test_outputs")

TEMPLATES = [
    ("东汇保安", r"D:\claude\claude_hk\backend\app\engine\templates\东汇-保安轮休表template.xlsx"),
    ("东汇保洁", r"D:\claude\claude_hk\backend\app\engine\templates\东汇-保洁轮休表_template.xlsx"),
    ("TY大元保安", r"D:\claude\claude_hk\backend\app\engine\templates\TY-2026.03- SG_SEC-Deploy Roster  Shortfall - template.xlsx"),
]

MONTHS = [
    ("2026-02", 28),
    ("2026-04", 30),
]


@pytest.fixture(scope="session", autouse=True)
def ensure_out_dir():
    OUT_DIR.mkdir(exist_ok=True)


class TestGenerateOutputs:
    def test_generate_all(self):
        for name, path in TEMPLATES:
            for month, days in MONTHS:
                wb = load_workbook(path, data_only=False)
                for ws in wb.worksheets:
                    if ws.title == "Data":
                        continue
                    date_row, first_col = _find_date_start(ws)
                    if not first_col:
                        continue
                    _adjust_day_columns(ws, days)
                    _update_dates(ws, month, date_row=date_row, day_start_col=first_col)

                out_name = f"{name}_{month}_{days}days.xlsx"
                out_path = OUT_DIR / out_name
                wb.save(str(out_path))
                sheets = ",".join([w.title for w in wb.worksheets if w.title != "Data"])
                print(f"\n  Saved: {out_path}  (sheets: {sheets})")
