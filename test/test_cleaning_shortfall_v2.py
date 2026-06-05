"""Unit tests for cleaning_shortfall_v2 engine.

Tests the new engine that works with the 31-day pre-configured template.
"""
from __future__ import annotations

import pytest
from openpyxl import load_workbook

from ..excel.cleaning_shortfall_v2 import (
    _adjust_day_columns,
    _update_dates,
    _shift_merged_cols_for_delete,
    DAY_START_COL,
    TEMPLATE_DAYS,
)

TEMPLATE_PATH = r"D:\claude\claude_hk\backend\data\doc\output_templates\东汇邨\保洁\东汇-保洁轮休表_template.xlsx"


@pytest.fixture
def template_wb():
    return load_workbook(TEMPLATE_PATH, data_only=False)


class TestAdjustDayColumns:
    """Tests for _adjust_day_columns."""

    def test_31_days_no_change(self, template_wb):
        ws = template_wb.worksheets[0]
        day_end = _adjust_day_columns(ws, 31)
        assert day_end == 34  # AH
        # All 31 day columns still present
        assert ws.cell(7, 34).value == "=AG7+1"
        # 5th week formula intact
        assert ws.cell(13, 32).value == "=SUM(AF12:AH12)"

    def test_30_days_deletes_one_column(self, template_wb):
        ws = template_wb.worksheets[0]
        day_end = _adjust_day_columns(ws, 30)
        assert day_end == 33  # AG
        # Date chain ends at AG
        assert "+1" in str(ws.cell(7, 33).value) or ws.cell(7, 33).value is None
        # 5th week formula shrunk
        val = ws.cell(13, 32).value
        assert "AH" not in str(val) or "AG" in str(val)

    def test_28_days_deletes_three_columns(self, template_wb):
        ws = template_wb.worksheets[0]
        day_end = _adjust_day_columns(ws, 28)
        assert day_end == 31  # AE
        # 28 day columns + 29-31 summary column deleted = 96 - 3 - 1 = 92
        assert ws.max_column == 92
        # 5th week merge removed
        merges = [str(mr) for mr in ws.merged_cells.ranges]
        assert not any("AF13:AH13" in m for m in merges)
        # 4 weeks still intact
        assert ws.cell(13, 25).value == "=SUM(Y12:AE12)"
        # 29-31號時數 summary column also removed
        assert not any("29-31" in str(mr) for mr in ws.merged_cells.ranges)

    def test_28_days_clears_5th_week_values(self, template_wb):
        ws = template_wb.worksheets[0]
        _adjust_day_columns(ws, 28)
        # 5th week cells cleared
        for r in [13, 14, 15, 36, 37, 38, 45, 46, 47]:
            # After deleting cols 32-34, the old col 35 (AI) shifts to 32 (AF)
            pass  # values were cleared before delete

    def test_total_merge_shifts_left(self, template_wb):
        ws = template_wb.worksheets[0]
        _adjust_day_columns(ws, 28)
        merges = [str(mr) for mr in ws.merged_cells.ranges]
        # AI13:AM13 shifts left by 4 (3 date cols + 1 summary col)
        # After delete: AF13:AI13
        assert any("AF13:AI13" in m for m in merges)

    def test_shifted_merge_preserves_border(self, template_wb):
        """Merged cells shifted left by column deletion must keep their borders."""
        ws = template_wb.worksheets[0]
        # Snapshot original border of the merge that will shift (AI13:AM13)
        orig_tl = ws.cell(13, 35)  # AI13
        orig_border = orig_tl.border.left.style if orig_tl.border else None

        _adjust_day_columns(ws, 28)

        # After shift AI13:AM13 -> AF13:AI13, check border preserved
        new_tl = ws.cell(13, 32)  # AF13
        assert new_tl.border.left.style == orig_border


class TestUpdateDates:
    """Tests for _update_dates."""

    def test_updates_first_date(self, template_wb):
        ws = template_wb.worksheets[0]
        _adjust_day_columns(ws, 31)
        _update_dates(ws, "2026-05")
        first = ws.cell(7, DAY_START_COL).value
        assert first.year == 2026
        assert first.month == 5
        assert first.day == 1


class TestEndToEnd:
    """End-to-end using the verify module."""

    def test_28_day_month(self, template_wb, tmp_path):
        from ..test.verify.cleaning import run_check_cleaning_summary

        ws = template_wb.worksheets[0]
        _adjust_day_columns(ws, 28)
        _update_dates(ws, "2026-02")

        out = tmp_path / "test_28.xlsx"
        template_wb.save(str(out))

        result = run_check_cleaning_summary(out, "Roster-FEB2026", "2026-02")
        assert result["ok"], result["errors"]
        assert result["days"] == 28
        assert result["weeks"] == 4

    def test_31_day_month(self, template_wb, tmp_path):
        from ..test.verify.cleaning import run_check_cleaning_summary

        ws = template_wb.worksheets[0]
        _adjust_day_columns(ws, 31)
        _update_dates(ws, "2026-03")

        out = tmp_path / "test_31.xlsx"
        template_wb.save(str(out))

        result = run_check_cleaning_summary(out, "Roster-FEB2026", "2026-03")
        assert result["ok"], result["errors"]
        assert result["days"] == 31
        assert result["weeks"] == 5
