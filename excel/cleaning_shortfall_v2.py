"""Built-in engine: cleaning roster fill v2 for the new 31-day template.

The new template ships with 31 days pre-configured, including 5th-week
formulas using COUNT() for proportional required hours.  For months with
<31 days we delete excess day columns and let the template's formulas
adapt automatically.
"""
from __future__ import annotations

import calendar
import re
from copy import copy
from datetime import date
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import Workbook

from ..common.col_adjust import (
    update_all_formulas_for_col_change,
    col_num_to_letter,
)

DAY_START_COL = 4   # col D
RANK_COL = 2        # col B
NAME_COL = 3        # col C
TEMPLATE_DAYS = 31  # new template has 31 days pre-configured

_TITLE_RE = re.compile(r"^(\d+)\.\s")
_SUM_SINGLE_RE = re.compile(r"^=SUM\(([A-Z]+)(\d+):\1(\d+)\)$")
_CELL_REF_RE = re.compile(r"(\$?[A-Z]{1,3}\$?)(\d+)")
_RANGE_RE = re.compile(r"([A-Z$]+)(\d+):\1(\d+)")


# ---------------------------------------------------------------------------
# Template layout detection
# ---------------------------------------------------------------------------

_CROSS_SHEET_DATE_RE = re.compile(r"^=[^!'\"\s]+![A-Z]+\d+$")


def _find_date_start(ws):
    """Find the row and column of the first date cell in the worksheet.

    Checks for:
    1. Direct datetime value (e.g. 2026-03-01)
    2. Cross-sheet reference (e.g. =早!I6) — used by multi-sheet templates
       where subsequent sheets reference the first sheet's date anchor.
    """
    for r in range(1, min(ws.max_row, 15) + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v is not None and hasattr(v, "year"):
                return r, c
            if isinstance(v, str) and _CROSS_SHEET_DATE_RE.match(v):
                return r, c
    return None, None


# ---------------------------------------------------------------------------
# Column deletion helpers
# ---------------------------------------------------------------------------

def _delete_one_column(ws, col: int) -> None:
    """Delete a single column, adjusting merged cells and formulas.

    openpyxl handles cell-data shifting; we only need to:
      1. unmerge ranges that touch the deleted column,
      2. update formulas so column refs stay correct,
      3. delete the column,
      4. re-merge adjusted ranges.
    """
    from openpyxl.worksheet.cell_range import CellRange

    to_unmerge: List[str] = []
    to_merge: List[str] = []
    for mr in list(ws.merged_cells.ranges):
        if mr.min_col > col:
            # Entirely to the right — shift left by 1
            new = CellRange(
                min_col=mr.min_col - 1, min_row=mr.min_row,
                max_col=mr.max_col - 1, max_row=mr.max_row,
            )
            to_unmerge.append(str(mr))
            to_merge.append(str(new))
        elif mr.min_col <= col <= mr.max_col:
            # Overlaps deleted column — shrink
            new_max = mr.max_col - 1
            if new_max < mr.min_col:
                to_unmerge.append(str(mr))
            else:
                new = CellRange(
                    min_col=mr.min_col, min_row=mr.min_row,
                    max_col=new_max, max_row=mr.max_row,
                )
                to_unmerge.append(str(mr))
                to_merge.append(str(new))

    for old in to_unmerge:
        ws.unmerge_cells(old)

    update_all_formulas_for_col_change(ws, col, -1, is_delete=True)
    ws.delete_cols(col, 1)

    for new in to_merge:
        ws.merge_cells(new)


# ---------------------------------------------------------------------------
# Day-column adjustment (delete excess from 31-day template)
# ---------------------------------------------------------------------------

def _adjust_day_columns(ws, days_in_month: int) -> int:
    """Adjust day columns from template's 31 days to target month days.

    1. Auto-detects the first date column (works for both cleaning and
       security templates regardless of where the date row starts).
    2. Deletes excess day columns one-by-one from the right edge.
    3. Handles the cleaning-template's "29-31號時數" summary column.
    """
    date_row, first_day_col = _find_date_start(ws)
    if not first_day_col:
        return ws.max_column

    if days_in_month >= TEMPLATE_DAYS:
        return first_day_col + TEMPLATE_DAYS - 1

    # Delete excess day columns from right to left (31, 30, 29).
    # Because each deleted column is to the right of the next target,
    # column numbers of remaining targets do not shift.
    for day in [31, 30, 29]:
        if days_in_month < day:
            col = first_day_col + day - 1
            _delete_one_column(ws, col)

    # Handle the "29-31號時數" summary column
    _adjust_fifth_week_summary_column(ws, days_in_month)

    return first_day_col + days_in_month - 1


def _adjust_fifth_week_summary_column(ws, days_in_month: int) -> None:
    """Delete or rename the '29-31號時數' summary column based on month length.

    - 28 days: delete the entire column (no days 29-31 exist).
    - 30 days: rename header to '29-30號時數'.
    - 31 days: keep as-is.
    """
    fifth_week_col = None
    for c in range(1, ws.max_column + 1):
        for r in range(1, min(ws.max_row, 15) + 1):
            v = ws.cell(r, c).value
            if v and isinstance(v, str) and "29-31" in v:
                fifth_week_col = c
                break
        if fifth_week_col:
            break

    if not fifth_week_col:
        return

    if days_in_month <= 28:
        _delete_one_column(ws, fifth_week_col)
    elif days_in_month == 30:
        for r in range(1, min(ws.max_row, 15) + 1):
            cell = ws.cell(r, fifth_week_col)
            if cell.value and isinstance(cell.value, str):
                cell.value = cell.value.replace("31", "30")


# ---------------------------------------------------------------------------
# Date update
# ---------------------------------------------------------------------------

def _update_dates(ws, month: str, date_row: int | None = None, day_start_col: int | None = None) -> None:
    """Update date row to target month (first cell = date, rest = +1 chain)."""
    if date_row is None or day_start_col is None:
        date_row, day_start_col = _find_date_start(ws)
    if not date_row:
        return

    year, month_num = int(month[:4]), int(month[5:7])
    days_in_month = calendar.monthrange(year, month_num)[1]

    ws.cell(date_row, day_start_col, value=date(year, month_num, 1))

    weekday_row = date_row + 2
    if weekday_row <= ws.max_row:
        for c in range(day_start_col, day_start_col + days_in_month):
            v = ws.cell(weekday_row, c).value
            if isinstance(v, str) and "TEXT" in v:
                col_letter = col_num_to_letter(c)
                ws.cell(weekday_row, c,
                        value=f'=SUBSTITUTE(TEXT({col_letter}{date_row},"aaa"),"週","")')


# ---------------------------------------------------------------------------
# Row insertion helpers (reused from v1)
# ---------------------------------------------------------------------------

def _find_segment_title_rows(ws) -> Dict[str, int]:
    out: Dict[str, int] = {}
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, RANK_COL).value
        if v and isinstance(v, str):
            m = _TITLE_RE.match(v)
            if m:
                out[m.group(1) + "."] = r
    return out


def _find_data_row(ws, start_row: int, end_row: int, day_start_col: int | None = None) -> int | None:
    r"""Find the first personnel data row in a segment.

    v1 templates have =SUM(D\d+:D\d+) with equal row numbers on data rows.
    v2 templates use plain values; we fall back to rank pattern (A\d+) in col B.
    """
    if day_start_col is None:
        _, day_start_col = _find_date_start(ws)
    if not day_start_col:
        return None
    # v1 template: look for single-row SUM formula
    for r in range(start_row + 1, end_row + 1):
        v = ws.cell(r, day_start_col).value
        if v and isinstance(v, str):
            m = _SUM_SINGLE_RE.match(v)
            if m and m.group(2) == m.group(3):
                return int(m.group(2))
    # v2 template: look for rank pattern in col B
    for r in range(start_row + 1, end_row + 1):
        b = ws.cell(r, RANK_COL).value
        if b and isinstance(b, str) and _RANK_RE.match(b.strip()):
            return r
    return None


_RANK_RE = re.compile(r"^A\d+$")


def _count_template_personnel_rows(ws, start_row: int, end_row: int) -> int:
    """Count how many template personnel rows exist between start and end."""
    count = 0
    for r in range(start_row, end_row + 1):
        b = ws.cell(r, RANK_COL).value
        if b and isinstance(b, str) and _RANK_RE.match(b.strip()):
            count += 1
    return count


def _insert_and_style(ws, insert_at: int, amount: int, source_row: int) -> None:
    from openpyxl.worksheet.cell_range import CellRange
    affected: List[Tuple[str, str]] = []
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row >= insert_at:
            new = CellRange(min_col=mr.min_col, min_row=mr.min_row + amount,
                            max_col=mr.max_col, max_row=mr.max_row + amount)
        elif mr.max_row >= insert_at:
            new = CellRange(min_col=mr.min_col, min_row=mr.min_row,
                            max_col=mr.max_col, max_row=mr.max_row + amount)
        else:
            continue
        affected.append((str(mr), str(new)))
    for old, _ in affected:
        try:
            ws.unmerge_cells(old)
        except KeyError:
            # Range may reference cells already removed by a prior delete_rows.
            pass
    ws.insert_rows(insert_at, amount=amount)
    for _, new in affected:
        try:
            ws.merge_cells(new)
        except ValueError:
            # Range may already exist after prior insert_rows adjustments.
            pass
    _shift_print_area(ws, insert_at, amount)
    for off in range(amount):
        target_row = insert_at + off
        for col in range(1, ws.max_column + 1):
            src = ws.cell(source_row, col)
            dst = ws.cell(target_row, col)
            if src.has_style:
                dst.font = copy(src.font)
                dst.border = copy(src.border)
                dst.fill = copy(src.fill)
                dst.number_format = copy(src.number_format)
                dst.protection = copy(src.protection)
                dst.alignment = copy(src.alignment)


def _shift_print_area(ws, insert_at: int, amount: int) -> None:
    _PRINT_AREA_RE = re.compile(r"\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)")
    pa = ws.print_area
    if pa:
        def repl(m: re.Match) -> str:
            c1, r1, c2, r2 = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))
            if r1 >= insert_at: r1 += amount
            if r2 >= insert_at: r2 += amount
            return f"${c1}${r1}:${c2}${r2}"
        new_pa = _PRINT_AREA_RE.sub(repl, str(pa))
        if new_pa != pa:
            ws.print_area = new_pa
    titles = ws.print_title_rows
    if titles:
        m = re.match(r"\$(\d+):\$(\d+)", titles)
        if m:
            r1, r2 = int(m.group(1)), int(m.group(2))
            if r1 >= insert_at: r1 += amount
            if r2 >= insert_at: r2 += amount
            new_titles = f"${r1}:${r2}"
            if new_titles != titles:
                ws.print_title_rows = new_titles


def _shift_formulas(ws, insert_at: int, amount: int) -> None:
    def repl(m: re.Match) -> str:
        col, r = m.group(1), int(m.group(2))
        if r >= insert_at:
            return f"{col}{r + amount}"
        return m.group(0)
    for row in range(1, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            v = ws.cell(row, col).value
            if isinstance(v, str) and v.startswith("="):
                new = _CELL_REF_RE.sub(repl, v)
                if new != v:
                    ws.cell(row, col).value = new


def _expand_range(ws, segment_data_row: int, segment_end_after: int) -> None:
    def repl(m: re.Match) -> str:
        col, r1, r2 = m.group(1), int(m.group(2)), int(m.group(3))
        if r1 == r2 == segment_data_row:
            return f"{col}{r1}:{col}{segment_end_after}"
        return m.group(0)
    for row in range(1, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            v = ws.cell(row, col).value
            if isinstance(v, str) and v.startswith("="):
                new = _RANGE_RE.sub(repl, v)
                if new != v:
                    ws.cell(row, col).value = new


def _copy_row_formulas(ws, source_row: int, insert_at: int, amount: int) -> None:
    for off in range(amount):
        target_row = insert_at + off
        for col in range(1, ws.max_column + 1):
            sv = ws.cell(source_row, col).value
            if isinstance(sv, str) and sv.startswith("="):
                def _rewrite(m: re.Match) -> str:
                    col_part, ref_row = m.group(1), int(m.group(2))
                    if ref_row == source_row:
                        return f"{col_part}{target_row}"
                    return m.group(0)
                ws.cell(target_row, col).value = _CELL_REF_RE.sub(_rewrite, sv)


def _fill_data_row(ws, r: int, row_data: Dict[str, Any], days_in_month: int, day_start_col: int | None = None) -> None:
    if day_start_col is None:
        _, day_start_col = _find_date_start(ws)
    if not day_start_col:
        return
    ws.cell(r, RANK_COL, value=row_data.get("rank_seq") or row_data.get("employee_no") or "")
    ws.cell(r, NAME_COL, value=row_data.get("name") or "")
    cells = {c["day"]: c for c in row_data.get("cells", [])}
    for d in range(1, days_in_month + 1):
        col = day_start_col + d - 1
        c = cells.get(d)
        if not c:
            ws.cell(r, col, value="")
            continue
        v = c.get("value")
        if v is None:
            v = c.get("code")
        try:
            num = float(v)
            ws.cell(r, col, value=int(num) if num == int(num) else num)
        except (TypeError, ValueError):
            ws.cell(r, col, value=str(v) if v else "")


# ---------------------------------------------------------------------------
# Main entry
# ---------------------------------------------------------------------------

def _fetch_data(context: Dict[str, Any]) -> Dict[str, Any]:
    """Call the API to fetch roster data.  Extracted so tests can monkey-patch."""
    from ...api.shortfall_engine import api_shortfall_engine
    return api_shortfall_engine(
        shift="A", project_id=context["project_id"],
        category_id=context["category_id"], month=context["month"],
    )


def run(wb: Workbook, context: Dict[str, Any]) -> None:
    """Fill the cleaning roster template (v2) for the 31-day pre-configured template."""
    data = _fetch_data(context)
    if not data.get("has_data"):
        return

    segments = data.get("segments", [])
    days_in_month = data.get("days_in_month", 31)

    # Phase 0 + 1: Adjust day columns and update dates on EVERY sheet
    for ws in wb.worksheets:
        if ws.title == "Data":
            continue
        # Detect template layout (date row + first date column)
        date_row, day_start_col = _find_date_start(ws)
        if not day_start_col:
            continue
        _adjust_day_columns(ws, days_in_month)
        _update_dates(ws, context["month"], date_row=date_row, day_start_col=day_start_col)

    # -----------------------------------------------------------------------
    # Data fill: currently only processes the first non-Data sheet.
    # Multi-sheet templates (e.g. TY 早/中/夜) need per-shift data.
    # -----------------------------------------------------------------------
    ws = wb.worksheets[0]
    if ws.title == "Data" and len(wb.worksheets) > 1:
        ws = wb.worksheets[1]

    date_row, day_start_col = _find_date_start(ws)
    if not day_start_col:
        return

    # Discover segment positions
    title_rows = _find_segment_title_rows(ws)
    if not title_rows:
        return
    sorted_titles = sorted(title_rows.items(), key=lambda x: x[1])
    section_bounds: Dict[str, Dict[str, int]] = {}
    for i, (prefix, tr) in enumerate(sorted_titles):
        end = sorted_titles[i + 1][1] - 1 if i + 1 < len(sorted_titles) else ws.max_row
        dr = _find_data_row(ws, tr, end, day_start_col=day_start_col)
        if dr is not None:
            section_bounds[prefix] = {"title_row": tr, "data_row": dr, "end_row": end}

    # Match API segments
    api_by_prefix: Dict[str, Dict[str, Any]] = {}
    for s in segments:
        m = _TITLE_RE.match(s.get("title", ""))
        if m:
            api_by_prefix[m.group(1) + "."] = s

    # Bottom-up: insert + fill per section
    ordered = sorted(section_bounds.keys(), key=lambda p: section_bounds[p]["data_row"], reverse=True)
    for prefix in ordered:
        sb = section_bounds[prefix]
        seg = api_by_prefix.get(prefix)
        if not seg or not seg.get("rows"):
            continue
        rows = seg["rows"]
        n = len(rows)
        template_data_row = sb["data_row"]
        template_count = _count_template_personnel_rows(ws, template_data_row, sb["end_row"])

        if n < template_count:
            # Remove excess template rows
            delete_from = template_data_row + n
            delete_count = template_count - n
            ws.delete_rows(delete_from, delete_count)
            # Update row numbers for all subsequent sections
            for other_sb in section_bounds.values():
                for key in ("title_row", "data_row", "end_row"):
                    if other_sb[key] > delete_from:
                        other_sb[key] -= delete_count
        elif n > template_count:
            # Insert additional rows
            insert_at = template_data_row + template_count
            insert_count = n - template_count
            _insert_and_style(ws, insert_at, insert_count, source_row=template_data_row)
            _shift_formulas(ws, insert_at, insert_count)
            _expand_range(ws, template_data_row, template_data_row + n - 1)
            _copy_row_formulas(ws, template_data_row, insert_at, insert_count)
            # Update row numbers for all subsequent sections
            for other_sb in section_bounds.values():
                for key in ("title_row", "data_row", "end_row"):
                    if other_sb[key] >= insert_at:
                        other_sb[key] += insert_count

        for idx, row_data in enumerate(rows):
            _fill_data_row(ws, template_data_row + idx, row_data, days_in_month, day_start_col=day_start_col)
