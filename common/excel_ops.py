"""Generic Excel segment expand/shrink + cell fill helpers.

Originally lived in `spike_render/spike.py` as part of the experimental
spike for the new fill pipeline. Moved here so output_engines builtins can
reuse them without depending on `spike_render/` (which keeps its CLI entry
points by re-exporting from this module).

Public API:
  expand_segment_for_count(ws, base, tpl_count, needed) -> int
      Add rows to a data segment so it fits `needed` people. Returns rows added.
  shrink_segment_for_count(ws, base, tpl_count, needed) -> int
      Remove rows from a data segment so it fits `needed` people. Returns
      (negative) rows removed.
  clear_data_row(ws, row, *, last_col) -> None
      Clear cells B..last_col on a row, preserving trailing formula cells.
  fill_data_row(ws, row, *, person, seq_label, day_start, day_end, col_seq,
                col_emp, col_name, col_rank, col_pos, col_hours, col_shift)
      Fill one person's row, with optional columns (None = skip).

All mutations honor the same openpyxl-bug workarounds discovered during the
spike (snapshot+rewrite to dodge save-time corruption, in-place
MergedCellRange mutation to avoid cell-wipe side effect of unmerge+merge).
"""
from __future__ import annotations

import re
from copy import copy
from typing import Any, Dict, List, Optional

# Default column positions for the security template family. Callers can
# override every position via fill_data_row's kwargs to handle cleaning or
# other templates that lack some columns.
DAY_START_COL = 9        # I
DAY_END_COL = 36         # AJ
COL_SEQ = 2              # B — rank seq (M01, N02, ...)
COL_EMP = 3              # C — employee number
COL_NAME = 4             # D — name (often a HYPERLINK formula)
COL_RANK = 5             # E — Rank
COL_POS = 6              # F — 職銜
COL_HOURS = 7            # G — 工作時數
COL_SHIFT = 8            # H — 更份


# ---------------------------------------------------------------------------
# Formula helpers
# ---------------------------------------------------------------------------


def _expand_range_end(formula: str, old_end: int, new_end: int) -> str:
    """Extend VERTICAL ranges (multi-row) ending at row `old_end` to `new_end`.

    Only multi-row ranges (row1 != row2) get extended. A single-row
    horizontal range like =COUNT(I10:AJ10) must stay at row 10 — extending
    it would double-count cells.
    """
    if not formula or not formula.startswith("="):
        return formula
    pat = re.compile(r'(\$?[A-Z]+\$?)(\d+):(\$?[A-Z]+\$?)(\d+)')

    def repl(m: re.Match) -> str:
        col1, row1, col2, row2 = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))
        if row1 != row2 and row2 == old_end:
            row2 = new_end
        return f"{col1}{row1}:{col2}{row2}"

    return pat.sub(repl, formula)


def _retarget_row(formula: str, src_row: int, dst_row: int) -> str:
    """Replace any row ref to src_row with dst_row. Used when cloning per-row
    formulas from a baseline row to a newly inserted row."""
    if not formula or not formula.startswith("="):
        return formula
    pat = re.compile(r'(\$?[A-Z]+\$?)(\d+)')

    def repl(m: re.Match) -> str:
        col, row = m.group(1), int(m.group(2))
        if row == src_row:
            row = dst_row
        return f"{col}{row}"

    return pat.sub(repl, formula)


def _shift_formula_refs(formula: str, insert_at: int, shift: int) -> str:
    """Shift any A1-style row reference >= insert_at by +shift.

    openpyxl's `insert_rows(...)` moves cells physically but does NOT update
    formula text inside them. We do it explicitly so formulas like
    =IF((N14-N15)<0,...) become =IF((N15-N16)<0,...) after inserting at row 11.
    """
    if not formula or not formula.startswith("="):
        return formula
    pat = re.compile(r'(\$?[A-Z]+\$?)(\d+)')

    def repl(m: re.Match) -> str:
        col, row = m.group(1), int(m.group(2))
        if row >= insert_at:
            row += shift
        return f"{col}{row}"

    return pat.sub(repl, formula)


# ---------------------------------------------------------------------------
# Row style / snapshot helpers
# ---------------------------------------------------------------------------


def _copy_row_style_and_per_row_formulas(ws, src_row: int, dst_row: int) -> None:
    """Clone src_row's STYLE + per-row FORMULAS onto dst_row.

    - Style (font/fill/border/etc) always copied.
    - Per-row formulas: refs ONLY to src_row → retarget to dst_row; refs to
      no row at all → copy as-is.
    - Data values (literals like '8' or 'RD') are NOT copied.
    """
    only_row_pat = re.compile(r'\$?[A-Z]+\$?(\d+)')

    for c in range(1, ws.max_column + 1):
        src_cell = ws.cell(src_row, c)
        dst_cell = ws.cell(dst_row, c)
        if src_cell.has_style:
            dst_cell.font = copy(src_cell.font)
            dst_cell.fill = copy(src_cell.fill)
            dst_cell.border = copy(src_cell.border)
            dst_cell.alignment = copy(src_cell.alignment)
            dst_cell.number_format = src_cell.number_format
            dst_cell.protection = copy(src_cell.protection)
        v = src_cell.value
        if isinstance(v, str) and v.startswith("="):
            rows_in_formula = {int(r) for r in only_row_pat.findall(v)}
            if not rows_in_formula or rows_in_formula == {src_row}:
                dst_cell.value = _retarget_row(v, src_row, dst_row)


def _snapshot_rows(ws, rows: List[int]) -> Dict[int, Dict[int, Any]]:
    """Snapshot cell values for the given rows using ws._cells dict directly,
    bypassing openpyxl normal-load reader's empty-<v> bug.

    Some cells with <f>...</f><v></v> shape return None via ws.cell().value
    even though the formula is set. Reading from _cells works around it.
    """
    out: Dict[int, Dict[int, Any]] = {}
    max_col = ws.max_column
    for r in rows:
        row_data: Dict[int, Any] = {}
        for c in range(1, max_col + 1):
            cell = ws._cells.get((r, c))
            if cell is None:
                continue
            v = cell.value
            if v is None:
                v = getattr(cell, "_value", None)
            if v is not None and v != "":
                row_data[c] = v
        if row_data:
            out[r] = row_data
    return out


# ---------------------------------------------------------------------------
# Public: segment expand / shrink
# ---------------------------------------------------------------------------


def expand_segment_for_count(ws, base: int, tpl_count: int, needed: int) -> int:
    """Expand template's data slot count from tpl_count to needed.

    Returns the number of rows inserted (0 if no expansion).
    """
    if needed <= tpl_count:
        return 0
    extra = needed - tpl_count
    last_tpl = base + tpl_count - 1
    insert_at = last_tpl + 1
    new_last = last_tpl + extra
    max_col = ws.max_column

    middle_src = base + 1 if tpl_count >= 2 else base
    last_borders = [copy(ws.cell(last_tpl, c).border) for c in range(1, max_col + 1)]
    middle_borders = [copy(ws.cell(middle_src, c).border) for c in range(1, max_col + 1)]

    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if isinstance(v, str) and v.startswith("="):
                new_v = _shift_formula_refs(v, insert_at, extra)
                if new_v != v:
                    cell.value = new_v

    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if isinstance(v, str) and v.startswith("="):
                new_v = _expand_range_end(v, last_tpl, new_last)
                if new_v != v:
                    cell.value = new_v

    affected_rows = list(range(insert_at, ws.max_row + 1))
    snapshot = _snapshot_rows(ws, affected_rows)

    # Pre-shift merged ranges (insert_rows doesn't auto-update merged_cells).
    # Mutate MergedCellRange in-place — unmerge + merge wipes UNRELATED cells
    # in distant columns (openpyxl bug).
    for mr in ws.merged_cells.ranges:
        if mr.min_row >= insert_at:
            mr.shift(0, extra)
        elif mr.max_row >= insert_at:
            mr.max_row += extra

    for _ in range(extra):
        ws.insert_rows(insert_at)

    for old_r, row_data in snapshot.items():
        new_r = old_r + extra
        for c, v in row_data.items():
            ws.cell(new_r, c).value = v

    for r in range(insert_at, new_last + 1):
        _copy_row_style_and_per_row_formulas(ws, base, r)

    for r in range(insert_at, new_last):
        for ci, b in enumerate(middle_borders, start=1):
            ws.cell(r, ci).border = copy(b)
    for ci, b in enumerate(last_borders, start=1):
        ws.cell(new_last, ci).border = copy(b)
    for ci, b in enumerate(middle_borders, start=1):
        ws.cell(last_tpl, ci).border = copy(b)

    return extra


def shrink_segment_for_count(ws, base: int, tpl_count: int, needed: int) -> int:
    """Mirror of expand_segment_for_count for the case engine < tpl_count.

    DELETEs (tpl_count - needed) rows from the end of the segment to avoid
    leaving an empty slot inside the COUNTIF range (which would break daily-
    count formulas).

    Returns the (negative) number of rows removed.
    """
    if needed >= tpl_count:
        return 0
    delete_count = tpl_count - needed
    last_tpl = base + tpl_count - 1
    new_last = base + needed - 1
    delete_from = new_last + 1
    max_col = ws.max_column

    last_borders = [copy(ws.cell(last_tpl, c).border) for c in range(1, max_col + 1)]

    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if isinstance(v, str) and v.startswith("="):
                new_v = _expand_range_end(v, last_tpl, new_last)
                if new_v != v:
                    cell.value = new_v

    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if isinstance(v, str) and v.startswith("="):
                new_v = _shift_formula_refs(v, delete_from + delete_count, -delete_count)
                if new_v != v:
                    cell.value = new_v

    for mr in ws.merged_cells.ranges:
        if mr.min_row > new_last:
            mr.shift(0, -delete_count)
        elif mr.max_row > new_last:
            mr.max_row -= delete_count

    affected_rows = list(range(delete_from + delete_count, ws.max_row + 1))
    snapshot = _snapshot_rows(ws, affected_rows)

    ws.delete_rows(delete_from, delete_count)

    for old_r, row_data in snapshot.items():
        new_r = old_r - delete_count
        for c, v in row_data.items():
            ws.cell(new_r, c).value = v

    for ci, b in enumerate(last_borders, start=1):
        ws.cell(new_last, ci).border = copy(b)

    return -delete_count


# ---------------------------------------------------------------------------
# Public: data row clear + fill
# ---------------------------------------------------------------------------


def clear_data_row(ws, row: int, *, last_col: int) -> None:
    """Clear cells B..last_col on this row.

    Uses cell.value = "" (not None) — None triggers an openpyxl save-time
    bug that corrupts cell storage in distant rows (e.g. row N+12).
    """
    for c in range(COL_SEQ, last_col + 1):
        ws.cell(row, c).value = ""


def fill_data_row(ws, row: int, *, person: Dict[str, Any], seq_label: str,
                  day_start: int = DAY_START_COL, day_end: int = DAY_END_COL,
                  col_seq: int = COL_SEQ, col_emp: Optional[int] = COL_EMP,
                  col_name: Optional[int] = COL_NAME,
                  col_rank: Optional[int] = COL_RANK,
                  col_pos: Optional[int] = COL_POS,
                  col_hours: Optional[int] = COL_HOURS,
                  col_shift: Optional[int] = COL_SHIFT) -> None:
    """Fill one person's row. Any col_* set to None is skipped — handles
    templates lacking some columns (e.g. cleaning roster has only seq+name).
    Day cells always written to [day_start, day_end]."""
    ws.cell(row, col_seq).value = seq_label
    emp = (person.get("employee_no") or "").strip()
    name = (person.get("name") or "").strip()
    pid = person.get("page_id")
    if col_emp is not None:
        ws.cell(row, col_emp).value = emp
    if col_name is not None:
        ws.cell(row, col_name).value = name
    if col_rank is not None:
        ws.cell(row, col_rank).value = person.get("rank") or person.get("rank_code") or ""
    if col_pos is not None:
        ws.cell(row, col_pos).value = person.get("position") or person.get("position_text") or ""
    if col_hours is not None:
        contracted_h = person.get("contracted_hours")
        if contracted_h is not None:
            ws.cell(row, col_hours).value = contracted_h
    if col_shift is not None:
        ws.cell(row, col_shift).value = person.get("shift_label") or person.get("shift") or ""
    cells = person.get("cells") or []
    for c in cells:
        day = c.get("day")
        code = c.get("code") or ""
        if day is None:
            continue
        col = day_start + day - 1
        if not (day_start <= col <= day_end):
            continue
        ws.cell(row, col).value = 8 if code == "8" else code


__all__ = [
    "DAY_START_COL", "DAY_END_COL",
    "COL_SEQ", "COL_EMP", "COL_NAME", "COL_RANK", "COL_POS", "COL_HOURS", "COL_SHIFT",
    "expand_segment_for_count", "shrink_segment_for_count",
    "clear_data_row", "fill_data_row",
]
