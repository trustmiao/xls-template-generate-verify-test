"""JSON mapping rule executor for simple output templates.

Rules JSON shape:
{
  "sheets": [
    {
      "sheet_index": 0,
      "header_row": 1,
      "data_start_row": 2,
      "columns": [
        {"field": "employee_no", "col": 1},
        {"field": "name", "col": 2},
        {"field": "shift", "col": 3}
      ],
      "day_columns": {
        "source": "entries",
        "day_start_col": 4,
        "value_field": "clock_in",
        "match_by": ["page_id"]
      }
    }
  ]
}

For now, a minimal placeholder that can be expanded.
"""
from __future__ import annotations

from typing import Any, Dict, List, Optional

from openpyxl import Workbook


def apply(wb: Workbook, context: Dict[str, Any], rules: Dict[str, Any]) -> None:
    """Apply JSON mapping rules to a workbook."""
    sheets = rules.get("sheets", [])
    for sheet_rule in sheets:
        idx = sheet_rule.get("sheet_index", 0)
        if idx >= len(wb.worksheets):
            continue
        ws = wb.worksheets[idx]
        header_row = sheet_rule.get("header_row", 1)
        data_start_row = sheet_rule.get("data_start_row", header_row + 1)

        # Build a row-by-row mapping from context employees
        employees: List[Dict[str, Any]] = context.get("employees", [])
        columns: List[Dict[str, Any]] = sheet_rule.get("columns", [])
        day_columns = sheet_rule.get("day_columns")

        for row_offset, emp in enumerate(employees):
            row_num = data_start_row + row_offset
            for col_rule in columns:
                field = col_rule.get("field")
                col = col_rule.get("col")
                if field and col:
                    val = emp.get(field, "")
                    ws.cell(row=row_num, column=col, value=val if val is not None else "")

            if day_columns:
                _fill_day_columns(ws, row_num, emp, context, day_columns)


def _fill_day_columns(
    ws,
    row_num: int,
    emp: Dict[str, Any],
    context: Dict[str, Any],
    day_columns: Dict[str, Any],
) -> None:
    """Fill day-based columns from entries data."""
    day_start_col = day_columns.get("day_start_col", 4)
    value_field = day_columns.get("value_field", "clock_in")
    match_by = day_columns.get("match_by", ["employee_no"])

    entries: List[Dict[str, Any]] = context.get("entries", [])
    pages: List[Dict[str, Any]] = context.get("pages", [])
    page_map = {p["id"]: p for p in pages}

    # Build day -> value mapping for this employee
    day_values: Dict[int, Any] = {}
    for e in entries:
        page = page_map.get(e.get("page_id"))
        if not page:
            continue
        matched = True
        for key in match_by:
            if key == "employee_no":
                if (page.get("employee_no") or "").strip() != (emp.get("employee_no") or "").strip():
                    matched = False
                    break
            elif key == "name":
                if (page.get("name") or "").strip() != (emp.get("name") or "").strip():
                    matched = False
                    break
        if matched:
            day = e.get("day")
            if day:
                day_values[day] = e.get(value_field)

    for day, val in day_values.items():
        col = day_start_col + day - 1
        ws.cell(row=row_num, column=col, value=val if val is not None else "")
