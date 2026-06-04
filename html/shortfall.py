"""HTML rendering engine for shortfall roster pages.

Extracted from api/shortfall_engine.py so the same logic can be reused by
both the web frontend and Excel export engines without going through HTTP.
"""
from __future__ import annotations

import calendar
import re as _re
from datetime import date
from typing import Any, Dict, List, Optional

from ...config import DOC_DIR
from ...shortfall_engine.parser import TemplateStructure
from ..common import (
    SHIFT_INFO,
    _get_holidays_for_month,
    _weekday_strip,
    load_template_workbook,
    parse_template,
    render_sheet,
)
from ..common.data_source import load_pages, resolve_template


def _to_json(sheet, has_data: bool = True) -> Dict[str, Any]:
    """Serialize RenderedSheet to frontend-friendly JSON."""
    def row_to_dict(r):
        return {
            "page_id": r.page_id,
            "rank_seq": r.rank_seq,
            "employee_no": r.employee_no,
            "name": r.name,
            "rank_code": r.rank_code,
            "position_text": r.position_text,
            "shift_label": r.shift_label,
            "workdays": r.workdays,
            "rest_days": r.rest_days,
            "total_hours": r.total_hours,
            "user_verified": r.user_verified,
            "is_part_time": r.is_part_time,
            "is_supervisor": r.is_supervisor,
            "segment": r.segment,
            "cells": [{"day": c.day, "code": c.code, "edited": c.edited, "value": c.value} for c in r.cells],
            "is_missing": r.is_missing,
            "contracted_hours": r.contracted_hours,
        }

    def seg_to_dict(seg):
        return {
            "title": seg.title,
            "type": seg.type,
            "rows": [row_to_dict(r) for r in seg.rows] if has_data else [],
            "summary": seg.summary,
        }

    return {
        "shift": sheet.shift,
        "label": sheet.label,
        "title": sheet.title,
        "month": sheet.month,
        "year": sheet.year,
        "month_num": sheet.month_num,
        "days_in_month": sheet.days_in_month,
        "weekday_strip": sheet.weekday_strip,
        "holiday_strip": sheet.holiday_strip,
        "sundays": sheet.sundays,
        "contract_required": sheet.contract_required,
        "segments": [seg_to_dict(s) for s in sheet.segments],
        "extra_rows": [row_to_dict(r) for r in sheet.extra_rows] if has_data else [],
        "template_path": sheet.template_path,
        "has_data": has_data,
        "has_supervisor": any(s.type == "supervisor" for s in sheet.segments),
        "bottom_legend": getattr(sheet, "bottom_legend", None),
    }


def run(
    shift: str,
    project_id: Optional[int] = None,
    category_id: Optional[int] = None,
    month: Optional[str] = None,
) -> Dict[str, Any]:
    """Render shortfall roster data for a shift.

    Returns a frontend-friendly JSON dict with segments, rows, cells, etc.
    """
    if shift not in SHIFT_INFO:
        raise ValueError(f"shift must be A/B/C, got {shift}")
    info = SHIFT_INFO[shift]

    excel_path, rel_path = resolve_template(project_id, category_id, month, "roster_shortfall")

    pages = load_pages(shift, project_id=project_id, category_id=category_id, month=month)
    has_data = len(pages) > 0

    if not excel_path or not excel_path.exists():
        if not pages:
            effective_month = month or "2026-03"
            try:
                y, m = effective_month.split("-")
                year, month_num = int(y), int(m)
            except Exception:
                year, month_num = 2026, 3
            days_in_month = calendar.monthrange(year, month_num)[1]
            weekday_strip = _weekday_strip(year, month_num, days_in_month)
            holiday_map = _get_holidays_for_month(year, month_num)
            holiday_strip = [holiday_map.get(d, None) for d in range(1, days_in_month + 1)]
            return {
                "shift": shift,
                "label": info["label"],
                "title": info["title"],
                "month": effective_month,
                "year": year,
                "month_num": month_num,
                "days_in_month": days_in_month,
                "weekday_strip": weekday_strip,
                "holiday_strip": holiday_strip,
                "sundays": sum(1 for d in range(1, days_in_month + 1) if date(year, month_num, d).weekday() == 6),
                "contract_required": 0,
                "segments": [],
                "extra_rows": [],
                "template_path": "",
                "has_data": False,
            }
        empty_tpl = TemplateStructure(
            sheet_name="",
            header_row=0,
            data_start_row=0,
            data_end_row=0,
            col_map={},
            segments=[],
            summary_rows=[],
            styles={},
            day_start_col=0,
            total_days_col=None,
            total_hours_col=None,
            required_counts={},
        )
        sheet = render_sheet(empty_tpl, pages, shift, month=month, template_path="")
        return _to_json(sheet, has_data=True)

    tpl = parse_template(str(excel_path), shift)
    sheet = render_sheet(tpl, pages, shift, month=month, template_path=rel_path or "")

    # Post-process: extract title and bottom legend from template
    try:
        wb_do = load_template_workbook(str(excel_path), data_only=True)
        wb_fo = load_template_workbook(str(excel_path), data_only=False)
        ws_do = wb_do[tpl.sheet_name] if tpl.sheet_name in wb_do.sheetnames else wb_do.worksheets[0]
        ws_fo = wb_fo[tpl.sheet_name] if tpl.sheet_name in wb_fo.sheetnames else wb_fo.worksheets[0]

        found = None
        for row_d, row_f in zip(
            ws_do.iter_rows(min_row=1, max_row=6, max_col=6, values_only=True),
            ws_fo.iter_rows(min_row=1, max_row=6, max_col=6, values_only=True),
        ):
            for v_d, v_f in zip(row_d, row_f):
                if isinstance(v_d, str) and "出勤" in v_d:
                    found = v_d.strip(); break
                if isinstance(v_f, str) and v_f.startswith("=") and "出勤" in v_f:
                    m = _re.search(r'&"月([^"]*)"', v_f)
                    if m:
                        found = f"{sheet.year}年{sheet.month_num}月{m.group(1).strip()}"
                        break
            if found: break
        if found:
            m2 = _re.match(r"^\d{4}年\d{1,2}月\s*(.+)$", found)
            sheet.title = m2.group(1) if m2 else found

        leg = None
        for row in ws_do.iter_rows(min_row=1, max_row=ws_do.max_row,
                                    max_col=ws_do.max_column, values_only=True):
            for v in row:
                if isinstance(v, str) and "【AL=" in v:
                    leg = v.strip(); break
            if leg: break
        if leg:
            sheet.bottom_legend = leg
    except Exception:
        pass

    return _to_json(sheet, has_data=has_data)
