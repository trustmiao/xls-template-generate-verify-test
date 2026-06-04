"""Shared engine infrastructure: template parsing, data rendering, data sources."""
from __future__ import annotations

from .template_parser import parse_template, TemplateStructure
from .render_sheet import render_sheet, RenderedSheet, RenderedRow, RenderedSegment, RenderedCell
from .shift_info import (
    SHIFT_INFO,
    SHIFT_REQUIRED,
    WEEKDAY_CH,
    _build_remark_map,
    _cell_code,
    _get_holidays_for_month,
    _weekday_strip,
    load_template_workbook,
    clear_template_caches,
)
from .excel_ops import (
    DAY_START_COL,
    DAY_END_COL,
    COL_SEQ,
    COL_EMP,
    COL_NAME,
    COL_RANK,
    COL_POS,
    COL_HOURS,
    COL_SHIFT,
    expand_segment_for_count,
    shrink_segment_for_count,
    clear_data_row,
    fill_data_row,
)

__all__ = [
    # template parser
    "parse_template",
    "TemplateStructure",
    # render sheet
    "render_sheet",
    "RenderedSheet",
    "RenderedRow",
    "RenderedSegment",
    "RenderedCell",
    # shift info
    "SHIFT_INFO",
    "SHIFT_REQUIRED",
    "WEEKDAY_CH",
    "_build_remark_map",
    "_cell_code",
    "_get_holidays_for_month",
    "_weekday_strip",
    "load_template_workbook",
    "clear_template_caches",
    # excel ops
    "DAY_START_COL",
    "DAY_END_COL",
    "COL_SEQ",
    "COL_EMP",
    "COL_NAME",
    "COL_RANK",
    "COL_POS",
    "COL_HOURS",
    "COL_SHIFT",
    "expand_segment_for_count",
    "shrink_segment_for_count",
    "clear_data_row",
    "fill_data_row",
]
