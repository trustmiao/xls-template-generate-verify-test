"""HTML engine: web roster — adjust template and export to HTML.

This engine is completely independent from the Excel export engines.
It loads a template workbook, adjusts it for the target month, fills
personnel data, and exports the result as HTML.
"""
from __future__ import annotations

import calendar
import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles.colors import COLOR_INDEX

from ..common.data_source import resolve_template
from ..common.shift_info import _get_holidays_for_month, WEEKDAY_CH
from .. import excel_row_ops, excel_column_ops
from ..common.col_adjust import update_all_formulas_for_col_change


# ── Constants ──
DAY_START_COL = 4
RANK_COL = 2
NAME_COL = 3
TEMPLATE_DAYS = 31

_TITLE_RE = re.compile(r"^(\d+)\.\s")
_RANK_RE = re.compile(r"^[A-Z]{1,2}\d+$")
_CROSS_SHEET_DATE_RE = re.compile(r"^=[^!'\"\s]+![A-Z]+\d+$")


# ── Inlined from cleaning_shortfall_v2 (to keep engines independent) ──

def _find_date_start(ws):
    """Find the row and column of the first date cell in the worksheet."""
    for r in range(1, min(ws.max_row, 15) + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v is not None and hasattr(v, "year"):
                return r, c
            if isinstance(v, str) and _CROSS_SHEET_DATE_RE.match(v):
                return r, c
    return None, None


def col_num_to_letter(n: int) -> str:
    """Convert 1-based column number to Excel column letter(s)."""
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _delete_one_column(ws, col: int) -> None:
    """Delete a single column using safe helpers (no phantom cells)."""
    update_all_formulas_for_col_change(ws, col, -1, is_delete=True)
    excel_column_ops.safe_delete_cols(ws, col, 1)
    excel_column_ops.fix_merged_cells(ws, col)
    excel_column_ops.adjust_conditional_formatting(ws, col)


def _adjust_fifth_week_summary_column(ws, days_in_month: int) -> None:
    """Delete or rename the '29-31號時數' summary column based on month length."""
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


def _adjust_day_columns(ws, days_in_month: int) -> int:
    """Adjust day columns from template's 31 days to target month days."""
    date_row, first_day_col = _find_date_start(ws)
    if not first_day_col:
        return ws.max_column

    if days_in_month >= TEMPLATE_DAYS:
        return first_day_col + TEMPLATE_DAYS - 1

    for day in [31, 30, 29]:
        if days_in_month < day:
            col = first_day_col + day - 1
            _delete_one_column(ws, col)

    _adjust_fifth_week_summary_column(ws, days_in_month)
    return first_day_col + days_in_month - 1


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


# ── Core roster processing ──

def _html_escape(s: str) -> str:
    return (
        s.replace("&", "&amp;")
         .replace("<", "&lt;")
         .replace(">", "&gt;")
         .replace('"', "&quot;")
    )


def _color_to_hex(color) -> str | None:
    """Convert openpyxl Color to CSS #RRGGBB string.

    Handles indexed, rgb, and theme color types.  Returns None for
    transparent / auto / unresolvable colours.
    """
    if color is None:
        return None

    # Indexed colours (most common in xlsx templates)
    if getattr(color, "type", None) == "indexed":
        idx = color.value
        # idx == 64 means "auto" – use default (None)
        if idx is not None and 0 <= idx < len(COLOR_INDEX):
            rgb = COLOR_INDEX[idx]
            if rgb and isinstance(rgb, str) and len(rgb) == 8 and rgb != "00000000":
                return f"#{rgb[2:]}"
        return None

    # Direct RGB (8-char AARRGGBB or 6-char RRGGBB)
    raw = None
    if getattr(color, "type", None) == "rgb":
        raw = color.value
    elif getattr(color, "rgb", None):
        raw = color.rgb
        # openpyxl rgb descriptor may raise on read; guard against error text
        if isinstance(raw, str) and "must be of type" in raw:
            raw = None
    if raw and isinstance(raw, str):
        raw = raw.strip()
        if len(raw) == 8 and raw != "00000000":
            return f"#{raw[2:]}"
        if len(raw) == 6:
            return f"#{raw}"

    # Theme colours – simplified fallback (would need workbook theme lookup)
    if getattr(color, "type", None) == "theme":
        return None

    return None


# Excel border style → CSS width mapping
_BORDER_WIDTH_MAP = {
    "hair": "0.5px",
    "thin": "1px",
    "medium": "2px",
    "thick": "3px",
    "double": "3px",
    "dashed": "1px",
    "dotted": "1px",
    "mediumDashDot": "2px",
    "mediumDashDotDot": "2px",
    "slantDashDot": "1px",
}


# Excel border style → CSS border-style
_BORDER_STYLE_MAP = {
    "double": "double",
    "dashed": "dashed",
    "dotted": "dotted",
    "mediumDashDot": "dashed",
    "mediumDashDotDot": "dotted",
    "slantDashDot": "dashed",
}


def _shift_from_sheet_name(sheet_name: str) -> str:
    mapping = {"早": "A", "中": "B", "夜": "C"}
    return mapping.get(sheet_name, "A")


def _fetch_data(context: Dict[str, Any], shift: str) -> Dict[str, Any]:
    from ...api.shortfall_engine import api_shortfall_engine
    return api_shortfall_engine(
        shift=shift,
        project_id=context["project_id"],
        category_id=context["category_id"],
        month=context["month"],
    )


def _detect_roster_regions(ws) -> List[Dict[str, Any]]:
    regions = []
    for row in range(1, ws.max_row + 1):
        b_val = ws.cell(row=row, column=RANK_COL).value
        if b_val and isinstance(b_val, str) and _TITLE_RE.match(b_val):
            title_row = row
            data_start = None
            for r in range(title_row + 1, ws.max_row + 1):
                b = ws.cell(row=r, column=RANK_COL).value
                if b and isinstance(b, str) and _RANK_RE.match(b.strip()):
                    data_start = r
                    break
            if data_start is None:
                continue
            data_end = data_start
            for r in range(data_start + 1, ws.max_row + 1):
                b = ws.cell(row=r, column=RANK_COL).value
                if b and isinstance(b, str) and _RANK_RE.match(b.strip()):
                    data_end = r
                else:
                    break
            regions.append({
                "index": len(regions) + 1,
                "title_row": title_row,
                "data_start": data_start,
                "data_end": data_end,
            })
    return regions


def _shrink_region(ws, region: Dict[str, Any], target_count: int) -> int:
    current = region["data_end"] - region["data_start"] + 1
    to_delete = current - target_count
    if to_delete <= 0:
        return 0
    for _ in range(to_delete):
        row_to_delete = region["data_end"]
        excel_row_ops.delete_row_and_fixup(ws, row_to_delete)
        region["data_end"] -= 1
    return to_delete


def _safe_write(ws, row: int, col: int, value) -> bool:
    cell = ws.cell(row, col)
    if isinstance(cell, MergedCell):
        return False
    cell.value = value
    return True


def _fill_data_row(
    ws, row: int, row_data: Dict[str, Any], day_start_col: int, days_in_month: int
) -> None:
    _safe_write(ws, row, RANK_COL, row_data.get("rank_seq") or row_data.get("employee_no") or "")
    _safe_write(ws, row, NAME_COL, row_data.get("name") or "")
    cells = {c["day"]: c for c in row_data.get("cells", [])}
    for d in range(1, days_in_month + 1):
        col = day_start_col + d - 1
        c = cells.get(d)
        if not c:
            _safe_write(ws, row, col, "")
            continue
        v = c.get("value")
        if v is None:
            v = c.get("code")
        try:
            num = float(v)
            _safe_write(ws, row, col, int(num) if num == int(num) else num)
        except (TypeError, ValueError):
            _safe_write(ws, row, col, str(v) if v else "")


def _adjust_number_formats(ws, day_start_col: int, days_in_month: int) -> None:
    for row in range(1, ws.max_row + 1):
        for col in range(day_start_col, day_start_col + days_in_month):
            cell = ws.cell(row, col)
            v = cell.value
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                if isinstance(v, int) or v == int(v):
                    cell.number_format = "0"
                else:
                    cell.number_format = "0.00"


# ── HTML export helpers ──

def _cell_to_css(cell) -> str:
    """Convert openpyxl cell styles to a CSS style string."""
    styles: List[str] = []

    # ── Font ──
    if cell.font:
        if cell.font.name:
            styles.append(f"font-family:{cell.font.name}")
        if cell.font.size:
            styles.append(f"font-size:{cell.font.size}pt")
        if cell.font.bold:
            styles.append("font-weight:bold")
        if cell.font.italic:
            styles.append("font-style:italic")
        if cell.font.underline and cell.font.underline != "none":
            styles.append("text-decoration:underline")
        color_hex = _color_to_hex(cell.font.color)
        if color_hex:
            styles.append(f"color:{color_hex}")

    # ── Fill (background) ──
    if cell.fill:
        # Prefer fgColor; fallback to start_color for older openpyxl
        fill_color = None
        if cell.fill.fgColor:
            fill_color = _color_to_hex(cell.fill.fgColor)
        if not fill_color and hasattr(cell.fill, "start_color") and cell.fill.start_color:
            fill_color = _color_to_hex(cell.fill.start_color)
        if fill_color:
            styles.append(f"background-color:{fill_color}")

    # ── Border ──
    if cell.border:
        for side_name, side in [
            ("top", cell.border.top),
            ("bottom", cell.border.bottom),
            ("left", cell.border.left),
            ("right", cell.border.right),
        ]:
            if side and side.style:
                width = _BORDER_WIDTH_MAP.get(side.style, "1px")
                bstyle = _BORDER_STYLE_MAP.get(side.style, "solid")
                color_hex = _color_to_hex(side.color)
                color_str = color_hex if color_hex else "#000000"
                styles.append(f"border-{side_name}:{width} {bstyle} {color_str}")

    # ── Alignment ──
    if cell.alignment:
        if cell.alignment.horizontal:
            styles.append(f"text-align:{cell.alignment.horizontal}")
        if cell.alignment.vertical:
            v = cell.alignment.vertical
            # openpyxl uses "center" for both; CSS needs explicit mapping
            if v == "center":
                styles.append("vertical-align:middle")
            else:
                styles.append(f"vertical-align:{v}")
        # Text wrap
        if cell.alignment.wrapText:
            styles.append("white-space:normal;word-wrap:break-word")

    return ";".join(styles)


def _format_cell_value(value, computed_value=None) -> str:
    """Format a cell value for HTML display.

    Args:
        value: The raw cell value (may be a formula string).
        computed_value: The computed/cached value from data_only=True load.
            If provided and not None, it takes precedence over ``value``.
    """
    # Prefer the pre-computed (data_only) value when available
    if computed_value is not None:
        value = computed_value

    if value is None:
        return ""
    if hasattr(value, "year"):
        # date/datetime – show day number for date rows
        return str(value.day)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        # Format numbers consistently with Excel
        if isinstance(value, int) or value == int(value):
            return str(int(value))
        return f"{value:.2f}"
    return _html_escape(str(value))


def _find_weekday_row(ws, date_row):
    for r in range(date_row + 1, min(ws.max_row, date_row + 5) + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and "TEXT" in v and "aaa" in v:
                return r
    return None


def workbook_to_html(wb: Workbook, month: str, value_wb: Workbook | None = None) -> str:
    """Convert an openpyxl workbook to a single HTML string.

    Args:
        wb: Workbook loaded with data_only=False (styles, formulas, merged cells).
        month: Target month (YYYY-MM) for holiday/ weekday rendering.
        value_wb: Optional workbook loaded with data_only=True; its cell
            values (pre-computed formula results) are preferred over the
            raw formula strings in ``wb``.

    Features:
      - Renders every non-Data sheet as a separate <table>
      - Preserves merged cells (rowspan/colspan)
      - Preserves cell styles (font, colour, border, alignment)
      - Date row shows day number + public-holiday badge
      - Weekday row shows Chinese weekday characters
      - Formula cells display their computed values
    """
    year, month_num = int(month[:4]), int(month[5:7])
    days_in_month = calendar.monthrange(year, month_num)[1]
    holidays = _get_holidays_for_month(year, month_num)

    # Build a map from sheet title → value_ws for fast lookup
    value_ws_map: Dict[str, Any] = {}
    if value_wb:
        for vws in value_wb.worksheets:
            value_ws_map[vws.title] = vws

    html_parts = [
        "<!DOCTYPE html>",
        '<html>',
        '<head>',
        '<meta charset="utf-8">',
        "<title>Web Roster</title>",
        "<style>",
        "body { font-family: Arial, 'Microsoft JhengHei', sans-serif; margin: 20px; }",
        "table { border-collapse: collapse; margin-bottom: 30px; table-layout: fixed; }",
        "td, th { padding: 2px 4px; white-space: nowrap; overflow: hidden; }",
        '.holiday { color: #c00; font-size: 8pt; font-weight: bold; }',
        "</style>",
        "</head>",
        "<body>",
    ]

    for ws in wb.worksheets:
        if ws.title == "Data":
            continue

        value_ws = value_ws_map.get(ws.title)

        html_parts.append(f'<h2>{_html_escape(ws.title)}</h2>')
        html_parts.append("<table>")

        # ── Column widths (via <col> tags) ──
        for col in range(1, ws.max_column + 1):
            letter = col_num_to_letter(col)
            width = ws.column_dimensions[letter].width if letter in ws.column_dimensions else None
            if width:
                # Excel width → px (approximate: ~7 px per Excel unit)
                px = int(width * 7)
                html_parts.append(f'<col style="width:{px}px">')
            else:
                html_parts.append('<col>')

        date_row, day_start_col = _find_date_start(ws)
        weekday_row = _find_weekday_row(ws, date_row) if date_row else None

        merged_covered: set[tuple[int, int]] = set()
        merged_map: Dict[tuple[int, int], tuple[int, int]] = {}
        for mr in ws.merged_cells.ranges:
            for r in range(mr.min_row, mr.max_row + 1):
                for c in range(mr.min_col, mr.max_col + 1):
                    if (r, c) != (mr.min_row, mr.min_col):
                        merged_covered.add((r, c))
            merged_map[(mr.min_row, mr.min_col)] = (
                mr.max_row - mr.min_row + 1,
                mr.max_col - mr.min_col + 1,
            )

        for row in range(1, ws.max_row + 1):
            # Row height
            row_height = ws.row_dimensions[row].height if row in ws.row_dimensions else None
            row_style = ""
            if row_height:
                px = int(row_height * 1.33)  # Excel point → px (approximate)
                row_style = f' style="height:{px}px"'
            html_parts.append(f"<tr{row_style}>")

            for col in range(1, ws.max_column + 1):
                if (row, col) in merged_covered:
                    continue

                cell = ws.cell(row, col)
                style = _cell_to_css(cell)
                value = cell.value

                # Prefer pre-computed value from data_only workbook
                computed_value = None
                if value_ws:
                    try:
                        computed_value = value_ws.cell(row, col).value
                    except Exception:
                        pass

                if row == date_row and day_start_col and col >= day_start_col:
                    day = col - day_start_col + 1
                    if 1 <= day <= days_in_month:
                        val_str = str(day)
                        if day in holidays:
                            val_str += (
                                f'<br><span class="holiday">'
                                f'{_html_escape(holidays[day])}</span>'
                            )
                    else:
                        val_str = _format_cell_value(value, computed_value)
                elif row == weekday_row and day_start_col and col >= day_start_col:
                    day = col - day_start_col + 1
                    if 1 <= day <= days_in_month:
                        val_str = WEEKDAY_CH[date(year, month_num, day).weekday()]
                    else:
                        val_str = _format_cell_value(value, computed_value)
                else:
                    val_str = _format_cell_value(value, computed_value)

                attrs = ""
                if (row, col) in merged_map:
                    rs, cs = merged_map[(row, col)]
                    if rs > 1:
                        attrs += f' rowspan="{rs}"'
                    if cs > 1:
                        attrs += f' colspan="{cs}"'
                if style:
                    attrs += f' style="{style}"'

                html_parts.append(f"<td{attrs}>{val_str}</td>")
            html_parts.append("</tr>")

        html_parts.append("</table>")

    html_parts.append("</body></html>")
    return "\n".join(html_parts)


# ── Main engine entry (mutates workbook in-place) ──

def _process_workbook(wb: Workbook, context: Dict[str, Any]) -> None:
    """Adjust the workbook in-place for web display.

    Steps per sheet (skips "Data" sheets):
      1. Adjust day columns to match month length
      2. Update date row to target month
      3. Detect roster regions
      4. Fetch API data for the shift
      5. Shrink regions to match actual headcount (delete from bottom)
      6. Fill retained rows with API data
      7. Normalise number formats
    """
    month = context["month"]
    year, month_num = int(month[:4]), int(month[5:7])
    days_in_month = calendar.monthrange(year, month_num)[1]

    for ws in wb.worksheets:
        if ws.title == "Data":
            continue

        # Phase 1: Adjust day columns and dates
        _adjust_day_columns(ws, days_in_month)
        _update_dates(ws, month)

        # Phase 2: Detect roster regions
        regions = _detect_roster_regions(ws)
        if not regions:
            continue

        # Phase 3: Fetch data for this shift
        shift = _shift_from_sheet_name(ws.title)
        try:
            data = _fetch_data(context, shift)
        except Exception:
            # DB / API not available — leave template data as-is
            continue

        if not data.get("has_data"):
            # No data — clear all data rows
            for region in regions:
                for r in range(region["data_start"], region["data_end"] + 1):
                    for col in range(1, ws.max_column + 1):
                        ws.cell(r, col).value = ""
            continue

        # Match API segments to regions by index
        segments = data.get("segments", [])
        api_by_index: Dict[int, Dict[str, Any]] = {}
        for i, seg in enumerate(segments):
            api_by_index[i + 1] = seg

        # Phase 4-6: Process regions bottom-up (shrink → re-detect → fill).
        for region in sorted(regions, key=lambda r: r["data_start"], reverse=True):
            idx = region["index"]
            seg = api_by_index.get(idx)
            if not seg:
                continue
            rows = seg.get("rows", [])
            target = len(rows)
            _shrink_region(ws, region, target)

            # Re-detect to get updated row numbers after this shrink
            fresh_regions = _detect_roster_regions(ws)
            fresh_region = next((r for r in fresh_regions if r["index"] == idx), None)
            if not fresh_region:
                continue

            # Phase 5: Fill data for this region
            date_row, day_start_col = _find_date_start(ws)
            for i, row_data in enumerate(rows):
                row = fresh_region["data_start"] + i
                _fill_data_row(ws, row, row_data, day_start_col, days_in_month)

        # Phase 6: Adjust number formats
        date_row, day_start_col = _find_date_start(ws)
        if day_start_col:
            _adjust_number_formats(ws, day_start_col, days_in_month)


# ── Public entry point (registered in HTML_ENGINES) ──

def run(
    shift: str,
    project_id: Optional[int] = None,
    category_id: Optional[int] = None,
    month: Optional[str] = None,
) -> Dict[str, Any]:
    """Load template, process workbook, return HTML.

    Args:
        shift: Ignored — web_roster processes all sheets in the workbook.
        project_id: Project ID for template lookup.
        category_id: Category ID for template lookup.
        month: Target month (YYYY-MM).

    Returns:
        {"html": <HTML string>, "month": month, "template_path": path}
    """
    effective_month = month or "2026-03"

    # Try to resolve template from DB
    excel_path, rel_path = resolve_template(
        project_id, category_id, effective_month, "cleaning_shortfall_v2"
    )
    if not excel_path:
        excel_path, rel_path = resolve_template(
            project_id, category_id, effective_month, "roster_shortfall"
        )

    # Fallback: local templates directory (for standalone testing)
    if not excel_path:
        tpl_dir = Path(__file__).resolve().parent.parent.parent / "engine" / "templates"
        candidates = [
            tpl_dir / "TY-2026.03- SG_SEC-Deploy Roster  Shortfall - template.xlsx",
            tpl_dir / "东汇-保安轮休表template.xlsx",
            tpl_dir / "东汇-保洁轮休表_template.xlsx",
        ]
        for c in candidates:
            if c.exists():
                excel_path = c
                rel_path = str(c.name)
                break

    if not excel_path or not excel_path.exists():
        return {
            "html": "<p>No template found</p>",
            "month": effective_month,
            "error": "template_not_found",
        }

    # Load workbook twice:
    #   - data_only=False  → styles, formulas, merged cells, structure
    #   - data_only=True   → cached formula values (computed by Excel)
    wb = load_workbook(str(excel_path), data_only=False, keep_links=False)
    wb_values = load_workbook(str(excel_path), data_only=True, keep_links=False)

    context = {
        "project_id": project_id,
        "category_id": category_id,
        "month": effective_month,
    }

    _process_workbook(wb, context)

    # Keep wb_values' column structure in sync with wb so that cell
    # coordinates line up when we read computed values in workbook_to_html.
    year, month_num = int(effective_month[:4]), int(effective_month[5:7])
    days_in_month = calendar.monthrange(year, month_num)[1]
    for vws in wb_values.worksheets:
        if vws.title == "Data":
            continue
        _adjust_day_columns(vws, days_in_month)
        _update_dates(vws, effective_month)

    html = workbook_to_html(wb, effective_month, value_wb=wb_values)

    return {
        "html": html,
        "month": effective_month,
        "template_path": str(rel_path) if rel_path else "",
        "_wb": wb,           # internal: processed workbook for testing
        "_wb_values": wb_values,  # internal: value workbook for testing
    }
