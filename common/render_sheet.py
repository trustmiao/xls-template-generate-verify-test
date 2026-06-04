"""Map database pages/entries onto a parsed Shortfall template."""
from __future__ import annotations

import calendar
from copy import deepcopy
from dataclasses import dataclass, field
from datetime import date
from typing import Any, Dict, List, Optional

from .template_parser import TemplateStructure, TemplateRow


@dataclass
class RenderedCell:
    day: int
    code: str
    edited: bool = False
    value: Optional[str] = None


@dataclass
class RenderedRow:
    page_id: Optional[int]
    rank_seq: str
    employee_no: str
    name: str
    rank_code: str
    position_text: str
    shift_label: str
    workdays: int
    rest_days: int
    total_hours: int
    user_verified: bool
    is_part_time: bool
    is_supervisor: bool
    segment: str
    cells: List[RenderedCell]
    is_missing: bool = False
    contracted_hours: int = 8


@dataclass
class RenderedSegment:
    title: str
    type: str
    rows: List[RenderedRow] = field(default_factory=list)
    summary: Dict[str, Any] = field(default_factory=dict)


@dataclass
class RenderedSheet:
    shift: str
    label: str
    title: str
    month: str
    year: int
    month_num: int
    days_in_month: int
    weekday_strip: List[str]
    holiday_strip: List[Optional[str]]
    sundays: int
    contract_required: int
    segments: List[RenderedSegment]
    extra_rows: List[RenderedRow]
    template_path: str


# Reuse shortfall logic
from .shift_info import (
    SHIFT_INFO,
    SHIFT_REQUIRED,
    WEEKDAY_CH,
    _build_remark_map,
    _cell_code,
    _get_holidays_for_month,
    _weekday_strip,
)


def _normalize_name(v: str) -> str:
    return v.strip().replace(" ", "").replace("\u3000", "")


def _normalize_emp(v: str) -> str:
    return v.strip().lstrip("0")


def _resolve_month(pages: List[Dict[str, Any]]) -> str:
    for p in pages:
        m = (p.get("month") or "").strip()
        if m:
            return m
    return "2026-03"


def _build_pages_lookup(pages: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    """Build lookup by emp_no and by name."""
    by_emp: Dict[str, Dict[str, Any]] = {}
    by_name: Dict[str, Dict[str, Any]] = {}
    for p in pages:
        emp = _normalize_emp(p.get("employee_no") or "")
        name = _normalize_name(p.get("name") or "")
        if emp:
            by_emp[emp] = p
        if name:
            by_name[name] = p
    return by_emp, by_name


def _build_original_cells(page: Dict[str, Any], days_in_month: int) -> Dict[int, str]:
    days_map = {e["day"]: e for e in page.get("entries", [])}
    result: Dict[int, str] = {}
    for d in range(1, days_in_month + 1):
        e = days_map.get(d)
        if not e:
            result[d] = ""
            continue
        result[d] = _cell_code(e)
    return result


def _make_rendered_row(
    tpl: TemplateRow,
    page: Optional[Dict[str, Any]],
    cells: List[RenderedCell],
    segment_type: str,
    days_in_month: int,
    is_missing: bool = False,
    has_supervisor: bool = False,
) -> RenderedRow:
    workdays = sum(1 for c in cells if c.code == "8")
    rest_days = sum(1 for c in cells if c.code == "RD")
    if page:
        rank_code = page.get("rank_code") or tpl.rank_code or ""
        position_text = page.get("position_text") or tpl.position_text or ""
        ratio = float(page.get("project_time_ratio") or 1.0)
        contracted_hours = int(round(ratio * 8))
        # 所有时间统一 cap 到 8 小时
        is_pt = bool(page.get("is_part_time")) or ratio < 1.0
        if not has_supervisor:
            # Sum only THIS segment's cells. When the same page appears in
            # multiple segments (e.g. a worker counted as 清潔科文 some days
            # and 清潔工人 other days), each segment must report only its own
            # hours — using raw page.entries would double-count and produce
            # 200h for a person with workdays=2 in one segment.
            def _cell_hours(c) -> float:
                if c.value not in (None, ""):
                    try:
                        return min(float(c.value), 8.0)
                    except (TypeError, ValueError):
                        return 0.0
                return 8.0 if c.code == "8" else 0.0
            total_hours = sum(_cell_hours(c) for c in cells)
        else:
            total_hours = workdays * 8 if is_pt else workdays * 8
        return RenderedRow(
            page_id=page["id"],
            rank_seq=tpl.rank_seq,
            employee_no=(page.get("employee_no") or "").zfill(7) if page.get("employee_no") else "",
            name=page.get("name") or "",
            rank_code=rank_code or position_text,
            position_text=position_text,
            shift_label=tpl.shift_label or SHIFT_INFO.get("", {}).get("label", ""),
            workdays=workdays,
            rest_days=rest_days,
            total_hours=total_hours,
            user_verified=bool(page.get("user_verified")),
            is_part_time=bool(page.get("is_part_time")),
            is_supervisor=segment_type == "supervisor",
            segment=segment_type,
            cells=cells,
            is_missing=is_missing,
            contracted_hours=contracted_hours,
        )
    else:
        return RenderedRow(
            page_id=None,
            rank_seq=tpl.rank_seq,
            employee_no=tpl.employee_no.zfill(7) if tpl.employee_no else "",
            name=tpl.name,
            rank_code=tpl.rank_code,
            position_text=tpl.position_text,
            shift_label=tpl.shift_label,
            workdays=workdays,
            rest_days=rest_days,
            total_hours=workdays * 8,
            user_verified=False,
            is_part_time=False,
            is_supervisor=segment_type == "supervisor",
            segment=segment_type,
            cells=cells,
            is_missing=is_missing,
            contracted_hours=8,
        )


def _rank(p: Dict[str, Any]) -> str:
    """Normalized OCR-original rank_code (case/whitespace-insensitive).

    Writes are normalized in update_page_fields; this is a defensive read.
    """
    return (p.get("rank_code") or "").strip().upper()


def _cp(p: Dict[str, Any]) -> str:
    """Normalized cross-posted rank ('GS'/'CF'/'VO' or '').

    Cross-posting decouples 兼岗 from the OCR-original rank_code so the
    person's original role is preserved while they're "borrowed" to cover
    another role. NULL/empty means "not cross-posted". See plan + memory:
    shortfall-no-template-people.
    """
    return (p.get("cross_post_rank") or "").strip().upper()


def _is_cross_posted(p: Dict[str, Any], role: Optional[str] = None) -> bool:
    cpr = _cp(p)
    if not cpr:
        return False
    return role is None or cpr == role


def _compute_role_call(
    role_pages: List[Dict[str, Any]],
    original_cells_map: Dict[int, Dict[int, str]],
    days_in_month: int,
    required_daily: int,
) -> Dict[int, set]:
    """For each day, compute which part-time role members are called (person-count)."""
    fulltime = [p for p in role_pages if not p.get("is_part_time")]
    parttime = [p for p in role_pages if p.get("is_part_time")]

    called_by_day: Dict[int, set] = {}
    for d in range(1, days_in_month + 1):
        ft_working = sum(
            1
            for p in fulltime
            if original_cells_map.get(p["id"], {}).get(d) == "8"
        )
        needed = max(0, required_daily - ft_working)
        available = sorted(
            [p for p in parttime if original_cells_map.get(p["id"], {}).get(d) == "8"],
            key=lambda p: p.get("employee_no") or "",
        )
        called_by_day[d] = {p["id"] for p in available[:needed]}
    return called_by_day


def _compute_role_hours(
    role_pages: List[Dict[str, Any]],
    original_cells_map: Dict[int, Dict[int, str]],
    work_hours_map: Dict[int, Dict[int, float]],
    days_in_month: int,
    required_daily_hours: float,
    cap_per_person: float = 8.0,
) -> Dict[int, Dict[int, float]]:
    """Full-time fills first (capped at cap_per_person), gap filled by part-time sequentially.
    Returns {day: {page_id: supervisor_hours}}."""
    fulltime = [p for p in role_pages if not p.get("is_part_time")]
    parttime = sorted(
        [p for p in role_pages if p.get("is_part_time")],
        key=lambda p: p.get("employee_no") or "",
    )

    result: Dict[int, Dict[int, float]] = {}
    for d in range(1, days_in_month + 1):
        ft_hours = 0.0
        for p in fulltime:
            if original_cells_map.get(p["id"], {}).get(d) == "8":
                actual = work_hours_map.get(p["id"], {}).get(d, cap_per_person)
                ft_hours += min(actual, cap_per_person)

        gap = max(0.0, required_daily_hours - ft_hours)

        result[d] = {}
        for p in parttime:
            if gap <= 0:
                break
            if original_cells_map.get(p["id"], {}).get(d) != "8":
                continue
            actual = work_hours_map.get(p["id"], {}).get(d, cap_per_person)
            can_do = min(actual, gap)
            result[d][p["id"]] = can_do
            gap -= can_do

    return result


def _compute_call_ids(
    pages: List[Dict[str, Any]],
    original_cells_map: Dict[int, Dict[int, str]],
    work_hours_map: Dict[int, Dict[int, float]],
    days_in_month: int,
    req: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Return {role_code: {day: set_of_page_ids}} for person-count roles,
    or {role_code: {day: {page_id: hours}}} for hour-based roles."""
    result: Dict[str, Any] = {}

    # GS roster = FT (rank='GS' AND not cross-posted) ∪ cross-posted-to-GS.
    # For _compute_role_call's purpose, "FT" is the page without cross-posting
    # (its actual rank), and "PT" is the cross-posted contributor. The function
    # reads `is_part_time` to split; we ferry the same intent via a temporary
    # overlay (set is_part_time=True for cross-posted).
    def _overlay_pt(plist):
        out = []
        for p in plist:
            if _is_cross_posted(p):
                mp = dict(p); mp["is_part_time"] = True; out.append(mp)
            else:
                out.append(p)
        return out

    gs_pages = _overlay_pt([p for p in pages
                            if _rank(p) == "GS" or _is_cross_posted(p, "GS")])
    if gs_pages:
        gs_needed = (req or {}).get("supervisor", 2)
        result["GS"] = _compute_role_call(gs_pages, original_cells_map, days_in_month, gs_needed)

    # CF roster = FT (rank='CF' not cross-posted) ∪ cross-posted-to-CF
    cf_pages = _overlay_pt([p for p in pages
                            if _rank(p) == "CF" or _is_cross_posted(p, "CF")])
    pt_cf = [p for p in cf_pages if p.get("is_part_time")]
    if pt_cf and req and req.get("supervisor", 0) > 0:
        cf_needed = req["supervisor"]
        result["CF"] = _compute_role_call(cf_pages, original_cells_map, days_in_month, cf_needed)

    # VO roster = FT (rank='VO' not cross-posted) ∪ cross-posted-to-VO
    vo_pages = _overlay_pt([p for p in pages
                            if _rank(p) == "VO" or _is_cross_posted(p, "VO")])
    pt_vo = [p for p in vo_pages if p.get("is_part_time")]
    if pt_vo and req and req.get("ordinary_weekly_hours", 0) > 0:
        daily_hours = req["ordinary_weekly_hours"] / 7.0
        result["VO"] = _compute_role_hours(vo_pages, original_cells_map, work_hours_map, days_in_month, daily_hours)

    return result


def _is_remark(code: str) -> bool:
    """Check if the code is a remark-type code (leave, holiday, etc.)."""
    return code in ("FS", "FH", "SL", "W", "WTT", "RP", "NP")


def _segment_cells(
    page: Dict[str, Any],
    original_cells: Dict[int, str],
    segment_type: str,
    called_by_day: Dict[str, Any],
    days_in_month: int,
    seg_title: str = "",
    work_hours_map: Optional[Dict[int, Dict[int, float]]] = None,
) -> List[RenderedCell]:
    cells: List[RenderedCell] = []
    # "is_pt" / "is_<role>" now mean "cross-posted to that role" — used to
    # decide cross-segment cell coding (e.g. PT GS day cell shows '8' in
    # supervisor seg when called, 'G' otherwise). The OCR-original rank_code
    # is irrelevant here since the segment we're rendering decides the role.
    is_pt = _is_cross_posted(page)
    is_gs = _is_cross_posted(page, "GS")
    is_cf = _is_cross_posted(page, "CF")
    is_vo = _is_cross_posted(page, "VO")

    for d in range(1, days_in_month + 1):
        orig = original_cells.get(d, "")

        # GS part-time logic
        if is_gs and is_pt:
            gs_calls = called_by_day.get("GS", {})
            if segment_type == "supervisor":
                if orig in ("RD",) or _is_remark(orig):
                    cells.append(RenderedCell(day=d, code=orig, edited=False))
                elif orig == "8":
                    if page["id"] in gs_calls.get(d, set()):
                        cells.append(RenderedCell(day=d, code="8", edited=False))
                    else:
                        cells.append(RenderedCell(day=d, code="G", edited=False))
                else:
                    cells.append(RenderedCell(day=d, code=orig, edited=False))
            else:  # security segment
                if orig in ("RD",) or _is_remark(orig):
                    cells.append(RenderedCell(day=d, code=orig, edited=False))
                elif orig == "8":
                    if page["id"] in gs_calls.get(d, set()):
                        cells.append(RenderedCell(day=d, code="GS", edited=False))
                    else:
                        cells.append(RenderedCell(day=d, code="8", edited=False))
                else:
                    cells.append(RenderedCell(day=d, code=orig, edited=False))

        # CF part-time logic (清洁科文兼岗 — person-count based)
        elif is_cf and is_pt:
            cf_calls = called_by_day.get("CF", {})
            is_cf_seg = "科文" in seg_title
            if orig in ("RD",) or _is_remark(orig):
                cells.append(RenderedCell(day=d, code=orig, edited=False))
            elif orig == "8":
                if is_cf_seg:
                    code = "8" if page["id"] in cf_calls.get(d, set()) else "C"
                else:
                    code = "CF" if page["id"] in cf_calls.get(d, set()) else "8"
                cells.append(RenderedCell(day=d, code=code, edited=False))
            else:
                cells.append(RenderedCell(day=d, code=orig, edited=False))

        # VO part-time logic (hour-based)
        elif is_vo and is_pt:
            vo_hours = called_by_day.get("VO", {})
            hours_as_vo = vo_hours.get(d, {}).get(page["id"], 0.0)
            is_called = hours_as_vo > 0
            is_vo_seg = "VO" in seg_title
            if orig in ("RD",) or _is_remark(orig):
                cells.append(RenderedCell(day=d, code=orig, edited=False))
            elif orig == "8":
                if is_vo_seg:
                    code = "8" if is_called else "CW"
                    value = None
                    if is_called and hours_as_vo < 8.0:
                        value = f"{hours_as_vo:.2f}".rstrip("0").rstrip(".")
                else:
                    # 工人段：全部8小时出去兼职才显示VO，否则显示剩余时间
                    if is_called:
                        actual = work_hours_map.get(page["id"], {}).get(d, 8.0)
                        ratio = float(page.get("project_time_ratio") or 1.0)
                        available = min(actual, ratio * 8.0)
                        remaining = available - hours_as_vo
                        if remaining <= 0:
                            code = "VO"
                            value = None
                        else:
                            code = "8"
                            value = f"{remaining:.2f}".rstrip("0").rstrip(".")
                    else:
                        code = "8"
                        value = None
                cells.append(RenderedCell(day=d, code=code, edited=False, value=value))
            else:
                cells.append(RenderedCell(day=d, code=orig, edited=False))

        # Full-time GS/CF/VO or other ranks: original code
        else:
            cells.append(RenderedCell(day=d, code=orig, edited=False))

    # 全局：code="8" 时如果实际工时 < 8，显示真实时间（work_hours_map 已截断到 8）
    if work_hours_map:
        for c in cells:
            if c.code == "8" and c.value is None:
                wh = work_hours_map.get(page["id"], {}).get(c.day, 8.0)
                if wh < 8.0:
                    c.value = f"{wh:.2f}".rstrip("0").rstrip(".")

    return cells


def _daily_counts(rows: List[RenderedRow], code: str = "8") -> List[int]:
    return [
        sum(1 for r in rows if r.cells[d - 1].code == code)
        for d in range(1, len(rows[0].cells) + 1 if rows else 1)
    ]


def _week_bounds(days_in_month: int) -> List[Tuple[int, int]]:
    """Return list of (start_day, end_day) for each week in the month."""
    weeks: List[Tuple[int, int]] = []
    day = 1
    while day <= days_in_month:
        end = min(day + 6, days_in_month)
        weeks.append((day, end))
        day = end + 1
    return weeks


def _allocate_hours(
    pages: List[Dict[str, Any]],
    work_hours_map: Dict[int, Dict[int, float]],
    days_in_month: int,
    req: Dict[str, int],
    called_by_day: Optional[Dict[str, Dict[int, set]]] = None,
) -> Dict[str, List[float]]:
    """Return allocated hours per segment per day for cleaning templates."""
    vo_weekly = req.get("ordinary_weekly_hours", 0)
    vo_daily = vo_weekly / 7.0 if vo_weekly > 0 else req.get("ordinary", 0) * 5.25
    daily_req = {
        "supervisor": req.get("supervisor", 0) * 8,
        "ordinary": vo_daily,
        "all": req.get("all", req.get("contract", 0)) * 8,
    }

    result: Dict[str, List[float]] = {
        "supervisor": [0.0] * days_in_month,
        "ordinary": [0.0] * days_in_month,
        "all": [0.0] * days_in_month,
        "overflow": [0.0] * days_in_month,
    }

    cf_calls = (called_by_day or {}).get("CF", {})
    vo_hours = (called_by_day or {}).get("VO", {})

    for p in pages:
        page_id = p["id"]
        rc = (p.get("rank_code") or "").strip().upper()
        ratio = float(p.get("project_time_ratio") or 1.0)
        is_pt = bool(p.get("is_part_time"))
        is_part_time = is_pt or ratio < 1.0

        wh_map = work_hours_map.get(page_id, {})
        for d in range(1, days_in_month + 1):
            raw = wh_map.get(d, 0.0)
            available = raw  # work_hours_map 已全局截断到 8

            if rc == "CF":
                if is_pt:
                    # Part-time CF: person-count based split
                    if page_id in cf_calls.get(d, set()):
                        result["supervisor"][d - 1] += available
                    else:
                        result["all"][d - 1] += available
                else:
                    # Full-time CF: all to supervisor
                    to_seg = min(available, daily_req["supervisor"])
                    result["supervisor"][d - 1] += to_seg
                    result["overflow"][d - 1] += available - to_seg
            elif rc == "VO":
                if is_pt:
                    # Part-time VO: hour-based split
                    hours_as_vo = vo_hours.get(d, {}).get(page_id, 0.0)
                    result["ordinary"][d - 1] += hours_as_vo
                    result["all"][d - 1] += (available - hours_as_vo)
                else:
                    # Full-time VO: all to ordinary
                    to_seg = min(available, daily_req["ordinary"])
                    result["ordinary"][d - 1] += to_seg
                    result["overflow"][d - 1] += available - to_seg
            else:
                # CW or null: all to 清潔工人段
                result["all"][d - 1] += available

    return result


def render_sheet(
    tpl: TemplateStructure,
    pages: List[Dict[str, Any]],
    shift: str,
    month: Optional[str] = None,
    template_path: str = "",
) -> RenderedSheet:
    info = SHIFT_INFO[shift]
    resolved_month = month if month else _resolve_month(pages)
    try:
        y, m = resolved_month.split("-")
        year, month_num = int(y), int(m)
    except Exception:
        year, month_num = 2026, 3

    days_in_month = calendar.monthrange(year, month_num)[1]
    weekday_strip = _weekday_strip(year, month_num, days_in_month)
    holiday_map = _get_holidays_for_month(year, month_num)
    holiday_strip = [holiday_map.get(d, None) for d in range(1, days_in_month + 1)]
    sundays = sum(1 for d in range(1, days_in_month + 1) if date(year, month_num, d).weekday() == 6)

    req = tpl.required_counts if tpl.required_counts else {}

    # Build lookups
    by_emp, by_name = _build_pages_lookup(pages)
    original_cells_map: Dict[int, Dict[int, str]] = {}
    for p in pages:
        original_cells_map[p["id"]] = _build_original_cells(p, days_in_month)

    # work_hours_map for shortfall is hour-sourced from entries.adjusted_hours
    # (HK gov-rounded) when present, falling back to work_hours (= Total Actual
    # Hours) otherwise. Per user spec: 工时检查 uses work_hours; shortfall +
    # deployment use adjusted_hours. See excel_import / e_face_daily parser.
    work_hours_map: Dict[int, Dict[int, float]] = {}
    for p in pages:
        days_map = {e["day"]: e for e in p.get("entries", [])}
        ratio = float(p.get("project_time_ratio") or 1.0)
        wh: Dict[int, float] = {}
        for d in range(1, days_in_month + 1):
            e = days_map.get(d)
            if e:
                src = e.get("adjusted_hours")
                if src is None:
                    src = e.get("work_hours")
                raw = min(float(src or 0), 8.0)
            else:
                raw = 0.0
            wh[d] = min(raw, ratio * 8.0)
        work_hours_map[p["id"]] = wh

    called_by_day = _compute_call_ids(pages, original_cells_map, work_hours_map, days_in_month, req)

    matched_page_ids: set[int] = set()
    segments: List[RenderedSegment] = []
    has_supervisor = any(s.type == "supervisor" for s in tpl.segments)

    # DB-derived rosters per role. PER USER RULE: the rendered person list is
    # OCR-derived only (memory:shortfall-no-template-people). 'FT in role' =
    # OCR rank matches AND not cross-posted to another role; 'cross-posted' =
    # explicit DB tag on the new cross_post_rank column.
    def _sort_idx(p):
        return p.get("page_index") or 0

    ft_gs_p = sorted([p for p in pages
                      if _rank(p) == "GS" and not _is_cross_posted(p)], key=_sort_idx)
    cp_gs_p = sorted([p for p in pages if _is_cross_posted(p, "GS")], key=_sort_idx)
    ft_cf_p = sorted([p for p in pages
                      if _rank(p) == "CF" and not _is_cross_posted(p)], key=_sort_idx)
    cp_cf_p = sorted([p for p in pages if _is_cross_posted(p, "CF")], key=_sort_idx)
    ft_vo_p = sorted([p for p in pages
                      if _rank(p) == "VO" and not _is_cross_posted(p)], key=_sort_idx)
    cp_vo_p = sorted([p for p in pages if _is_cross_posted(p, "VO")], key=_sort_idx)
    ft_gs_ids = {p["id"] for p in ft_gs_p}
    ft_cf_ids = {p["id"] for p in ft_cf_p}
    ft_vo_ids = {p["id"] for p in ft_vo_p}

    def _roster_for(seg) -> List[Dict[str, Any]]:
        """DB-driven roster for this segment, ordered per user rule:
          - supervisor / CF / VO segments: FT first, cross-posted at BOTTOM
          - security / 清潔工人 segments: cross-posted at TOP, then FT
        """
        if seg.type == "supervisor":
            return ft_gs_p + cp_gs_p   # PT at bottom
        if not has_supervisor:
            # cleaning template — distinguish by title
            if "科文" in seg.title:
                return ft_cf_p + cp_cf_p
            if "VO" in seg.title:
                return ft_vo_p + cp_vo_p
            if "清潔工人" in seg.title:
                # PT (cross-posted to CF or VO) at top, then ordinary cleaners
                ord_workers = sorted(
                    [p for p in pages
                     if p["id"] not in ft_cf_ids and p["id"] not in ft_vo_ids
                     and not _is_cross_posted(p)],
                    key=_sort_idx,
                )
                return cp_cf_p + cp_vo_p + ord_workers
            return []
        # security segment in a security template: cross-posted GS at top,
        # then all guards (excluding FT GS who live in supervisor)
        rest = sorted(
            [p for p in pages
             if p["id"] not in ft_gs_ids and not _is_cross_posted(p, "GS")],
            key=_sort_idx,
        )
        return cp_gs_p + rest

    # Style-template lookup: for each segment, build a map page_id → template
    # row that matched it via emp/name (for layout details like rank_seq).
    # Pages without a template match still render — they reuse the first
    # template data row as their style scaffold.
    for seg in tpl.segments:
        seg_rows: List[RenderedRow] = []

        # Pre-build template-row lookup for FT styling
        tr_by_page: Dict[int, "TemplateRow"] = {}  # noqa: F821
        for tr in seg.rows:
            emp_key = _normalize_emp(tr.employee_no)
            name_key = _normalize_name(tr.name)
            p = (by_emp.get(emp_key) if emp_key else None) or \
                (by_name.get(name_key) if name_key else None)
            if p:
                tr_by_page[p["id"]] = tr

        default_tr = seg.rows[0] if seg.rows else None
        roster = _roster_for(seg)

        for page in roster:
            tr = tr_by_page.get(page["id"], default_tr)
            if tr is None:
                continue
            matched_page_ids.add(page["id"])
            orig = original_cells_map[page["id"]]
            cells = _segment_cells(
                page, orig, seg.type, called_by_day, days_in_month,
                seg_title=seg.title, work_hours_map=work_hours_map,
            )
            row = _make_rendered_row(
                tr, page, cells, seg.type, days_in_month,
                is_missing=False, has_supervisor=has_supervisor,
            )
            # is_part_time on the rendered row signals "cross-posted" for
            # the frontend PT tint. Use the new column directly.
            if _is_cross_posted(page):
                row.is_part_time = True
                # When the row is rendered in the segment matching the cross-
                # post role, override the rank label to that role. (Some
                # templates list the cross-posted person in their primary-role
                # segment but with a non-matching rank code — e.g. 东汇 A
                # supervisor row 3 has 成小紅 with rank='SG/G'; we want her
                # row to read 'GS' there because that's the role she covers.)
                cp_role = _cp(page)
                if cp_role:
                    if seg.type == "supervisor" and cp_role == "GS":
                        row.rank_code = "GS"
                    elif not has_supervisor:
                        if "科文" in seg.title and cp_role == "CF":
                            row.rank_code = "CF"
                        elif "VO" in seg.title and cp_role == "VO":
                            row.rank_code = "VO"
            # PT CF per-segment hours fixup (person-count based)
            if not has_supervisor and _is_cross_posted(page, "CF"):
                if "科文" in seg.title:
                    row.total_hours = row.workdays * row.contracted_hours
                else:
                    row.total_hours = row.workdays * 8
            # PT VO per-segment hours fixup (hour-based)
            if not has_supervisor and _is_cross_posted(page, "VO"):
                vo_hours_map = called_by_day.get("VO", {})
                if "VO" in seg.title:
                    vo_days = sum(
                        1 for d in range(1, days_in_month + 1)
                        if vo_hours_map.get(d, {}).get(page["id"], 0.0) > 0
                    )
                    row.workdays = vo_days
                    row.total_hours = sum(
                        vo_hours_map.get(d, {}).get(page["id"], 0.0)
                        for d in range(1, days_in_month + 1)
                    )
                else:
                    worker_days = sum(
                        1 for d in range(1, days_in_month + 1)
                        if original_cells_map.get(page["id"], {}).get(d) == "8"
                        and vo_hours_map.get(d, {}).get(page["id"], 0.0) == 0
                    )
                    row.workdays = worker_days
                    row.total_hours = worker_days * 8
            seg_rows.append(row)

        # Compute segment summary
        if seg.type == "supervisor":
            actual = _daily_counts(seg_rows, "8")
            summary = {
                "actual_count": actual,
                "required_count": req.get("supervisor", 0),
            }
        elif has_supervisor:
            # Security segment in a template that also has a supervisor segment
            # (security guard template). Use temporary structure; post-process will fill special/ordinary/all.
            contract_count = _daily_counts(seg_rows, "8")
            summary = {
                "contract_count": contract_count,
                "required_special": req.get("special", 0),
                "required_ordinary": req.get("ordinary", 0),
                "required_all": req.get("all", 0),
            }
        else:
            # Security segment in a template with NO supervisor segment
            # (cleaning template). Use simple segment-level summary.
            actual = _daily_counts(seg_rows, "8")
            # Infer required key from segment title for multi-segment cleaning templates
            if "科文" in seg.title:
                required = req.get("supervisor", 0)
            elif "VO" in seg.title:
                required = req.get("ordinary", 0)
            else:
                required = req.get("all", req.get("contract", 0))
            summary = {
                "actual_count": actual,
                "required_count": required,
            }

        segments.append(RenderedSegment(title=seg.title, type=seg.type, rows=seg_rows, summary=summary))

    # If no segments were parsed, fall back to DB-driven segments (same as current shortfall.py)
    if not segments:
        all_people = []
        for idx, page in enumerate(pages, start=1):
            orig = original_cells_map[page["id"]]
            cells = [RenderedCell(day=d, code=orig.get(d, ""), edited=False) for d in range(1, days_in_month + 1)]
            workdays = sum(1 for c in cells if c.code == "8")
            ratio = float(page.get("project_time_ratio") or 1.0)
            contracted_hours = int(round(ratio * 8))
            all_people.append(RenderedRow(
                page_id=page["id"],
                rank_seq=f"{info['prefix']}{idx:02d}",
                employee_no=(page.get("employee_no") or "").zfill(7) if page.get("employee_no") else "",
                name=page.get("name") or "",
                rank_code=page.get("rank_code") or "",
                position_text=page.get("position_text") or "",
                shift_label=info["label"],
                workdays=workdays,
                rest_days=sum(1 for c in cells if c.code == "RD"),
                total_hours=workdays * 8,
                user_verified=bool(page.get("user_verified")),
                is_part_time=bool(page.get("is_part_time")),
                is_supervisor=_rank(page) == "GS",
                segment="security",
                cells=cells,
                contracted_hours=contracted_hours,
            ))
        segments.append(RenderedSegment(
            title="2. 清潔工人出勤時數",
            type="security",
            rows=all_people,
            summary={},
        ))
        matched_page_ids = {p["id"] for p in pages}

    # Post-process: fill security summary using supervisor actual
    # Only applies to security-guard templates that have both supervisor + security segments.
    if has_supervisor:
        supervisor_actual: Optional[List[int]] = None
        for seg in segments:
            if seg.type == "supervisor":
                supervisor_actual = seg.summary.get("actual_count", [0] * days_in_month)
                break

        # Defensive: when supervisor segment is empty (e.g. no FT GS yet AND
        # no cross-posted-to-GS), _daily_counts returns [], which would index-
        # error in the loop below. Pad to days_in_month.
        if not supervisor_actual or len(supervisor_actual) != days_in_month:
            supervisor_actual = [0] * days_in_month

        for seg in segments:
            if seg.type == "security":
                contract_count = seg.summary.get("contract_count", [0] * days_in_month)
                if not contract_count or len(contract_count) != days_in_month:
                    contract_count = [0] * days_in_month
                # "特别保安员每日(实际)" — semantically this row is fixed at
                # the contract-required headcount, not a count of OCR-clocked
                # 8s. Every template (东汇 row 33 hardcoded 6, 大元 row 62
                # hardcoded 2, etc.) writes the required value verbatim,
                # because the data model never reliably distinguishes
                # special vs ordinary at the OCR layer (rank_code comes back
                # as "SG/G" unsplit). Earlier this was set to
                # supervisor_actual, which matched the 大元 template by
                # coincidence (supervisor required==special required==2) and
                # was off by 4/day on 东汇.
                special_required = req.get("special", 0)
                special_actual = [special_required] * days_in_month
                ordinary_actual = [
                    contract_count[d - 1] - special_actual[d - 1]
                    for d in range(1, days_in_month + 1)
                ]
                all_actual = [
                    supervisor_actual[d - 1] + contract_count[d - 1]
                    for d in range(1, days_in_month + 1)
                ]
                # "all" required_count = sum of every sub-category's
                # contractual requirement (supervisor + special + ordinary).
                # This matches the template's formula
                #   =SUM(I16 + I63 + I66)
                # on the 保安每日工作(要求)總人數 row. Previously we read
                # req.get("all", 0) but parser never populated req["all"] for
                # security templates (the row's day cell is a formula, not
                # a literal number, so the parser's number-gate skipped it
                # and "all" stayed at the default 0). With required=0 the
                # frontend's (要求) row rendered 0 and (缺勤) rendered the
                # full actual, both wrong.
                required_all = (
                    req.get("supervisor", 0)
                    + req.get("special", 0)
                    + req.get("ordinary", 0)
                )
                seg.summary = {
                    "special": {
                        "actual_count": special_actual,
                        "required_count": req.get("special", 0),
                    },
                    "ordinary": {
                        "actual_count": ordinary_actual,
                        "required_count": req.get("ordinary", 0),
                    },
                    "all": {
                        "actual_count": all_actual,
                        "required_count": required_all,
                    },
                }

    # Build cleaning segment index map for rank-based auto-assignment
    cleaning_seg_map: Dict[str, int] = {}
    if not has_supervisor:
        for i, seg in enumerate(segments):
            if "科文" in seg.title:
                cleaning_seg_map["CF"] = i
            elif "VO" in seg.title:
                cleaning_seg_map["VO"] = i
            elif "清潔工人" in seg.title:
                cleaning_seg_map["CW"] = i

    # Generate rank_seq for cleaning-template DB-driven rows (blank template has no data rows)
    _rank_seq_counter = 1
    _page_rank_seq_map: Dict[int, str] = {}

    def _get_rank_seq(page_id: int) -> str:
        nonlocal _rank_seq_counter
        if page_id not in _page_rank_seq_map:
            _page_rank_seq_map[page_id] = f"{shift}{_rank_seq_counter:02d}"
            _rank_seq_counter += 1
        return _page_rank_seq_map[page_id]

    # Extra rows: DB pages not found in template
    extra_rows: List[RenderedRow] = []
    for p in pages:
        if p["id"] not in matched_page_ids:
            orig = original_cells_map[p["id"]]
            # Determine target segment title for CF part-time cell rendering
            rc = (p.get("rank_code") or "").strip().upper()
            target_seg_title = ""
            if not has_supervisor:
                if rc == "CF" and "CF" in cleaning_seg_map:
                    target_seg_title = segments[cleaning_seg_map["CF"]].title
                elif rc == "VO" and "VO" in cleaning_seg_map:
                    target_seg_title = segments[cleaning_seg_map["VO"]].title
                elif "CW" in cleaning_seg_map:
                    target_seg_title = segments[cleaning_seg_map["CW"]].title
            cells = _segment_cells(p, orig, "security", called_by_day, days_in_month, seg_title=target_seg_title, work_hours_map=work_hours_map)
            workdays = sum(1 for c in cells if c.code == "8")
            ratio = float(p.get("project_time_ratio") or 1.0)
            contracted_hours = int(round(ratio * 8))
            # 保洁模板 extra_rows: PT CF / PT VO 按段计算工时
            is_pt = bool(p.get("is_part_time")) or ratio < 1.0
            is_pt_cf = rc == "CF" and bool(p.get("is_part_time")) and not has_supervisor
            is_pt_vo = rc == "VO" and bool(p.get("is_part_time")) and not has_supervisor
            cf_calls = called_by_day.get("CF", {})
            vo_hours_map = called_by_day.get("VO", {})
            wh_map = work_hours_map.get(p["id"], {})
            # Chain: PT-CF and PT-VO each compute per-segment hours; regular
            # workers fall through to the cleaning default (or supervisor
            # default if the template has a supervisor segment).
            if is_pt_cf:
                # Use cells-derived workdays so total_hours stays consistent
                # with the rendered cells (prevents 200h shown for 2-workday
                # CF row, where called_by_day overcounts theoretical assignments).
                if "科文" in target_seg_title:
                    total_hours = workdays * contracted_hours
                else:
                    contracted_hours = 8
                    total_hours = workdays * 8
            elif is_pt_vo:
                if "VO" in target_seg_title:
                    total_hours = sum(vo_hours_map.get(d, {}).get(p["id"], 0.0) for d in range(1, days_in_month + 1))
                else:
                    # Worker segment: work_hours_map - 调配到 VO 的时间
                    total_hours = sum(
                        wh_map.get(d, 0.0) - vo_hours_map.get(d, {}).get(p["id"], 0.0)
                        for d in range(1, days_in_month + 1)
                    )
            elif not has_supervisor:
                total_hours = sum(wh_map.get(d, 0.0) for d in range(1, days_in_month + 1))
            else:
                total_hours = workdays * 8
            row = RenderedRow(
                page_id=p["id"],
                rank_seq=_get_rank_seq(p["id"]) if not has_supervisor else "",
                employee_no=(p.get("employee_no") or "").zfill(7) if p.get("employee_no") else "",
                name=p.get("name") or "",
                rank_code=p.get("rank_code") or "",
                position_text=p.get("position_text") or "",
                shift_label=info["label"],
                workdays=workdays,
                rest_days=sum(1 for c in cells if c.code == "RD"),
                total_hours=total_hours,
                user_verified=bool(p.get("user_verified")),
                is_part_time=bool(p.get("is_part_time")),
                is_supervisor=False,
                segment="security",
                cells=cells,
                contracted_hours=contracted_hours,
            )
            if not has_supervisor:
                rc = (p.get("rank_code") or "").strip().upper()
                is_pt = bool(p.get("is_part_time"))
                if rc == "CF" and "CF" in cleaning_seg_map:
                    segments[cleaning_seg_map["CF"]].rows.append(row)
                    # PT CF also appears in worker segment (兼岗)
                    if is_pt and "CW" in cleaning_seg_map:
                        worker_title = segments[cleaning_seg_map["CW"]].title
                        w_cells = _segment_cells(p, orig, "security", called_by_day, days_in_month, seg_title=worker_title, work_hours_map=work_hours_map)
                        w_workdays = sum(1 for c in w_cells if c.code == "8")
                        worker_days = sum(1 for d in range(1, days_in_month + 1)
                            if original_cells_map.get(p["id"], {}).get(d) == "8"
                            and p["id"] not in cf_calls.get(d, set()))
                        w_total_primary = worker_days * 8
                        w_row = RenderedRow(
                            page_id=p["id"],
                            rank_seq=_get_rank_seq(p["id"]),
                            employee_no=(p.get("employee_no") or "").zfill(7) if p.get("employee_no") else "",
                            name=p.get("name") or "",
                            rank_code=p.get("rank_code") or "",
                            position_text=p.get("position_text") or "",
                            shift_label=info["label"],
                            workdays=w_workdays,
                            rest_days=sum(1 for c in w_cells if c.code == "RD"),
                            total_hours=w_total_primary,
                            user_verified=bool(p.get("user_verified")),
                            is_part_time=bool(p.get("is_part_time")),
                            is_supervisor=False,
                            segment="security",
                            cells=w_cells,
                            contracted_hours=8,
                        )
                        segments[cleaning_seg_map["CW"]].rows.append(w_row)
                elif rc == "VO" and "VO" in cleaning_seg_map:
                    segments[cleaning_seg_map["VO"]].rows.append(row)
                    # 兼职 VO 才在清洁工人段出现（兼岗），全职 VO 只做 VO
                    if is_pt and "CW" in cleaning_seg_map:
                        worker_title = segments[cleaning_seg_map["CW"]].title
                        w_cells = _segment_cells(p, orig, "security", called_by_day, days_in_month, seg_title=worker_title, work_hours_map=work_hours_map)
                        w_workdays = sum(1 for c in w_cells if c.code == "8")
                        w_row = RenderedRow(
                            page_id=p["id"],
                            rank_seq=_get_rank_seq(p["id"]),
                            employee_no=(p.get("employee_no") or "").zfill(7) if p.get("employee_no") else "",
                            name=p.get("name") or "",
                            rank_code=p.get("rank_code") or "",
                            position_text=p.get("position_text") or "",
                            shift_label=info["label"],
                            workdays=w_workdays,
                            rest_days=sum(1 for c in w_cells if c.code == "RD"),
                            total_hours=w_workdays * 8,
                            user_verified=bool(p.get("user_verified")),
                            is_part_time=bool(p.get("is_part_time")),
                            is_supervisor=False,
                            segment="security",
                            cells=w_cells,
                            contracted_hours=8,
                        )
                        segments[cleaning_seg_map["CW"]].rows.append(w_row)
                elif "CW" in cleaning_seg_map:
                    # Default: CW or null/empty
                    segments[cleaning_seg_map["CW"]].rows.append(row)
                else:
                    extra_rows.append(row)
            else:
                extra_rows.append(row)

    # Sort rows within each segment for cleaning templates
    if not has_supervisor:
        for seg in segments:
            if "科文" in seg.title or "VO" in seg.title:
                # 全职在前，兼职在后，缺失行最后
                seg.rows.sort(key=lambda r: (r.is_part_time, r.is_missing))
            elif "清潔工人" in seg.title:
                # 兼职在前，全职在后，缺失行最后
                seg.rows.sort(key=lambda r: (not r.is_part_time, r.is_missing))

    # Recompute segment summaries for cleaning templates
    if not has_supervisor:
        for seg in segments:
            if not seg.rows:
                continue
            seg.summary["actual_count"] = _daily_counts(seg.rows, "8")
            if "科文" in seg.title:
                seg.summary["required_count"] = req.get("supervisor", 0)
            elif "VO" in seg.title:
                seg.summary["required_count"] = req.get("ordinary", 0)
            elif "清潔工人" in seg.title:
                seg.summary["required_count"] = req.get("all", req.get("contract", 0))

    # Post-process: compute daily_hours and weekly_summary for cleaning templates
    if not has_supervisor:
        weeks = _week_bounds(days_in_month)
        hours_alloc = _allocate_hours(pages, work_hours_map, days_in_month, req, called_by_day)
        for seg in segments:
            if not seg.rows:
                continue
            # 所有 cleaning segment（科文/VO/普通清洁工人）都基于渲染后的 cells 重新计算
            # daily_hours，不依赖 rank_code（因为 OCR 录入的 rank_code 常为 null）。
            # 这样 shortfall-engine 和 shortfall-gap 的口径一致。
            if "科文" in seg.title:
                daily_req_val = req.get("supervisor", 0) * 8
            elif "VO" in seg.title:
                vo_weekly = req.get("ordinary_weekly_hours", 0)
                daily_req_val = vo_weekly / 7.0 if vo_weekly > 0 else req.get("ordinary", 0) * 5.25
            else:
                daily_req_val = req.get("all", req.get("contract", 0)) * 8

            daily_hours = [0.0] * days_in_month
            for r in seg.rows:
                for c in r.cells:
                    d = c.day - 1
                    if c.value:
                        daily_hours[d] += float(c.value)
                    elif c.code == "8":
                        daily_hours[d] += 8.0
                    # 其他 code (RD/FH/SL/C/CW/空 等) 不算工时

            seg.summary["daily_hours"] = daily_hours
            weekly_actual = [sum(daily_hours[s - 1:e]) for s, e in weeks]
            weekly_required = [daily_req_val * (e - s + 1) for s, e in weeks]
            weekly_shortfall = [a - r for a, r in zip(weekly_actual, weekly_required)]
            seg.summary["weekly_summary"] = {
                "actual": weekly_actual,
                "required": weekly_required,
                "shortfall": weekly_shortfall,
            }

    # Compute contract_required from available required counts
    contract_required = req.get("contract", 0)
    if not contract_required:
        contract_required = sum(
            v for k, v in req.items() if k in ("supervisor", "special", "ordinary", "all")
        )

    return RenderedSheet(
        shift=shift,
        label=info["label"],
        title=info["title"],
        month=resolved_month,
        year=year,
        month_num=month_num,
        days_in_month=days_in_month,
        weekday_strip=weekday_strip,
        holiday_strip=holiday_strip,
        sundays=sundays,
        contract_required=contract_required,
        segments=segments,
        extra_rows=extra_rows,
        template_path=template_path,
    )
