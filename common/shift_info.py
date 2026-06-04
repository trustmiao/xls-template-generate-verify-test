"""Shared shortfall utilities (extracted from old shortfall.py)."""
from __future__ import annotations

import calendar
import json
import os
import time
from datetime import date
from functools import lru_cache
from typing import Any, Dict, List, Optional

from ...models import list_public_holidays


@lru_cache(maxsize=8)
def _load_wb_cached(abspath: str, mtime: float, data_only: bool):
    """Load a workbook with external-link parsing disabled (keep_links=False).

    keep_links=False skips openpyxl's external-link book parsing, which on
    these templates costs ~6 s per load. Result is cached by (path, mtime,
    data_only) so the 3 per-shift shortfall-engine calls (and the formula /
    data_only variants) reuse one parse instead of reloading the same file.
    Stale entries (changed mtime) fall out via the LRU bound.

    NOTE: returned Workbook is shared — callers MUST treat it as read-only.
    """
    from openpyxl import load_workbook
    return load_workbook(abspath, data_only=data_only, keep_links=False)


def load_template_workbook(path: str, data_only: bool = False):
    """Cached, keep_links=False workbook load for read-only template access."""
    ap = os.path.abspath(path)
    try:
        mtime = os.path.getmtime(ap)
    except OSError:
        mtime = 0.0
    return _load_wb_cached(ap, mtime, bool(data_only))


def clear_template_caches() -> Dict[str, int]:
    """Drop all cached template workbooks + parsed structures.

    The caches are keyed by file mtime, so a re-uploaded template normally
    refreshes automatically. Call this to force a recompute when that isn't
    enough — e.g. a template overwritten in place with an unchanged mtime, or
    after a code change to the parser while the server keeps running.
    Page/entry data is never cached (always read live from the DB), so this
    only affects template parsing.
    """
    from .shortfall_engine.parser import _parse_template_cached  # lazy: avoid import cycle
    evicted = {
        "workbooks": _load_wb_cached.cache_info().currsize,
        "templates": _parse_template_cached.cache_info().currsize,
    }
    _load_wb_cached.cache_clear()
    _parse_template_cached.cache_clear()
    return evicted


# A/B/C → (sheet name, full title, rank seq prefix, English title)
SHIFT_INFO: Dict[str, Dict[str, str]] = {
    "A": {"label": "早", "title": "早更 07:00 - 15:00", "prefix": "M"},
    "B": {"label": "中", "title": "中更 15:00 - 23:00", "prefix": "D"},
    "C": {"label": "夜", "title": "夜更 23:00 - 07:00", "prefix": "N"},
}

# Template fixed required counts (contractual)
# 早/中: supervisor=2, special=2, ordinary=29, all=33, contract=31
# 夜:    supervisor=2, special=2, ordinary=21, all=25, contract=23
SHIFT_REQUIRED: Dict[str, Dict[str, int]] = {
    "A": {"supervisor": 2, "special": 2, "ordinary": 29, "all": 33, "contract": 31},
    "B": {"supervisor": 2, "special": 2, "ordinary": 29, "all": 33, "contract": 31},
    "C": {"supervisor": 2, "special": 2, "ordinary": 21, "all": 25, "contract": 23},
}

# Monday=0 .. Sunday=6
WEEKDAY_CH = ["一", "二", "三", "四", "五", "六", "日"]


def _weekday_strip(year: int, month: int, days: int) -> List[str]:
    return [WEEKDAY_CH[date(year, month, d).weekday()] for d in range(1, days + 1)]


def _get_holidays_for_month(year: int, month: int) -> Dict[int, str]:
    """Return a map of day -> holiday name for the given month."""
    holidays = list_public_holidays(year=year)
    result: Dict[int, str] = {}
    for h in holidays:
        try:
            y, m, d = h["date"].split("-")
            if int(y) == year and int(m) == month:
                result[int(d)] = h["name"]
        except Exception:
            continue
    return result


_REMARK_MAP_CACHE: Optional[Dict[str, str]] = None
_REMARK_MAP_CACHE_TS: float = 0.0
_REMARK_MAP_TTL: float = 5.0


def _build_remark_map() -> Dict[str, str]:
    """Build remark map from defaults + user settings (5-second TTL)."""
    global _REMARK_MAP_CACHE, _REMARK_MAP_CACHE_TS
    now = time.time()
    if _REMARK_MAP_CACHE is not None and (now - _REMARK_MAP_CACHE_TS) < _REMARK_MAP_TTL:
        return _REMARK_MAP_CACHE
    defaults: Dict[str, str] = {
        "AL": "AL",
        "SH": "FH",
        "FH": "FH",
        "H": "FH",
        "NSH": "NSH",
        "NP": "NP",
        "RD": "RD",
        "RP": "RD",
        "R": "RD",
        "SL": "SL",
        "NC": "NC",
        "ND": "ND",
        "EC": "EC",
        # 假期检查 (holiday-check) 可写入这些假期码,之前不在映射里会被
        # _cell_code fallback 成 "X"(缺勤),导致假期在 shortfall 显示为 X。
        "PH": "PH",   # 公眾假期
        "PL": "PL",   # 其他有薪假
        "S": "S",     # Secondment 借調/其他更份
        "F": "FS",
        "FS": "FS",
        "FW": "FS",
        "W": "WTT",
        "WTT": "WTT",
    }
    try:
        from ...models import get_settings
        settings = get_settings()
        user = json.loads(settings.get("remark_map_json") or "{}")
        _REMARK_MAP_CACHE = {**defaults, **user}
    except Exception:
        _REMARK_MAP_CACHE = defaults
    _REMARK_MAP_CACHE_TS = now
    return _REMARK_MAP_CACHE


def _cell_code(entry: Dict[str, Any]) -> str:
    """Map an OCR entry to a shortfall code: RD / 8 / X / partial / FS / FH / SL / WTT / NP / RP."""
    remark_map = _build_remark_map()
    remark = (entry.get("remark") or "").strip()
    if remark in remark_map:
        return remark_map[remark]
    # fallback: Kimi sometimes puts remark codes into clock_in/clock_out
    ci = (entry.get("clock_in") or "").strip()
    co = (entry.get("clock_out") or "").strip()
    for val in (ci, co):
        if val in remark_map:
            return remark_map[val]
    if entry.get("is_rest_day"):
        return "RD"
    if ci and co:
        return "8"
    if ci or co:
        return "?"  # partial — only in or out
    return "X"


def _resolve_month(pages: List[Dict[str, Any]]) -> str:
    for p in pages:
        m = (p.get("month") or "").strip()
        if m:
            return m
    return "2026-03"
