"""Engine registry — maps builtin keys to engine callables."""
from __future__ import annotations

from typing import Any, Callable, Dict

from . import excel, html

# HTML engines (return JSON dict)
HTML_ENGINES: Dict[str, Callable] = {
    "shortfall": html.shortfall_run,
}

# Excel engines (mutate workbook in-place)
EXCEL_ENGINES: Dict[str, Callable] = {
    "roster_shortfall": excel.roster_shortfall_run,
    "generic_roster": excel.generic_roster_run,
    "generic_roster_cleaner": excel.generic_roster_cleaner_run,
    "deployment": excel.deployment_run,
    "cleaning_roster": excel.cleaning_roster_run,
    "cleaning_shortfall": excel.cleaning_shortfall_run,
}

# Legacy alias for backward compat
BUILTIN_ENGINES = EXCEL_ENGINES
