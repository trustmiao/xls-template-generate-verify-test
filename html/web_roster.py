"""HTML engine wrapper for web_roster.

Loads template → runs Excel engine → returns HTML string.
"""
from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, Optional

from openpyxl import load_workbook

from ..common.data_source import resolve_template
from ..excel.web_roster import run as excel_run, workbook_to_html


def run(
    shift: str,
    project_id: Optional[int] = None,
    category_id: Optional[int] = None,
    month: Optional[str] = None,
) -> Dict[str, Any]:
    """Load template, run web_roster engine, return HTML.

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

    wb = load_workbook(str(excel_path), data_only=False, keep_links=False)

    context = {
        "project_id": project_id,
        "category_id": category_id,
        "month": effective_month,
    }

    excel_run(wb, context)
    html = workbook_to_html(wb, effective_month)

    return {
        "html": html,
        "month": effective_month,
        "template_path": str(rel_path) if rel_path else "",
    }
