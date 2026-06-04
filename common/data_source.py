"""Data source helpers for engines: load pages, templates, build context."""
from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional

from ...config import DOC_DIR
from ...models import (
    find_output_template_by_type,
    get_data_context,
    get_latest_output_template_by_type,
    get_output_engine,
    get_output_template,
    get_output_template_by_type,
    get_output_template_for_month,
    list_pages_with_entries,
)


def load_pages(
    shift: str,
    project_id: Optional[int] = None,
    category_id: Optional[int] = None,
    month: Optional[str] = None,
) -> List[Dict[str, Any]]:
    """Load pages with entries for a shift.

    Thin wrapper around models.list_pages_with_entries that applies
    corrected values (employee_no_corrected, name_corrected) when present.
    """
    return list_pages_with_entries(
        shift=shift,
        project_id=project_id,
        category_id=category_id,
        month=month,
    )


def resolve_template(
    project_id: Optional[int] = None,
    category_id: Optional[int] = None,
    month: Optional[str] = None,
    template_type: str = "roster_shortfall",
) -> tuple[Optional[Path], Optional[str]]:
    """Look up roster template from DB by effective date range.

    3-tier fallback:
      1. Explicit date-range match (non-null start/end)
      2. Exact month match (legacy)
      3. Latest unrestricted template
    Returns (absolute_path, relative_path) or (None, None).
    """
    if project_id and category_id and month:
        tpl = find_output_template_by_type(project_id, category_id, month, template_type)
        if tpl:
            path = DOC_DIR / tpl["file_path"]
            if path.exists():
                return path, tpl["file_path"]

        tpl = get_output_template_by_type(project_id, category_id, month, template_type)
        if tpl:
            path = DOC_DIR / tpl["file_path"]
            if path.exists():
                return path, tpl["file_path"]

        tpl = get_latest_output_template_by_type(project_id, category_id, template_type)
        if tpl:
            path = DOC_DIR / tpl["file_path"]
            if path.exists():
                return path, tpl["file_path"]

    return None, None


def build_context(project_id: int, category_id: int, month: str) -> Dict[str, Any]:
    """Build the standard data context for output engines.

    Thin wrapper around models.get_data_context.
    """
    return get_data_context(project_id, category_id, month)


def get_engine_config(engine_id: int) -> Optional[Dict[str, Any]]:
    """Load engine config from DB."""
    return get_output_engine(engine_id)


def resolve_template_for_engine(
    engine: Dict[str, Any],
    month: str,
) -> tuple[Optional[Path], Optional[str]]:
    """Resolve template for a given engine config.

    Priority:
      1. Explicit output_template_id on the engine
      2. Fallback by month + builtin_key as template_type
    """
    from ...config import DOC_DIR

    project_id = engine["project_id"]
    category_id = engine["category_id"]
    output_template_id = engine.get("output_template_id")

    tpl_row = None
    if output_template_id:
        tpl_row = get_output_template(output_template_id)
    if not tpl_row:
        fallback_type = engine.get("builtin_key") if engine.get("engine_type") == "builtin" else None
        tpl_row = get_output_template_for_month(
            project_id, category_id, month, template_type=fallback_type
        )

    if tpl_row:
        path = DOC_DIR / tpl_row["file_path"]
        if path.exists():
            return path, str(path.relative_to(DOC_DIR))

    return None, None
