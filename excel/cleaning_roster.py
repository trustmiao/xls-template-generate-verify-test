"""Built-in engine: cleaning roster fill.

Wraps the existing cleaning_fill + template_filler logic so it can be invoked
via the generic output engine system.
"""
from __future__ import annotations

from typing import Any, Dict

from openpyxl import Workbook


def run(wb: Workbook, context: Dict[str, Any]) -> None:
    """Fill a cleaning roster template workbook.

    Uses the existing services.cleaning_fill and services.template_filler
    logic, but adapted to work inside the generic engine framework.
    """
    project_id = context["project_id"]
    category_id = context["category_id"]
    month = context["month"]

    # Import here to avoid circular dependencies at module load time
    from ...services.cleaning_fill import build_cleaning_fill
    from ...services.template_filler import fill_cleaning_template_from_workbook
    from ...models import list_role_requirements, list_employee_roles

    plan = build_cleaning_fill(project_id, category_id, month)
    reqs = list_role_requirements(project_id=project_id, category_id=category_id)
    employees = list_employee_roles(project_id=project_id, category_id=category_id, month=month)

    fill_cleaning_template_from_workbook(wb, plan, reqs, employees)
