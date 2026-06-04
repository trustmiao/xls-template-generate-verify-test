"""Excel verify wrapper — calls excel-verify skill scripts."""
from __future__ import annotations

import calendar
import json
import subprocess
import sys
from pathlib import Path
from typing import Any, Dict, Optional

SKILL_DIR = Path(r"D:\claude\claude_hk\backend\tools\skills\excel-verify\scripts")
PYTHON = sys.executable


def run_detect_regions(xlsx_path: Path, sheet: str) -> Dict[str, Any]:
    """Run detect_regions.py --json and return parsed output."""
    script = SKILL_DIR / "detect_regions.py"
    if not script.exists():
        return {"ok": False, "error": f"Script not found: {script}"}

    cmd = [str(PYTHON), str(script), "--xlsx", str(xlsx_path), "--sheet", sheet, "--json"]
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60, encoding="utf-8")
        stdout = result.stdout.strip()
        # detect_regions.py prints JSON first, then human-readable summary
        for line in stdout.splitlines():
            line_stripped = line.strip()
            if line_stripped.startswith("{"):
                try:
                    data = json.loads(line_stripped)
                    data["ok"] = result.returncode == 0
                    data["raw_stdout"] = stdout
                    return data
                except json.JSONDecodeError:
                    continue
        # Fallback
        if "{" in stdout:
            json_start = stdout.find("{")
            try:
                data = json.loads(stdout[json_start:])
                data["ok"] = result.returncode == 0
                data["raw_stdout"] = stdout
                return data
            except json.JSONDecodeError:
                pass
        return {"ok": result.returncode == 0, "raw_stdout": stdout, "raw_stderr": result.stderr}
    except Exception as e:
        return {"ok": False, "error": str(e)}


def run_check_roster_filled(
    xlsx_path: Path, template_path: Path, sheet: str,
    project_id: int, category_id: int, month: str, shift: str,
) -> Dict[str, Any]:
    """Run check_roster_filled.py and return result."""
    script = SKILL_DIR / "check_roster_filled.py"
    if not script.exists():
        return {"ok": False, "error": f"Script not found: {script}"}

    cmd = [
        str(PYTHON), str(script),
        "--xlsx", str(xlsx_path), "--template", str(template_path),
        "--sheet", sheet,
        "--project", str(project_id), "--category", str(category_id),
        "--month", month, "--shift", shift,
    ]
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120, encoding="utf-8")
        return {
            "ok": result.returncode == 0,
            "returncode": result.returncode,
            "stdout": result.stdout,
            "stderr": result.stderr,
        }
    except Exception as e:
        return {"ok": False, "error": str(e)}


def run_compare_template(
    xlsx_path: Path, template_path: Path, sheet: str,
    out_csv: Optional[Path] = None,
) -> Dict[str, Any]:
    """Run compare_template.py and return result.

    Compares generated xlsx against its template, skipping dynamic zones
    (roster + date), checking formula/value/style/merge in static skeleton.
    """
    script = SKILL_DIR / "compare_template.py"
    if not script.exists():
        return {"ok": False, "error": f"Script not found: {script}"}

    cmd = [
        str(PYTHON), str(script),
        "--template", str(template_path),
        "--generated", str(xlsx_path),
        "--sheet", sheet,
        "--ignore-font-name",
        "--ignore-nfmt",
    ]
    if out_csv:
        cmd += ["--out", str(out_csv)]
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120, encoding="utf-8")
        return {
            "ok": result.returncode == 0,
            "returncode": result.returncode,
            "stdout": result.stdout,
            "stderr": result.stderr,
        }
    except Exception as e:
        return {"ok": False, "error": str(e)}


def run_check_date_zone(xlsx_path: Path, sheet: str, month: str) -> Dict[str, Any]:
    """Check date zone column count matches target month days."""
    dr = run_detect_regions(xlsx_path, sheet)
    if not dr.get("ok"):
        return {"ok": False, "error": dr.get("error", "detect_regions failed")}

    dz = dr.get("date_zone", {})
    date_row = dz.get("date_row")
    col_start = dz.get("col_start")
    col_end = dz.get("col_end")

    if not date_row or col_start is None or col_end is None:
        return {"ok": True, "note": "no date_zone detected, skip"}

    try:
        expected = calendar.monthrange(int(month[:4]), int(month[5:7]))[1]
    except Exception as e:
        return {"ok": False, "error": f"invalid month {month}: {e}"}

    actual = col_end - col_start + 1
    if actual != expected:
        return {
            "ok": False,
            "error": f"date zone has {actual} day columns, expected {expected} for {month}",
            "expected": expected,
            "actual": actual,
        }

    return {"ok": True, "days": actual}


def run_check_cleaning_summary(xlsx_path: Path, sheet: str, month: str) -> Dict[str, Any]:
    """Run cleaning-specific checks (issue #4) on a generated xlsx."""
    try:
        from .cleaning import run_check_cleaning_summary as _check
        return _check(xlsx_path, sheet, month)
    except Exception as e:
        return {"ok": False, "error": str(e)}


def find_template_for_combo(project_id: int, category_id: int, month: str) -> Optional[Path]:
    """Look up the roster_shortfall template path for a given combo."""
    try:
        from ...common.data_source import resolve_template
        path, _ = resolve_template(project_id, category_id, month, "roster_shortfall")
        return path if path and path.exists() else None
    except Exception:
        return None
