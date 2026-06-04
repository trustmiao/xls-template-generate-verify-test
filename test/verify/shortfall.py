"""Shortfall roster verify wrapper.

NOTE: These scripts require backend + frontend dev servers running.
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path
from typing import Any, Dict, Optional

SKILL_DIR = Path(r"D:\claude\claude_hk\backend\tools\skills\shortfall-roster-verify\scripts")
PYTHON = sys.executable


def _run_script(script_name: str, args: list[str], timeout: int = 180) -> Dict[str, Any]:
    script = SKILL_DIR / script_name
    if not script.exists():
        return {"ok": False, "error": f"Script not found: {script}"}
    cmd = [str(PYTHON), str(script)] + args
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout, encoding="utf-8")
        return {
            "ok": result.returncode == 0,
            "returncode": result.returncode,
            "stdout": result.stdout,
            "stderr": result.stderr,
        }
    except Exception as e:
        return {"ok": False, "error": str(e)}


def run_check_roster_match(project_id: int, category_id: int, month: str, shift: Optional[str] = None) -> Dict[str, Any]:
    args = ["--project", str(project_id), "--category", str(category_id), "--month", month]
    if shift:
        args += ["--shift", shift]
    return _run_script("check_roster_match.py", args, timeout=120)


def run_check_zero_shortage(project_id: int, category_id: int, month: str) -> Dict[str, Any]:
    args = ["--project", str(project_id), "--category", str(category_id), "--month", month]
    return _run_script("check_zero_shortage.py", args, timeout=120)


def run_strict_grid(project_id: int, category_id: int, month: str, shift: str, out_path: Path) -> Dict[str, Any]:
    args = [
        "--project", str(project_id), "--category", str(category_id),
        "--month", month, "--shift", shift, "--out", str(out_path),
    ]
    return _run_script("strict_grid.py", args, timeout=300)


def run_check_cross_posting_display(project_id: int, category_id: int, month: str, shift: Optional[str] = None) -> Dict[str, Any]:
    args = ["--project", str(project_id), "--category", str(category_id), "--month", month]
    if shift:
        args += ["--shift", shift]
    return _run_script("check_cross_posting_display.py", args, timeout=120)


def run_check_holiday_shortfall(project_id: int, category_id: int, month: str, base_url: str = "http://127.0.0.1:8094") -> Dict[str, Any]:
    args = ["--project", str(project_id), "--category", str(category_id), "--month", month, "--base", base_url]
    return _run_script("check_holiday_shortfall.py", args, timeout=120)
