"""Test configuration for engine test suite."""
from __future__ import annotations

from pathlib import Path

# ---------------------------------------------------------------------------
# Output directory
# ---------------------------------------------------------------------------
OUTPUT_DIR = Path(__file__).parent / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)

# ---------------------------------------------------------------------------
# Test combinations: 5 project/category/month combos with recommended engines
# ---------------------------------------------------------------------------
COMBINATIONS = [
    {
        "label": "大元邨-保安-2026-02",
        "project_id": 1,
        "category_id": 1,
        "month": "2026-02",
        "engine_id": 3,
        "shifts": ["A", "B", "C"],
        "sheets": ["早", "中", "夜"],
    },
    {
        "label": "大元邨-保安-2026-03",
        "project_id": 1,
        "category_id": 1,
        "month": "2026-03",
        "engine_id": 3,
        "shifts": ["A", "B", "C"],
        "sheets": ["早", "中", "夜"],
    },
    {
        "label": "东汇邨-保安-2026-02",
        "project_id": 2,
        "category_id": 3,
        "month": "2026-02",
        "engine_id": 45,
        "shifts": ["A", "B", "C"],
        "sheets": ["早", "中", "夜"],
    },
    {
        "label": "东汇邨-保安-2026-03",
        "project_id": 2,
        "category_id": 3,
        "month": "2026-03",
        "engine_id": 45,
        "shifts": ["A", "B", "C"],
        "sheets": ["早", "中", "夜"],
    },
    {
        "label": "东汇邨-保洁-2026-02",
        "project_id": 2,
        "category_id": 2,
        "month": "2026-02",
        "engine_id": 70,
        "shifts": ["A"],
        "sheets": ["Roster-FEB2026"],
    },
    {
        "label": "东汇邨-保洁-2026-03",
        "project_id": 2,
        "category_id": 2,
        "month": "2026-03",
        "engine_id": 70,
        "shifts": ["A"],
        "sheets": ["Roster-FEB2026"],
    },
]


def get_shift_label(shift: str) -> str:
    """Return Chinese label for shift code."""
    return {"A": "早", "B": "中", "C": "夜"}.get(shift, shift)
