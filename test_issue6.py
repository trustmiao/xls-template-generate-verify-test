#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Issue #6 test: Adjust personnel counts for Feb and Apr across all templates.

For each template:
  - Feb (28 days): segment1→1 person, segment2→2 people, segment3→3 people (if exists)
  - Apr (30 days): segment1→1 person, segment2→2 people, segment3→3 people (if exists)

Steps per template/month:
  1. _adjust_day_columns() + _update_dates() to set the month
  2. detect roster regions
  3. delete rows from each region until target count is reached
"""

import re
import calendar
from pathlib import Path
from openpyxl import load_workbook

import sys
sys.path.insert(0, r"D:\claude\claude_hk\backend")
from app.engine.excel.cleaning_shortfall_v2 import (
    _adjust_day_columns,
    _update_dates,
)
from excel_row_ops import (
    adjust_all_formulas,
    safe_delete_rows,
    fix_merged_cells,
    adjust_print_area,
    adjust_conditional_formatting,
)

TEMPLATES = [
    ("东汇-保安", "templates/东汇-保安轮休表template.xlsx", [1, 2]),
    ("东汇-保洁", "templates/东汇-保洁轮休表_template.xlsx", [1, 2, 3]),
    ("TY-大元保安", "templates/TY-2026.03- SG_SEC-Deploy Roster  Shortfall - template.xlsx", [1, 2]),
]

OUT_DIR = Path("test_outputs_issue6_v2")
OUT_DIR.mkdir(exist_ok=True)

# ── Region detection ──
rank_pattern = re.compile(r'^[A-Z]+\d+$')
title_pattern = re.compile(r'^\d+\.\s+')


def detect_roster_regions(ws):
    regions = []
    for row in range(1, ws.max_row + 1):
        b_val = ws.cell(row=row, column=2).value
        if b_val and isinstance(b_val, str) and title_pattern.match(b_val):
            title_row = row
            data_start = None
            for r in range(title_row + 1, ws.max_row + 1):
                b = ws.cell(row=r, column=2).value
                if b and isinstance(b, str) and rank_pattern.match(b.strip()):
                    data_start = r
                    break
            if data_start is None:
                continue
            data_end = data_start
            for r in range(data_start + 1, ws.max_row + 1):
                b = ws.cell(row=r, column=2).value
                is_data = b and isinstance(b, str) and rank_pattern.match(b.strip())
                if is_data:
                    data_end = r
                else:
                    break
            regions.append({
                'index': len(regions) + 1,
                'title_row': title_row,
                'data_start': data_start,
                'data_end': data_end,
                'title': b_val,
            })
    return regions


# ── Core deletion ──
def delete_second_row_of_region(ws, region):
    """Delete the 2nd data row of a region and fix formulas + merged cells + print area + conditional formatting."""
    row_to_delete = region['data_start'] + 1
    if row_to_delete > region['data_end']:
        return False

    adjust_all_formulas(ws, row_to_delete)
    safe_delete_rows(ws, row_to_delete, 1)
    fix_merged_cells(ws, row_to_delete)
    adjust_print_area(ws, row_to_delete)
    adjust_conditional_formatting(ws, row_to_delete)
    return True


def process_template(name, path, targets, month, out_dir):
    """Generate output for one template + month with specified segment targets."""
    year, mon = map(int, month.split('-'))
    days_in_month = calendar.monthrange(year, mon)[1]

    wb = load_workbook(path, data_only=False)

    # Step 1: Adjust day columns and dates on every sheet
    for ws in wb.worksheets:
        if ws.title == "Data":
            continue
        _adjust_day_columns(ws, days_in_month)
        _update_dates(ws, month)

    # Step 2: Delete rows from each region to reach target counts
    for ws in wb.worksheets:
        if ws.title == "Data":
            continue
        regions = detect_roster_regions(ws)
        print(f"\n  Sheet '{ws.title}': {len(regions)} region(s)")
        for region in regions:
            idx = region['index']
            if idx > len(targets):
                print(f"    Region {idx}: skipped (no target defined)")
                continue
            target = targets[idx - 1]
            current = region['data_end'] - region['data_start'] + 1
            to_delete = current - target
            if to_delete <= 0:
                print(f"    Region {idx}: already has {current} person(s), no deletion needed")
                continue

            print(f"    Region {idx}: {current} → {target} (deleting {to_delete} row(s))")
            for _ in range(to_delete):
                # Re-detect because row numbers shift after each deletion
                regions_now = detect_roster_regions(ws)
                r = next((x for x in regions_now if x['index'] == idx), None)
                if r is None:
                    break
                ok = delete_second_row_of_region(ws, r)
                if not ok:
                    break

    out_name = f"{month}_{name}_adjusted.xlsx"
    out_path = out_dir / out_name
    wb.save(str(out_path))
    print(f"  💾 Saved: {out_path}")
    return out_path


def inspect_file(path, label):
    """Print final segment structure."""
    wb = load_workbook(path, data_only=False)
    print(f"\n{'='*60}")
    print(f"INSPECT: {label}")
    print(f"{'='*60}")
    for ws in wb.worksheets:
        if ws.title == "Data":
            continue
        regions = detect_roster_regions(ws)
        print(f"\n  Sheet '{ws.title}':")
        for r in regions:
            count = r['data_end'] - r['data_start'] + 1
            names = []
            for row in range(r['data_start'], r['data_end'] + 1):
                rank = ws.cell(row, 2).value or ""
                name = ws.cell(row, 3).value or ""
                names.append(f"{rank}/{name}")
            print(f"    Region {r['index']}: {count} person(s) — {names}")


def main():
    for name, path, targets in TEMPLATES:
        for month in ["2026-02", "2026-04"]:
            print(f"\n{'='*60}")
            print(f"Processing: {name} | {month} | targets={targets}")
            print(f"{'='*60}")
            out = process_template(name, path, targets, month, OUT_DIR)
            inspect_file(out, f"{name} {month}")

    print(f"\n{'='*60}")
    print(f"All outputs saved to: {OUT_DIR}")
    print(f"{'='*60}")


if __name__ == '__main__':
    main()
