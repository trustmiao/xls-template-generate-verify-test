#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Delete the 2nd data row from the m-th roster region in an Excel template.

Usage:
    python delete_one_row.py <input.xlsx> <output.xlsx> <region_index>

Example:
    python delete_one_row.py template.xlsx output.xlsx 2
    # Deletes the 2nd data row from region 2
"""

import re
import sys
from pathlib import Path

try:
    import openpyxl
except ImportError:
    print("Error: openpyxl is required. Install with: pip install openpyxl")
    sys.exit(1)

from excel_row_ops import (
    adjust_all_formulas,
    safe_delete_rows,
    fix_merged_cells,
    adjust_print_area,
    adjust_conditional_formatting,
)


# ──────────────────────────────────────────────────────────────────────────────
# Region Detection
# ──────────────────────────────────────────────────────────────────────────────

def detect_roster_regions(ws):
    """
    Detect all roster (personnel list) regions in a worksheet.

    A region is defined as:
      - Title row: B-column starts with "N. " (e.g. "1. 保安主管出勤時數")
      - Data rows: B-column contains a rank code like M01, A1, D01, etc.
      - Ends at the first non-data row below the data block.

    Returns a list of dicts:
        [{
            'index': 1-based region index,
            'title_row': row number of the title,
            'data_start': first data row,
            'data_end':   last data row,
            'title':      title text,
        }, ...]
    """
    regions = []
    rank_pattern = re.compile(r'^[A-Z]+\d+$')
    title_pattern = re.compile(r'^\d+\.\s+')

    for row in range(1, ws.max_row + 1):
        b_val = ws.cell(row=row, column=2).value
        if b_val and isinstance(b_val, str) and title_pattern.match(b_val):
            title_row = row

            # Find first data row (rank code in column B)
            data_start = None
            for r in range(title_row + 1, ws.max_row + 1):
                b = ws.cell(row=r, column=2).value
                if b and isinstance(b, str) and rank_pattern.match(b.strip()):
                    data_start = r
                    break

            if data_start is None:
                continue

            # Find data end: last consecutive row with a rank code in column B
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


# ──────────────────────────────────────────────────────────────────────────────
# Main
# ──────────────────────────────────────────────────────────────────────────────

def process_workbook(input_path, output_path, region_index):
    wb = openpyxl.load_workbook(input_path)

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        regions = detect_roster_regions(ws)

        print(f"\n📋 Sheet '{sheet_name}' — detected {len(regions)} roster region(s):")
        for r in regions:
            print(f"   Region {r['index']}: rows {r['data_start']}-{r['data_end']} "
                  f"({r['data_end'] - r['data_start'] + 1} rows) — {r['title']}")

        region = next((r for r in regions if r['index'] == region_index), None)
        if region is None:
            print(f"   ⚠️  Region {region_index} not found, skipping.")
            continue

        row_to_delete = region['data_start'] + 1
        if row_to_delete > region['data_end']:
            print(f"   ⚠️  Region {region_index} has only 1 row, nothing to delete.")
            continue

        print(f"   🗑️  Deleting row {row_to_delete} (2nd data row of region {region_index})")

        # 1. Adjust formulas
        adjust_all_formulas(ws, row_to_delete)

        # 2. Physically remove the row (safe replacement for ws.delete_rows)
        safe_delete_rows(ws, row_to_delete, 1)

        # 3. Fix merged cells (direct mutation — no unmerge/merge side effects)
        fix_merged_cells(ws, row_to_delete)

        # 4. Adjust print area
        adjust_print_area(ws, row_to_delete)

        # 5. Adjust conditional formatting
        adjust_conditional_formatting(ws, row_to_delete)

        print(f"   ✅ Done. Original row {row_to_delete} deleted.")

    wb.save(output_path)
    print(f"\n💾 Saved to: {output_path}")
    return wb


def main():
    if len(sys.argv) >= 4:
        input_path = sys.argv[1]
        output_path = sys.argv[2]
        region_index = int(sys.argv[3])
    else:
        # Default test
        input_path = "templates/东汇-保洁轮休表_template.xlsx"
        output_path = "test_outputs_delete_rows/东汇-保洁_区域2删第2行.xlsx"
        region_index = 2
        print("No arguments provided; running default test case.")
        print(f"  Input:  {input_path}")
        print(f"  Output: {output_path}")
        print(f"  Region: {region_index}")

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    process_workbook(input_path, output_path, region_index)


if __name__ == '__main__':
    main()
