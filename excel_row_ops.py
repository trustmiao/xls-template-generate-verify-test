#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Common row-deletion utilities for Excel templates.

Handles:
  - safe row deletion (replaces openpyxl's delete_rows)
  - formula adjustment
  - merged-cell range adjustment
  - print-area adjustment
  - conditional-formatting adjustment
"""

import re
import copy

# ── Formula Adjustment ──
cell_ref = re.compile(r'(?<![A-Z$])(\$?[A-Z]{1,3})(\$?)(\d+)')


def adjust_formula(formula, deleted_row):
    """
    Adjust cell references after deleting a row.

    - Range refs (A1:A5): both bounds use >= rule (shrink range)
    - Single refs (A1):    use > rule only (A1→#REF! if row==deleted)
    """
    range_positions = set()
    range_pattern = re.compile(
        r'(?<![A-Z$])(\$?[A-Z]{1,3})(\$?)(\d+)(:(\$?[A-Z]{1,3})(\$?)(\d+))'
    )
    for m in range_pattern.finditer(formula):
        range_positions.add((m.start(), m.end()))

    result = []
    last_end = 0

    for match in cell_ref.finditer(formula):
        in_range = any(
            start <= match.start() < end for start, end in range_positions
        )

        result.append(formula[last_end:match.start()])
        col_part = match.group(1)
        abs_row = match.group(2) == '$'
        row_num = int(match.group(3))

        if not abs_row:
            if in_range:
                if row_num >= deleted_row:
                    row_num -= 1
            else:
                if row_num > deleted_row:
                    row_num -= 1
                elif row_num == deleted_row:
                    result.append("#REF!")
                    last_end = match.end()
                    continue

        result.append(f"{col_part}{'$' if abs_row else ''}{row_num}")
        last_end = match.end()

    result.append(formula[last_end:])
    return ''.join(result)


def adjust_all_formulas(ws, deleted_row):
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        for cell in row:
            val = cell.value
            if val and isinstance(val, str) and val.startswith('='):
                cell.value = adjust_formula(val, deleted_row)


# ── Merged Cells ──
def adjust_merged_range(range_str, deleted_row):
    ref_pattern = re.compile(r'(\$?)([A-Z]{1,3})(\$?)(\d+)')

    def parse_ref(ref):
        m = ref_pattern.match(ref)
        if m:
            col_abs, col, row_abs, row = m.groups()
            row = int(row)
            if row > deleted_row:
                row -= 1
            return f"{col_abs}{col}{row_abs}{row}"
        return ref

    if ':' in range_str:
        start, end = range_str.split(':')
        return f"{parse_ref(start)}:{parse_ref(end)}"
    return parse_ref(range_str)


def fix_merged_cells(ws, deleted_row):
    """
    Adjust merged cell ranges after a row deletion WITHOUT unmerge+merge.

    openpyxl's unmerge_cells + merge_cells has a side effect of copying
    the top-left cell's style to ALL merged cells, which corrupts styles.
    We directly mutate MergedCellRange objects in place.
    """
    to_remove = []
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row > deleted_row:
            # Only adjust row numbers (shift would also move columns,
            # causing Invalid shift value when min_col==1)
            mr.min_row -= 1
            mr.max_row -= 1
        elif mr.max_row < deleted_row:
            pass
        else:
            if mr.min_row == mr.max_row == deleted_row:
                to_remove.append(mr)
            else:
                mr.max_row -= 1
                if mr.min_row > mr.max_row:
                    to_remove.append(mr)

    for mr in to_remove:
        ws.merged_cells.ranges.remove(mr)


# ── Print Area ──
def adjust_print_area(ws, deleted_row):
    if not ws.print_area or not ws._print_area:
        return
    for cr in ws._print_area.ranges:
        if cr.min_row > deleted_row:
            cr.min_row -= 1
        if cr.max_row > deleted_row:
            cr.max_row -= 1


# ── Conditional Formatting ──
def adjust_conditional_formatting(ws, deleted_row):
    """
    Adjust conditional-formatting ranges after a row deletion.

    Rules:
      - Ranges entirely below the deleted row → shift up by 1.
      - Ranges entirely above → unchanged.
      - Ranges that span the deleted row → shrink max_row by 1.
      - Single-row ranges on the deleted row → removed.

    Rebuilds the worksheet's ConditionalFormattingList from scratch to
    avoid internal key/hash inconsistencies when mutating existing objects.
    """
    from collections import OrderedDict
    from openpyxl.formatting.formatting import ConditionalFormatting
    from openpyxl.worksheet.cell_range import MultiCellRange

    new_entries = []          # list of (ConditionalFormatting, list_of_rules)

    for cf in ws.conditional_formatting:
        rules = list(cf.rules)
        adjusted_ranges = []

        for cr in cf.cells:
            if cr.min_row > deleted_row:
                # Entirely below — shift up by 1
                new_cr = copy.copy(cr)
                new_cr.min_row -= 1
                new_cr.max_row -= 1
                adjusted_ranges.append(str(new_cr))
            elif cr.max_row < deleted_row:
                # Entirely above — unchanged
                adjusted_ranges.append(str(cr))
            else:
                # deleted_row is inside or at the boundary
                if cr.min_row == cr.max_row == deleted_row:
                    # Single-row range completely deleted — drop it
                    continue
                else:
                    new_cr = copy.copy(cr)
                    new_cr.max_row -= 1
                    if new_cr.min_row > new_cr.max_row:
                        continue
                    adjusted_ranges.append(str(new_cr))

        if adjusted_ranges:
            new_cf = ConditionalFormatting()
            new_cf.cells = MultiCellRange(' '.join(adjusted_ranges))
            new_entries.append((new_cf, rules))

    # Rebuild _cf_rules from scratch
    ws.conditional_formatting._cf_rules = OrderedDict()
    for new_cf, rules in new_entries:
        ws.conditional_formatting._cf_rules[new_cf] = rules


# ── Safe Row Deletion (replaces ws.delete_rows) ──
def safe_delete_rows(ws, idx, amount=1):
    """
    Delete rows, correctly mimicking Excel behavior.
    Replaces openpyxl's ws.delete_rows() which has an iter_rows bug that
    creates phantom empty cells which then overwrite valid data during shift.
    """
    max_row = max((row for row, col in ws._cells.keys()), default=0)

    for r in range(idx, max_row + 1):
        source_r = r + amount
        if source_r > max_row:
            for key in list(ws._cells.keys()):
                if key[0] == r:
                    del ws._cells[key]
            continue

        target_cols = {col for row, col in ws._cells.keys() if row == r}
        source_cols = {col for row, col in ws._cells.keys() if row == source_r}

        for col in target_cols - source_cols:
            del ws._cells[(r, col)]

        for col in source_cols:
            if (source_r, col) in ws._cells:
                cell = ws._cells[(source_r, col)]
                ws._cells[(r, col)] = cell
                cell.row = r
                del ws._cells[(source_r, col)]

    # Update row dimensions (height + style + other attrs)
    for r in range(idx, max_row + 1):
        source_r = r + amount
        dst_rd = ws.row_dimensions[r]
        if source_r in ws.row_dimensions:
            src_rd = ws.row_dimensions[source_r]
            dst_rd.height = src_rd.height
            if src_rd._style is not None:
                dst_rd._style = copy.copy(src_rd._style)
            else:
                dst_rd._style = None
            dst_rd.hidden = src_rd.hidden
            dst_rd.outline_level = src_rd.outline_level
            dst_rd.collapsed = src_rd.collapsed
        else:
            dst_rd.height = None
            dst_rd._style = None
            dst_rd.hidden = False
            dst_rd.outline_level = 0
            dst_rd.collapsed = False

    ws._current_row = max(
        row for row, col in ws._cells.keys()
    ) if ws._cells else 0


# ── High-level: delete one row and fix everything ──
def delete_row_and_fixup(ws, row_to_delete):
    """
    Delete a single row and fix all dependent structures.

    Applies, in order:
      1. Formula adjustment
      2. Safe physical row deletion
      3. Merged-cell range adjustment
      4. Print-area adjustment
      5. Conditional-formatting adjustment
    """
    adjust_all_formulas(ws, row_to_delete)
    safe_delete_rows(ws, row_to_delete, 1)
    fix_merged_cells(ws, row_to_delete)
    adjust_print_area(ws, row_to_delete)
    adjust_conditional_formatting(ws, row_to_delete)
