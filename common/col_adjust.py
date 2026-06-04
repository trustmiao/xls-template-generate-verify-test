"""Excel column adjustment helpers for date-column expansion/shrinkage.

Used by roster_shortfall (security) and cleaning_shortfall (cleaning) engines
to match template day columns to the target month length.
"""
from __future__ import annotations

import re
from copy import copy
from typing import Optional

# Regex for column references in formulas (e.g., A1, $A$1, AA10, $AA$10)
_COL_REF_RE = re.compile(r'(\$?)([A-Z]{1,3})(\$?)(\d+)')

# Regex for horizontal ranges (e.g., I16:AM16, $I$16:$AM$16)
_H_RANGE_RE = re.compile(r'(\$?[A-Z]{1,3}\$?\d+):(\$?[A-Z]{1,3}\$?\d+)')


def col_letter_to_num(letter: str) -> int:
    """Convert Excel column letter(s) to 1-based column number."""
    result = 0
    for c in letter.upper():
        result = result * 26 + (ord(c) - ord("A") + 1)
    return result


def col_num_to_letter(n: int) -> str:
    """Convert 1-based column number to Excel column letter(s)."""
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


def shift_formula_cols(formula: str, col_idx: int, amount: int, *, is_delete: bool = False) -> str:
    """Shift column references in an Excel formula.

    Args:
        formula: The formula string (must start with '=').
        col_idx: The column index where the operation starts (1-based).
        amount: Number of columns inserted (positive) or deleted (negative).
        is_delete: True if this is a deletion operation.
    """
    if is_delete:
        delete_count = abs(amount)
        last_kept = col_idx - 1

        # Step 1: For horizontal ranges, shrink end refs that land in deleted cols
        def _shrink_hrange(m: re.Match) -> str:
            start = m.group(1)
            end = m.group(2)
            # Parse end ref
            em = _COL_REF_RE.match(end)
            if not em:
                return m.group(0)
            end_col = col_letter_to_num(em.group(2))
            if end_col >= col_idx and end_col < col_idx + delete_count:
                new_end_letter = col_num_to_letter(last_kept)
                new_end = f"{em.group(1)}{new_end_letter}{em.group(3)}{em.group(4)}"
                return f"{start}:{new_end}"
            return m.group(0)

        formula = _H_RANGE_RE.sub(_shrink_hrange, formula)

        # Step 2: Shift remaining refs that are past the deleted block
        def repl(m: re.Match) -> str:
            dollar1 = m.group(1)
            col_letter = m.group(2)
            dollar2 = m.group(3)
            row_num = m.group(4)
            col_num = col_letter_to_num(col_letter)
            if col_num >= col_idx + delete_count:
                new_col = col_num - delete_count
                new_letter = col_num_to_letter(new_col)
                return f"{dollar1}{new_letter}{dollar2}{row_num}"
            return m.group(0)

        return _COL_REF_RE.sub(repl, formula)
    else:
        # Insert: shift cols >= col_idx
        def repl(m: re.Match) -> str:
            dollar1 = m.group(1)
            col_letter = m.group(2)
            dollar2 = m.group(3)
            row_num = m.group(4)
            col_num = col_letter_to_num(col_letter)
            if col_num >= col_idx:
                new_col = col_num + amount
                new_letter = col_num_to_letter(new_col)
                return f"{dollar1}{new_letter}{dollar2}{row_num}"
            return m.group(0)

        return _COL_REF_RE.sub(repl, formula)


def update_all_formulas_for_col_change(ws, col_idx: int, amount: int, *, is_delete: bool = False) -> None:
    """Update all formulas in the worksheet after column insert/delete."""
    for row in range(1, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row, col)
            v = cell.value
            if isinstance(v, str) and v.startswith("="):
                new_v = shift_formula_cols(v, col_idx, amount, is_delete=is_delete)
                if new_v != v:
                    cell.value = new_v

    # Update print area
    pa = ws.print_area
    if pa:
        new_pa = shift_formula_cols(pa, col_idx, amount, is_delete=is_delete)
        if new_pa != pa:
            ws.print_area = new_pa


def delete_cols_with_formulas(ws, delete_from: int, delete_count: int) -> None:
    """Delete columns and update all formulas referencing them."""
    update_all_formulas_for_col_change(ws, delete_from, -delete_count, is_delete=True)
    ws.delete_cols(delete_from, delete_count)


def insert_cols_with_formulas(
    ws,
    insert_at: int,
    amount: int,
    *,
    date_row: Optional[int] = None,
    weekday_row: Optional[int] = None,
) -> None:
    """Insert columns and update all formulas, then copy style/formula from prev column."""
    update_all_formulas_for_col_change(ws, insert_at, amount, is_delete=False)
    ws.insert_cols(insert_at, amount)

    # Copy style and formulas from the column before insertion point
    source_col = insert_at - 1
    for row in range(1, ws.max_row + 1):
        src = ws.cell(row, source_col)
        for offset in range(amount):
            target_col = insert_at + offset
            dst = ws.cell(row, target_col)

            if src.has_style:
                dst.font = copy(src.font)
                dst.border = copy(src.border)
                dst.fill = copy(src.fill)
                dst.number_format = copy(src.number_format)
                dst.protection = copy(src.protection)
                dst.alignment = copy(src.alignment)

            # Date row: copy with shifted column reference
            if row == date_row and src.value:
                if isinstance(src.value, str) and src.value.startswith("="):
                    dst.value = shift_formula_cols(src.value, insert_at, offset + 1, is_delete=False)
                else:
                    dst.value = src.value

            # Weekday row: copy with shifted column reference
            if row == weekday_row and src.value:
                if isinstance(src.value, str) and src.value.startswith("="):
                    dst.value = shift_formula_cols(src.value, insert_at, offset + 1, is_delete=False)
                else:
                    dst.value = src.value
