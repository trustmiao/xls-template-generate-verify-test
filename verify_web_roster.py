#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""逐格验证：HTML 输出 vs Excel 模板 (Issue #5)

兼容新 API（run() 返回 sheets 字典）。
本脚本已更新为使用 verify_web.py 的核心逻辑。
"""
from __future__ import annotations

import sys

sys.path.insert(0, r"D:\claude\claude_hk\backend")

from app.engine.verify_web import (
    parse_html_table,
    verify_sheet,
    format_excel_value,
    parse_css_style,
    get_excel_cell_info,
    parse_html_border,
)
from app.engine.html.web_roster import run, _get_print_area_bounds
from openpyxl.utils import get_column_letter


def main():
    print("=" * 80)
    print("Web Roster 逐格验证 (Issue #5)")
    print("=" * 80)

    templates = [
        ("TY-大元保安", 1, 1),
        ("东汇-保安", 2, 3),
        ("东汇-保洁", 2, 2),
    ]

    all_ok = True
    for name, pid, cid in templates:
        for month in ["2026-02", "2026-04"]:
            print(f"\n【验证】{name} — {month}")

            result = run("A", project_id=pid, category_id=cid, month=month)
            sheets_html = result["sheets"]
            wb = result.get("_wb")
            wb_values = result.get("_wb_values")

            if wb is None:
                print(f"  ⚠ 无法获取处理后的 workbook")
                continue

            total_errors = 0
            for sheet_name, html in sheets_html.items():
                print(f"\n  Sheet: {sheet_name}")

                if sheet_name not in wb.sheetnames:
                    print(f"    ⚠ Sheet '{sheet_name}' 不在 workbook 中")
                    continue

                ws = wb[sheet_name]
                value_ws = wb_values[sheet_name] if wb_values and sheet_name in wb_values.sheetnames else None

                html_data = parse_html_table(html)
                print(f"    HTML: {html_data['max_row']} 行 × {html_data['max_col']} 列")

                bounds = _get_print_area_bounds(ws)
                if bounds:
                    print(f"    Print area: {get_column_letter(bounds[0])}{bounds[1]}:{get_column_letter(bounds[2])}{bounds[3]}")

                errors = verify_sheet(html_data, ws, value_ws=value_ws)
                if errors:
                    total_errors += len(errors)
                    print(f"    ❌ {len(errors)} 个差异")
                    for e in errors[:10]:
                        pos = e.get("pos", "")
                        print(f"      ({pos[0]},{pos[1]}) [{e['type']}] {e['msg'][:120]}")
                    if len(errors) > 10:
                        print(f"      ... 还有 {len(errors) - 10} 个")
                else:
                    print(f"    ✅ 无差异")

            if total_errors == 0:
                print(f"\n  ✅ {name} {month} — 全部通过")
            else:
                print(f"\n  ❌ {name} {month} — 共 {total_errors} 个差异")
                all_ok = False

    print("\n" + "=" * 80)
    if all_ok:
        print("✅ 所有模板全部通过逐格验证!")
    else:
        print("❌ 存在差异，需要进一步修复")
    print("=" * 80)


if __name__ == "__main__":
    main()
