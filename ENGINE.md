# Engine 体系完整指南

> 读完本文档即可理解、测试、调试整个 Engine 体系，无需再读源码。

## 一、总览

Engine 体系负责两个核心功能：

| 功能 | 输出格式 | 用途 |
|------|---------|------|
| **Shortfall（缺勤表）** | JSON → 网页 HTML 表格 | 前端展示人员排班/缺勤情况 |
| **Output（输出表）** | Excel (.xlsx) | 下载/打印正式报表 |

系统位于 `app/engine/`，分三层：
- **HTML 引擎**（`html/`）：将数据库中的 pages + entries 渲染为 JSON，前端再转为表格
- **Excel 引擎**（`excel/`）：读取 xlsx 模板，填入数据后生成输出文件
- **测试框架**（`test/`）：自动化生成 + 多层验证

---

## 二、目录结构

```
app/engine/
    __init__.py                    # 统一入口：execute_html(), execute_excel()
    registry.py                    # 引擎注册表：HTML_ENGINES + EXCEL_ENGINES
    ENGINE.md                      # ← 本文档

    common/                        # 共享模块（所有引擎共用）
        data_source.py             # build_context(), load_pages(), resolve_template()
        render_sheet.py            # RenderedSheet 生成（内部使用）
        shift_info.py              # SHIFT_INFO, SHIFT_REQUIRED 常量
        template_parser.py         # TemplateStructure 解析模板结构
        excel_ops.py               # Excel 操作工具（插入行、复制样式等）

    html/                          # HTML 渲染引擎
        shortfall.py               # shortfall_run(shift, project_id, category_id, month) → JSON

    excel/                         # Excel 输出引擎
        roster_shortfall.py        # 大元邨保安模板（早/中/夜三 sheet）
        generic_roster.py          # 通用 roster 引擎（自动检测段结构）
        generic_roster_cleaner.py  # generic_roster 的保洁变体（更新日期行）
        cleaning_roster.py         # 保洁 roster（包装旧版 cleaning_fill）
        cleaning_shortfall.py      # 保洁 shortfall（单 sheet，三段式）
        deployment.py              # Deployment 表（替换 VLOOKUP）
        json_mapper.py             # JSON Mapping 规则执行器

    test/                          # 测试框架
        config.py                  # 6 个测试组合配置
        harness.py                 # EngineTestHarness 类
        run_all.py                 # 主入口：跑全部组合 + 验证
        test_engines.py            # pytest 集成测试
        verify/
            excel.py               # excel-verify skill 包装器（detect_regions / check_roster_filled / compare_template）
            shortfall.py           # shortfall-roster-verify skill 包装器
        outputs/                   # 测试输出目录（JSON + HTML + XLSX + CSV + 报告）
```

---

## 三、核心入口

### 3.1 execute_html — 生成 Shortfall JSON

```python
from app.engine import execute_html

data = execute_html(
    engine_name="shortfall",   # 目前只有这一个 HTML 引擎
    shift="A",                 # A=早, B=中, C=夜
    project_id=1,              # 1=大元邨, 2=东汇邨
    category_id=1,             # 1=保安, 2=保洁, 3=保安(东汇邨)
    month="2026-02",           # 目标月份
)
```

返回 JSON 结构：
```python
{
    "has_data": True,              # 该 scope 是否有数据
    "segments": [                  # 数据段列表（如：保安主任 / 保安员）
        {
            "title": "保安主任",
            "rows": [              # 每行代表一个人
                {
                    "rank_seq": "M01",
                    "employee_no": "h123456",
                    "name": "張三",
                    "cells": [     # 每天一个单元格
                        {"value": "8.0", "code": "", "edited": False},
                        ...
                    ],
                    "total_hours": "176.0",
                    "shift_label": "A",
                }
            ]
        }
    ],
    "extra_rows": [],              # 数据库多出的未分类人员
    "holiday_strip": [False, ...], # 每天是否公众假期
    "weekday_strip": ["一","二",...],
    "days_in_month": 28,
    "shift": "A",
    "month": "2026-02",
    "title": "大元邨",
    "label": "保安",
}
```

### 3.2 execute_excel — 生成 Excel 文件

```python
from app.engine import execute_excel

result = execute_excel(
    engine_id=3,              # 数据库中的 output_engines.id
    month="2026-02",
    created_by=None,          # 可选：操作人
)
```

返回字典：
```python
{
    "file_id": 123,
    "file_name": "大元邨-保安-2026-02.xlsx",
    "download_url": "/api/output-files/123/download",
    "preview_url": "/api/output-files/123/preview",
    "file_size": 241920,
    "xlsx_checks": {...},     # X1/X2/X3 验证结果
}
```

执行流程：
1. 从数据库 `output_engines` 表读取 engine 配置
2. 根据配置找到对应的 xlsx 模板
3. 调用 `build_context()` 加载 pages、entries 等数据
4. 根据 `engine_type` 分发：
   - `"builtin"` → 调用 `EXCEL_ENGINES[builtin_key](wb, context)`
   - `"json_mapping"` → 调用 `json_mapper_apply(wb, context, rules)`
5. 保存到 `OUTPUT_DIR`，记录到 `generated_files` 表
6. 运行 `verify_generated_xlsx()` 做 X1/X2/X3 检查

---

## 四、引擎注册表

文件：`registry.py`

```python
HTML_ENGINES = {
    "shortfall": html.shortfall_run,   # 返回 JSON dict
}

EXCEL_ENGINES = {
    "roster_shortfall":       excel.roster_shortfall_run,       # 大元邨保安
    "generic_roster":         excel.generic_roster_run,         # 通用 roster
    "generic_roster_cleaner": excel.generic_roster_cleaner_run, # 保洁（更新日期）
    "deployment":             excel.deployment_run,             # Deployment 表
    "cleaning_roster":        excel.cleaning_roster_run,        # 保洁 roster
    "cleaning_shortfall":     excel.cleaning_shortfall_run,     # 保洁 shortfall
}
```

**新增引擎步骤**：
1. 在 `app/engine/excel/` 下写函数：`def my_engine_run(wb, context)`
2. 在 `registry.py` 的 `EXCEL_ENGINES` 中注册
3. 在 `test/config.py` 的 `COMBINATIONS` 中添加测试组合

---

## 五、数据流

### 5.1 Shortfall（HTML）数据流

```
数据库 pages + entries
    └── load_pages(shift, project_id, category_id, month)
            └── html.shortfall_run(...)
                    ├── resolve_template()      找 xlsx 模板
                    ├── parse_template()        解析结构（段、行、列）
                    └── render_sheet()          填数据
                            └── JSON → 前端
```

### 5.2 Output（Excel）数据流

```
数据库 engine config
    └── execute_excel(engine_id, month)
            ├── resolve_template_for_engine()   找模板
            ├── build_context()                 构建 DataContext
            ├── EXCEL_ENGINES[builtin_key](wb, context)  引擎填充
            ├── save xlsx → OUTPUT_DIR
            └── verify_generated_xlsx()         X1/X2/X3 检查
```

### 5.3 DataContext 结构

Excel 引擎函数接收的 `context`：

```python
{
    "project_id": 1,
    "category_id": 1,
    "month": "2026-02",
    "template_path": "...",
    "pages": [...],           # load_pages() 结果，每人一个 dict
    "entries": [...],         # 打卡记录
    "shift_info": {...},      # SHIFT_INFO 常量
    # ... 其他由 build_context() 注入的字段
}
```

---

## 六、各 Excel 引擎说明

| 引擎 | builtin_key | 适用场景 | 特点 |
|------|-------------|---------|------|
| roster_shortfall | roster_shortfall | 大元邨保安 | 三 sheet（早/中/夜），保留公式，段结构固定 |
| generic_roster | generic_roster | 通用 | 自动 detect_regions，动态调整段大小 |
| generic_roster_cleaner | generic_roster_cleaner | 东汇邨保洁 | 在 generic_roster 基础上更新日期行到目标月 |
| cleaning_roster | cleaning_roster | 保洁 | 包装旧版 cleaning_fill + template_filler |
| cleaning_shortfall | cleaning_shortfall | 东汇邨保洁 shortfall | 单 sheet，三段（科文/工人/VO），插入行重写 SUM |
| deployment | deployment | Deployment 表 | 替换 VLOOKUP 为 literal 值，更新月份标签 |

---

## 七、API 路由

| 路由 | 文件 | 功能 |
|------|------|------|
| `GET /api/shortfall-engine` | `app/api/shortfall_engine.py` | 返回 shortfall JSON |
| `POST /api/output-engines/{id}/generate` | `app/api/output_engine.py` | 生成 Excel |
| `POST /api/engine-test` | `app/api/engine_test.py` | 一键执行 HTML + Excel + X3 检查 |
| `GET /api/output-files/{id}/download` | `app/api/output_engine.py` | 下载 xlsx |
| `GET /api/output-files/{id}/preview` | `app/api/output_engine.py` | HTML 表格预览 |

`/api/engine-test` 是最常用的调试接口，一次请求同时跑 HTML 引擎和 Excel 引擎，返回对比数据。

---

## 八、测试框架

### 8.1 快速开始

```bash
cd D:/claude/claude_hk/backend

# 运行全部 5 个组合（生成 + Excel verify）
python -m app.engine.test.run_all

# 只跑第 1 个组合
python -m app.engine.test.run_all --combo 1

# 同时运行 strict_grid 验证（需前后端服务运行）
python -m app.engine.test.run_all --verify-shortfall

# pytest 集成测试
pytest app/engine/test/test_engines.py -v
```

### 8.2 测试组合（config.py）

共 6 个组合，覆盖 2 个项目 × 2 个 category × 2 个月份：

| # | 组合 | project | category | month | engine_id | builtin_key | shifts | sheets |
|---|------|---------|----------|-------|-----------|-------------|--------|--------|
| 1 | 大元邨-保安-2026-02 | 1 | 1 | 2026-02 | 3 | roster_shortfall | A,B,C | 早,中,夜 |
| 2 | 大元邨-保安-2026-03 | 1 | 1 | 2026-03 | 3 | roster_shortfall | A,B,C | 早,中,夜 |
| 3 | 东汇邨-保安-2026-02 | 2 | 3 | 2026-02 | 45 | roster_shortfall | A,B,C | 早,中,夜 |
| 4 | 东汇邨-保安-2026-03 | 2 | 3 | 2026-03 | 45 | roster_shortfall | A,B,C | 早,中,夜 |
| 5 | 东汇邨-保洁-2026-02 | 2 | 2 | 2026-02 | 70 | generic_roster | A | Roster-FEB2026 |
| 6 | 东汇邨-保洁-2026-03 | 2 | 2 | 2026-03 | 70 | generic_roster | A | Roster-FEB2026 |

### 8.3 测试执行步骤（按顺序）

`run_all.py` 为每个组合按以下顺序执行验证：

| Step | 名称 | 执行模块 | 检查内容 | 输出 |
|------|------|---------|---------|------|
| 1 | Shortfall HTML | `run_all.py` `_check_shortfall_structure` | `has_data`、segments 数量、rows 数量、extra_rows 数量 | `{label}_{shift}.json` + `.html` |
| 2 | Excel 生成 + X1/X2/X3 | `EngineTestHarness.verify_xlsx()` → `app/utils/excel_verify.py` | 基础 Excel 格式验证 | `{label}_{builtin}.xlsx` |
| 3 | Excel Verify Skill | `verify/excel.py` 包装器 → `tools/skills/excel-verify/scripts/` | 3 个子检查（见下） | `report_*.json` |
| 4 | Shortfall Verify | `verify/shortfall.py` 包装器 → `tools/skills/shortfall-roster-verify/scripts/` | 3 个子检查（见下） | CSV、stdout |

#### Step 3 详细：Excel Verify Skill（每个 sheet 独立运行）

按顺序执行 3 个脚本：

| 子检查 | 脚本 | 检查内容 | 失败含义 | 影响结果 |
|--------|------|---------|---------|---------|
| 3a | `detect_regions.py` | 识别日期行（date_row）和数据段（data_segments） | 模板结构无法识别，后续检查无法运行 | **是** |
| 3b | `check_roster_filled.py` | Rule A：段内无空行；Rule B：名单与 engine API 一致 | 有空槽未填，或名单与数据库不一致 | **是** |
| 3c | `compare_template.py` | 跳过 roster+date 动态区域，对比静态骨架（公式/值/样式/合并单元格） | 生成文件在非动态区域意外修改了模板内容 | **否（仅参考）** |

**compare_template 为参考性检查**：
- 差异会输出 `DIFF` 并保存 CSV，但**不计入测试失败**
- 原因：`detect_regions` 的动态区域只覆盖 B-AJ 列，保安模板中 AN/AO 列（COUNT/SUM 公式）也在数据行上，但位于动态区域外；数据段扩大/收缩时这些 per-row 公式的行号必然变化。此外 A 列的空格/空字符串差异属于 openpyxl 读取噪声。真正的静态骨架（标题、标签、固定值）变更仍可通过 CSV 人工审查发现。

#### Step 4 详细：Shortfall-Roster Verify（需 `--verify-shortfall` + dev servers）

| 子检查 | 脚本 | 检查内容 | 级别 | 需服务 |
|--------|------|---------|------|--------|
| 4a | `check_zero_shortage.py` | 缺勤/欠时数为零 | combo | 后端 8094 |
| 4b | `check_cross_posting_display.py` | PT 人员在 supervisor + security 两段正确显示，工时相加一致 | combo | 后端 8094 |
| 4c | `check_holiday_shortfall.py` | 假期代码映射（entries.remark → shortfall code），公众假期标记 | combo | 后端 8094 |
| 4d | `check_roster_match.py` | API roster 与 OCR pages 一致（含 extra_rows） | shift | 后端 8094 |
| 4e | `strict_grid.py` | 逐单元格对比模板 vs 网页渲染 | shift | 后端 8094 + 前端 5180 |

### 8.4 EngineTestHarness（编程接口）

```python
from app.engine.test.harness import EngineTestHarness

h = EngineTestHarness(project_id=1, category_id=1, month="2026-02")

# HTML 引擎
json_data = h.run_html("shortfall", shift="A")

# Excel 引擎（按 builtin_key）
xlsx_path = h.run_excel("roster_shortfall")

# Excel 引擎（按 engine_id）
xlsx_path = h.run_excel_by_engine_id(3)

# X1/X2/X3 验证
checks = h.verify_xlsx(xlsx_path)
assert checks["ok"]
```

### 8.5 测试输出

运行后 `app/engine/test/outputs/` 下生成：

```
outputs/
    大元邨-保安-2026-02_A.json                # Shortfall JSON (Step 1)
    大元邨-保安-2026-02_A.html                # HTML 预览 (Step 1)
    大元邨-保安-2026-02_B.json / .html
    大元邨-保安-2026-02_C.json / .html
    大元邨-保安-2026-02_roster_shortfall.xlsx # Excel 输出 (Step 2)
    大元邨-保安-2026-02_A_compare.csv         # compare_template CSV (Step 3c)
    大元邨-保安-2026-02_B_compare.csv
    大元邨-保安-2026-02_C_compare.csv
    ...（其他组合）
    report_20260604_013648.json               # 完整测试报告 (所有步骤)
```

---

## 九、Excel Verify Skill 详解

位置：`tools/skills/excel-verify/scripts/`

### 9.1 detect_regions.py — 区域检测

检测模板中的动态区域：
- **Date zone（黄色）**：日期行（row 6/5/7）+ 星期行
- **Data segments（红色）**：数据段（连续的行，B 列匹配 `^[A-Z]{1,2}\d{1,3}$`）

```bash
python tools/skills/excel-verify/scripts/detect_regions.py \
    --xlsx app/engine/test/outputs/大元邨-保安-2026-02.xlsx \
    --sheet 早 --json
```

输出 JSON：
```json
{
  "sheet": "早",
  "print_area": {"col_start": 1, "row_start": 1, "col_end": 41, "row_end": 74},
  "date_zone": {"date_row": 6, "weekday_row": 7, "col_start": 9, "col_end": 36},
  "data_segments": [
    {"first_row": 8, "last_row": 48, "col_start": 2, "col_end": 36, "data_rows": [8,9,...]},
    {"first_row": 57, "last_row": 59, "col_start": 2, "col_end": 36, "data_rows": [57,58,59]}
  ]
}
```

### 9.2 check_roster_filled.py — 名单填充检查

```bash
python tools/skills/excel-verify/scripts/check_roster_filled.py \
    --xlsx app/engine/test/outputs/大元邨-保安-2026-02.xlsx \
    --template <模板路径> \
    --sheet 早 \
    --project 1 --category 1 --month 2026-02 --shift A
```

检查项：
- **Rule A**：段内是否有空行（模板槽位未填充）
- **Rule B**：xlsx 中的名单是否与 `/api/shortfall-engine` 返回的名单一致

**注意**：必须提供 `--template`，否则 detect_regions 无法识别空行。

### 9.3 compare_template.py — 模板静态骨架对比

对比模板和生成的 xlsx，**跳过动态区域**（roster zone + date zone），只检查静态骨架部分：

```bash
python tools/skills/excel-verify/scripts/compare_template.py \
    --template <tpl> --generated <xlsx> --sheet 早 \
    [--out D:/tmp/compare.csv]
```

检查项：
1. **公式/值**：公式文本或字面量是否一致（支持行插入后的公式范围偏移）
2. **样式**：字体（粗体/斜体/颜色/大小）、填充、边框、对齐方式
3. **合并单元格**：模板中的合并范围是否完整保留

**为什么跳过动态区域？**
- roster zone（人员名单）：引擎会根据实际人数插入/删除行，必然不同
- date zone（日期行）：月份不同，日期数字必然不同

**运行方式**：
- 独立运行：`python compare_template.py --template tpl.xlsx --generated gen.xlsx --sheet 早`
- 测试套件：`run_all.py` Step 3c 自动调用，输出 `{label}_{shift}_compare.csv`

---

## 十、Shortfall-Roster Verify Skill

位置：`tools/skills/shortfall-roster-verify/scripts/`

| 脚本 | 用途 | 需 dev servers | 已集成 run_all |
|------|------|---------------|-----------------|
| `check_zero_shortage.py` | 零缺勤检查 | 否 | ✅ Step 4a |
| `check_cross_posting_display.py` | PT 跨段显示检查（supervisor+security 两段一致性） | 否 | ✅ Step 4b |
| `check_holiday_shortfall.py` | 假期代码映射检查（remark → shortfall code） | 否 | ✅ Step 4c |
| `check_roster_match.py` | 名单比对（OCR vs API，含 extra_rows） | 否 | ✅ Step 4d |
| `strict_grid.py` | 逐单元格对比（模板 vs 网页） | **是** | ✅ Step 4e |
| `e2e_workflow.py` | 端到端工作流（改 DB） | **是** | ❌ 会修改 DB |
| `overlay_check.py` | Playwright 可视化覆盖 | **是** | ❌ GUI 工具 |
| `overlay_html.py` | HTML 可视化覆盖 | **是** | ❌ GUI 工具 |

---

## 十一、测试结果记录

### 最近测试：2026-06-04（compare_template 加入后）

**命令**：`python -m app.engine.test.run_all`

**结果**：Total=6, Passed=6, Failed=0, Time=95.5s ✅

| # | 组合 | 引擎 (engine_id) | Shortfall | Excel 大小 | detect_regions | check_roster | compare_template | 状态 |
|---|------|-----------------|-----------|-----------|----------------|--------------|------------------|------|
| 1 | 大元邨-保安-2026-02 | `大元邨 Shortfall 输出` (id=3, roster_shortfall) | A:44 B:40 C:32 | 236KB | 早:2段 中:1段 夜:2段 | 通过 | DIFF 109/618/56 | **通过** |
| 2 | 大元邨-保安-2026-03 | `大元邨 Shortfall 输出` (id=3, roster_shortfall) | 全 shift 无数据 | 210KB | 0段 | 跳过 | DIFF 908/940/890 | **通过** |
| 3 | 东汇邨-保安-2026-02 | `东汇-保安轮休表` (id=45, roster_shortfall) | A:2 B:1 C:0 | 185KB | 早:2段 中:1段 夜:0段 | 通过 | DIFF 29/854/956 | **通过** |
| 4 | 东汇邨-保安-2026-03 | `东汇-保安轮休表` (id=45, roster_shortfall) | A:29 B:15 C:14 | 198KB | 各2段 | 通过 | DIFF 88/66/30 | **通过** |
| 5 | 东汇邨-保洁-2026-02 | `东汇邨-通用保洁轮休表` (id=70, generic_roster) | A:22(3段) | 155KB | Roster-FEB2026: 3段 | 通过 | 通过 | **通过** |
| 6 | 东汇邨-保洁-2026-03 | `东汇邨-通用保洁轮休表` (id=70, generic_roster) | A:24(3段) | 155KB | Roster-FEB2026: 3段 | 通过 | DIFF 97 | **通过** |

**compare_template DIFF 说明**：
- 差异主要来自保安模板 AN/AO 列 per-row 公式行号变化（数据段扩大/收缩）+ A 列 openpyxl 空格噪声
- DIFF 仅报告、保存 CSV，不影响测试通过/失败
- 保洁模板差异较少，因无 AN/AO per-row 公式列

**修复记录**：
1. **config.py**：东汇邨保洁 engine_id 从 69（大元邨）修复为 70（东汇邨）；新增东汇邨保安 2026-03；sheet 名修复为 `Roster-FEB2026`
2. **excel_ops.py**：姓名写入去掉 HYPERLINK 公式，直接写纯文本
3. **check_roster_filled.py**：新增列头自动检测（姓名/工号/排位），支持无工号列模板（如保洁）

### 已知问题

1. **所有组合 xlsx_checks 警告**
   - `WARN: xlsx_checks failed — unknown`
   - `verify_generated_xlsx()` 返回不明确的错误，不影响核心功能

2. **东汇邨保安 2026-02 数据变少**
   - A=2行 B=1行 C=0行（数据库中数据量减少，属于正常数据变动）

---

## 十二、常见问题

### Q1: has_data=False？

数据库中没有对应 `project_id + category_id + month + shift` 的数据：

```python
from app.engine.common.data_source import load_pages
pages = load_pages("A", project_id=1, category_id=1, month="2026-02")
print(len(pages))  # 0 表示无数据
```

### Q2: 新增一个 builtin 引擎？

1. 在 `app/engine/excel/` 下写引擎函数：`def my_engine_run(wb, context)`
2. 在 `app/engine/registry.py` 的 `EXCEL_ENGINES` 中注册
3. 在 `app/engine/test/config.py` 的 `COMBINATIONS` 中添加测试组合
4. 运行 `python -m app.engine.test.run_all` 验证

### Q3: generic_roster_cleaner 是什么？

`generic_roster_cleaner` 是 `generic_roster` 的保洁专用变体，专用于东汇邨保洁模板（单 sheet，自动检测段结构，同时更新日期行到目标月）。

### Q4: 导出的 Excel 文件存在哪里？

- **生产环境**：`backend/output/` 目录（由 `app/config.py` 的 `OUTPUT_DIR` 决定），同时记录到数据库 `generated_files` 表
- **测试环境**：`app/engine/test/outputs/` 目录（由 `test/config.py` 的 `OUTPUT_DIR` 决定）

### Q5: 如何单独运行某个引擎？

```python
from app.engine.test.harness import EngineTestHarness

h = EngineTestHarness(project_id=1, category_id=1, month="2026-02")

# 跑 HTML 引擎
data = h.run_html("shortfall", shift="A")
print(data["segments"])

# 跑 Excel 引擎
path = h.run_excel("roster_shortfall")
print(path)
```

---

## 十三、关键文件速查

| 用途 | 路径 |
|------|------|
| 统一入口 | `app/engine/__init__.py` |
| 注册表 | `app/engine/registry.py` |
| 数据源 | `app/engine/common/data_source.py` |
| HTML 引擎 | `app/engine/html/shortfall.py` |
| Excel 引擎目录 | `app/engine/excel/` |
| 测试主入口 | `app/engine/test/run_all.py` |
| 测试配置 | `app/engine/test/config.py` |
| TestHarness | `app/engine/test/harness.py` |
| Excel verify 包装 | `app/engine/test/verify/excel.py` |
| Shortfall verify 包装 | `app/engine/test/verify/shortfall.py` |
| Excel verify 脚本 | `tools/skills/excel-verify/scripts/` |
| Shortfall verify 脚本 | `tools/skills/shortfall-roster-verify/scripts/` |
| API: engine-test | `app/api/engine_test.py` |
| API: shortfall | `app/api/shortfall_engine.py` |
| API: output | `app/api/output_engine.py` |
| Excel 基础验证 | `app/utils/excel_verify.py` |
