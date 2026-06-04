# Engine 目录

> **完整文档见 [ENGINE.md](./ENGINE.md)** — 涵盖引擎体系、API、测试流程、验证工具、常见问题等全部内容。

## 一句话说明

Engine 负责将数据库中的排班/打卡数据渲染为：
- **HTML JSON**（前端展示缺勤表）
- **Excel 文件**（下载打印正式报表）

## 快速测试

```bash
cd D:/claude/claude_hk/backend

# 运行全部 5 个组合（生成 + Excel verify）
python -m app.engine.test.run_all

# 只跑第 1 个组合
python -m app.engine.test.run_all --combo 1

# pytest 集成测试
pytest app/engine/test/test_engines.py -v
```

## 核心入口

```python
from app.engine import execute_html, execute_excel

# HTML 引擎 → JSON
data = execute_html("shortfall", shift="A", project_id=1, category_id=1, month="2026-02")

# Excel 引擎 → xlsx
result = execute_excel(engine_id=3, month="2026-02")
```

## 目录结构

```
engine/
    __init__.py              # execute_html(), execute_excel()
    registry.py              # HTML_ENGINES + EXCEL_ENGINES
    ENGINE.md                # ← 完整使用指南
    common/                  # 共享模块（数据源、模板解析、Excel 工具）
    html/                    # HTML 渲染引擎
    excel/                   # Excel 输出引擎
    test/                    # 测试框架 + 验证工具包装器
```
