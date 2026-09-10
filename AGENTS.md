# QAtools Agent Guide

本文件是新会话的第一入口，作用域覆盖整个仓库。进入子目录后，如有更近的
`AGENTS.md`，同时遵循其专项约束。

## 开始前

1. 先读根目录 `README.md` 和 `docs/README.md`。
2. 修改某项工具前，读 `tools/<tool>/README.md` 和对应测试。
3. PhraseLoom 修改还需读 `phraseloom/AGENTS.md`。

当前行为以代码和测试为准，其次是当前文档。`docs/archive/` 只提供历史背景，
不能作为当前实现依据。

## 仓库入口

- [仓库地图](docs/repository-map.md)：调用链、模块职责、工具与测试对应关系、扩展步骤、大文件评估。
- [CLI 手册](docs/cli-usage.md)：已注册命令与调用；仅 GUI 功能不能假设存在 CLI。
- [待办](docs/backlog.md)：尚未完成的工作；不等于已实现能力。
- `tools/qt_pages.py` 仅做注册与兼容导出；统一页面实现放在工具包内的 `qt_page.py`。
- `docs/qa-workflow-guide/` 是操作指南；截图生成和页面验证见同目录 README，业务规则仍以工具 README 为准。

## 修改约定

- 业务规则放在工具包中；GUI 只负责收集输入、调用业务层和展示结果。
- 新增 GUI 工具时保持模块导入和页面构造轻量；重型可选依赖和文件读取放到用户触发操作后。
- 新增正式命令时同步更新 `qatools/cli.py`、`docs/cli-usage.md` 和测试。
- 保持既有 CLI、兼容脚本、工作表名称与输出格式，除非任务明确要求迁移。
- 使用 `openpyxl` 时显式关闭工作簿，避免 Windows 文件锁。
- 不提交 `build/`、`dist/`、`outputs/`、`testfiles/` 或生成的 Excel 文件。
- 改行为时同时更新对应工具 README；不要把当前规则写进历史设计稿。

## 验证

先运行相关测试，再运行全量检查：

```powershell
python -m unittest discover -s tests -v
python -m compileall -q qatools phraseloom tools tests
git diff --check
python scripts/check_docs.py
```

如果只改 PhraseLoom，可先运行：

```powershell
python -m unittest discover -s tests/phraseloom_tests -v
```

## 文档职责

- `README.md`：产品概览、最短上手路径和能力地图。
- `docs/README.md`：按任务路由用户和 agent，不重复业务规则。
- `docs/cli-usage.md`：统一 CLI 的唯一完整调用手册。
- `tools/<tool>/README.md`：单项工具的输入、输出与业务规则。
- `phraseloom/README.md`：Strings 工作簿与用户流程契约。
- `AGENTS.md`：agent 的入口、开发约束和验证方式；详细地图在 `docs/repository-map.md`。
- `docs/backlog.md`：唯一待办清单；完成后更新当前文档，不在旧方案里继续加需求。
- `docs/archive/`：已完成方案和迁移记录，不维护为当前事实。

一个事实只保留一个权威位置。其他文档使用链接引用，不复制完整参数或规则。
