# PhraseLoom Agent Guide

本文件补充根目录 `AGENTS.md`，仅作用于 `phraseloom/`。

## 产品边界

PhraseLoom 是确定性的 Excel Strings 预处理与回填工具，只有两个核心动作：

1. `export`：清洗未翻译 Source，生成自包含的 Strings 工作簿。
2. `restore`：把译文写回工作簿内嵌的原始文件副本。

不要在此流程中加入 TM 预填、术语抽取、翻译推断或自动生成译文。

## 模块地图

- `strings_workflow.py`：导出与回填编排。
- `strings_package.py`：Strings 工作簿格式、元数据与样式。
- `workbook_io.py`：源工作簿读取。
- `cleaning.py`：去重与模板压缩。
- `tag_engine.py`、`tag_rules.py`：保护内容的提取、校验与恢复。
- `template_engine.py`：数字、颜色和序列模板。
- `string_cluster.py`：可选的相似句分组，仅影响分组和顺序。
- `cli.py`、`interactive.py`、`gui.py`：用户入口。
- `workbook_schema.py`：当前工作簿常量。

当前用户流程和工作簿契约见 `README.md`；行为细节以
`tests/phraseloom_tests/` 为准。

## 不变量

- 未明确要求迁移时，不改变 Strings 工作簿 schema。
- `tag_engine.py`、`template_engine.py`、`string_cluster.py` 不做 Excel I/O。
- 相似句分组不能合并普通字符串或生成 Target。
- 回填必须保留已有 Target、原工作表结构、可见性和格式。
- 始终关闭 `openpyxl` 工作簿。

修改工作流后至少运行：

```powershell
python -m unittest discover -s tests/phraseloom_tests -p test_strings_workflow_e2e.py -v
python -m unittest discover -s tests/phraseloom_tests -v
```
