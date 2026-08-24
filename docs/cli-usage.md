# QAtools CLI 使用指南

这份文档是 QAtools 命令行的统一入口，面向人工调用、agent、自动化脚本和
批处理任务。各工具的业务规则仍以对应工具 README 为准。

## 安装与入口

推荐在仓库根目录安装：

```bash
python -m pip install -e .
```

安装后统一使用：

```bash
qatools <命令> [参数]
```

不安装也可以使用完全等价的模块入口：

```bash
python -m qatools <命令> [参数]
```

查看命令列表和原生参数：

```bash
qatools --help
qatools list
qatools help qa
qatools tag-check --help
```

## 命令目录

| 命令 | 用途 |
|---|---|
| `qatools gui` | 打开统一 Toolshub GUI |
| `qatools qa` | 一次执行多项质量检查并生成统一报告 |
| `qatools phraseloom` | 导出 Strings 并在翻译后回填 |
| `qatools term-check` | 术语 mark 与历史 TB 检查 |
| `qatools tag-check` | 常规 Tag、Placeholder、换行标记与 memoQ Marker 检查 |
| `qatools line-break-check` | 真实换行数量检查 |
| `qatools consistency-check` | 同源译文一致性检查 |
| `qatools chinese-check` | Target 中文字符与中文标点检查 |
| `qatools french-nbsp` | 法语 NBSP 恢复 |
| `qatools batch` | 按行拆分 Excel batch，并在处理后复原 |
| `qatools xbench` | Xbench QA Report 转换 |

可用别名：

- `qatools workflow` → `qatools qa`
- `qatools strings` → `qatools phraseloom`
- `qatools source-consistency` → `qatools consistency-check`
- `qatools target-chinese` → `qatools chinese-check`
- `qatools excel-batch` → `qatools batch`
- `qatools xbench-transform` → `qatools xbench`

## 自动化调用约定

- 始终显式传入输入文件、工作表、列和输出路径。
- 列参数使用 Excel 列字母，例如 `A`、`B`、`AA`。
- 不要依赖缺参后的终端交互。
- 除特别说明外，工具生成新 Excel，不覆盖输入文件。
- GUI 的工作表和列自动识别不属于 CLI 默认行为。
- 每个子命令的 `--help` 直接来自对应工具参数解析器，是参数名称的权威来源。

## 一键质量检查

默认运行术语、Tag、换行数量、同源译文一致性、Target 中文和 Target 文本规范检查：

```bash
qatools qa ./input.xlsx \
  -s Sheet1 \
  -c A \
  -t B \
  --start-row 2 \
  -o ./artifacts/workflow_check_input.xlsx
```

用 `--check` 重复选择部分检查：

```bash
qatools qa ./input.xlsx -c A -t B \
  --check tag \
  --check line-break \
  --check consistency
```

可选检查值：

```text
term
tag
line-break
consistency
chinese
text
```

Target 文本规范检查支持单独选择规则；不传 `--text-rule` 时默认运行全部三项：

```bash
qatools qa ./input.xlsx -c A -t B \
  --check text \
  --text-rule abnormal-punctuation \
  --text-rule consecutive-spaces \
  --text-rule mixed-width
```

术语和 Tag 的高级参数：

```bash
qatools qa ./input.xlsx -c A -t B \
  --term-mark-style '【】' \
  --term-mark-style '[]' \
  --history-tb ./history_tb.xlsx \
  --history-sheet Glossary \
  --tag-token-type angle \
  --tag-token-type brace
```

只使用历史 TB、不从 mark 提取新术语：

```bash
qatools qa ./input.xlsx -c A -t B \
  --check term \
  --no-term-mark \
  --history-tb ./history_tb.xlsx
```

## PhraseLoom

导出待翻译 Strings：

```bash
qatools phraseloom export ./source.xlsx \
  --source-col source \
  --target-col target \
  --context-col context \
  -o ./source_strings.xlsx
```

可选相似结构分组：

```bash
qatools phraseloom export ./source.xlsx --group-similar
```

默认会按换行拆分多行 Source；如需整格导出：

```bash
qatools phraseloom export ./source.xlsx --no-split-lines
```

回填翻译：

```bash
qatools phraseloom restore ./source_strings.xlsx \
  -o ./source_translated.xlsx
```

Strings 工作簿内嵌原工作簿、行映射和 Tag 规则，回填时无需再次传入原始
Excel。完整规则见 [PhraseLoom README](../phraseloom/README.md)。

## 单项质量检查

术语检查：

```bash
qatools term-check ./input.xlsx \
  -s Sheet1 -c A -t B --start-row 2 \
  --mark-style '【】' \
  --history-tb ./history_tb.xlsx \
  -o ./artifacts/term_pair_check_input.xlsx
```

Tag / Placeholder 检查：

```bash
qatools tag-check ./input.xlsx \
  -s Sheet1 -c A -t B --start-row 2 \
  --token-type angle \
  --token-type square_color \
  --token-type brace \
  --token-type newline \
  -o ./artifacts/tag_check_input.xlsx
```

换行数量检查：

```bash
qatools line-break-check ./input.xlsx \
  -s Sheet1 -c A -t B --start-row 2 \
  -o ./artifacts/line_break_check_input.xlsx
```

同源译文一致性：

```bash
qatools consistency-check ./input.xlsx \
  -s Sheet1 -c A -t B --start-row 2 \
  -o ./artifacts/source_consistency_check_input.xlsx
```

Target 中文检查：

```bash
qatools chinese-check ./input.xlsx \
  -s Sheet1 -c A -t B --start-row 2 \
  -o ./artifacts/target_chinese_check_input.xlsx
```

## 文本修复与转换

法语 NBSP 恢复：

```bash
qatools french-nbsp ./input.xlsx \
  -s Sheet1 -t B -r C --start-row 2 \
  -o ./artifacts/input_french_nbsp_restored.xlsx
```

Xbench 报告转换：

```bash
qatools xbench ./Xbench_QA_Report.xlsx \
  -s "Xbench QA" \
  -o ./artifacts/xbench_flat.xlsx
```

Excel batch 拆分：

```bash
qatools batch split ./input.xlsx \
  --sheet Sheet1 \
  --batch-size 1000 \
  --header-rows 1 \
  --output-dir ./artifacts/input_batches
```

完成分批作业后复原：

```bash
qatools batch restore ./artifacts/input_batches \
  --output ./artifacts/input_restored.xlsx
```

拆分目录中的 `batch_manifest.json` 和 `_qatools_restore_source_*` 模板是复原所需
文件，应与 batch 一起保留。完整规则见
[Excel batch 工具](../tools/excel_batcher/README.md)。

## 兼容入口

统一 CLI 是新文档和新自动化的首选入口。以下旧入口继续可用：

| 统一命令 | 兼容入口 |
|---|---|
| `qatools gui` | `python toolshub_gui.py`、`toolshub` |
| `qatools phraseloom` | `phraseloom`、`python -m phraseloom.cli` |
| `qatools term-check` | `python tools/term_pair_checker/extract_terms_from_excel.py` |
| `qatools tag-check` | `python tools/tag_placeholder_checker/check_tags_and_placeholders.py` |
| `qatools line-break-check` | `python tools/line_break_checker/check_line_breaks.py` |
| `qatools consistency-check` | `python tools/source_consistency_checker/check_source_consistency.py` |
| `qatools chinese-check` | `python tools/chinese_target_checker/check_chinese_target.py` |
| `qatools french-nbsp` | `python tools/french_nbsp_restorer/restore_french_nbsp.py` |
| `qatools batch` | `python tools/excel_batcher/excel_batcher.py` |
| `qatools xbench` | `python tools/xbench_report_transformer/transform_xbench_report.py` |

`qatools qa` 是统一新增的一键检查 CLI；此前该流程只有 GUI。

## 扩展新命令

新增 CLI 工具时：

1. 在独立模块中保留业务逻辑和原生参数解析器。
2. 在 `qatools/cli.py` 的 `COMMANDS` 中登记命令、说明和模块。
3. 在本文件加入至少一个 `qatools <命令>` 示例。
4. 添加转发测试和工具自身测试。
5. 运行 `qatools <命令> --help`、完整 unittest 和 wheel 构建。

测试会校验每个正式命令都已出现在本指南中，从而减少注册表与文档漂移。
