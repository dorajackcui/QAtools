# QAtools CLI 使用指南

这份文档是 QAtools 命令行的统一入口，面向人工调用、AI agent 和自动化脚本。
参数名称以各命令的 `--help` 为准；业务判定规则以对应工具 README 为准。

## 安装与入口

推荐在仓库根目录安装：

```bash
python -m pip install -e .
```

安装后统一使用：

```bash
qatools <命令> [参数]
```

在仓库根目录、已安装依赖的 Python 环境中，也可以使用等价的模块入口：

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
| `qatools gui` | 打开统一 GUI |
| `qatools qa` | 一次执行多项质量检查并生成统一报告 |
| `qatools phraseloom` | 导出 Strings 并在翻译后回填 |
| `qatools term-check` | 术语 mark 与历史 TB 检查 |
| `qatools tag-check` | 常规 Tag、Placeholder、换行标记与 memoQ Marker 检查 |
| `qatools line-break-check` | 真实换行数量检查 |
| `qatools consistency-check` | 同 Source 不同 Target 检查 |
| `qatools chinese-check` | Target 中文字符与中文标点检查 |
| `qatools french-nbsp` | 法语 NBSP 恢复 |
| `qatools batch` | 按行拆分 Excel batch，并在处理后复原 |
| `qatools merge-sheets` | 合并目录内所有 Excel 文件的活动工作表 |
| `qatools xbench` | Xbench QA Report 转换 |

可用别名：

- `qatools workflow` → `qatools qa`
- `qatools strings` → `qatools phraseloom`
- `qatools source-consistency` → `qatools consistency-check`
- `qatools target-chinese` → `qatools chinese-check`
- `qatools excel-batch` → `qatools batch`
- `qatools merge-active-sheets` → `qatools merge-sheets`
- `qatools xbench-transform` → `qatools xbench`

## 自动化调用约定

- 始终显式传入输入文件、工作表、列和输出路径。
- QA、单项检查和 NBSP 的列参数使用 Excel 列字母，例如 `A`、`B`、`AA`；PhraseLoom 使用表头名或从 1 开始的列索引。
- 不要依赖缺参后的终端交互。
- 除特别说明外，工具生成新 Excel，不覆盖输入文件。
- GUI 的工作表和列自动识别不属于 CLI 默认行为。
- 每个子命令的 `--help` 直接来自对应工具参数解析器，是参数名称的权威来源。

本页多行示例使用 Bash 的 `\` 续行。PowerShell 请合并为单行，或使用反引号续行；含空格路径加引号。

```powershell
python -m qatools qa "D:\work\input file.xlsx" -s Sheet1 -c A -t B --start-row 2 -o "D:\work\qa.xlsx"
if ($LASTEXITCODE -ne 0) { throw "QAtools 执行失败" }
```

退出码：未知命令和参数解析错误通常为 `2`；统一分发捕获的文件不存在、键或值错误为 `1`；
正常完成为 `0`。其他异常可能由原生入口直接退出，不应依赖固定异常文本。
`0` 不表示 QA 零问题，也不表示批处理没有跳过文件；自动化还需检查输出摘要和报告。
当前 CLI 没有统一 JSON 输出协议。

## 功能可用范围

内容同步的两个方向、列操作、兼容性重存、同名文件替换和未翻译统计目前**仅有 GUI**，
没有已注册的统一 CLI 命令。实现计划见 [CLI-01](backlog.md#cli-01-新增工具的-cli)。
不要把 Python 业务函数名、GUI key 或待定命令名当作正式命令调用。

同 Target 不同 Source、数字、URL、Target 文本规范通过 `qa --check` 使用，未单独登记命令。
QA 的“应用修订”目前是 GUI 动作，也没有对应的 `qa` 子命令。

## 一键质量检查

默认运行八项常用检查；“同 Target 不同 Source”因合理复用较常见而默认关闭。
组合运行时，所有检查共用一次主工作簿读取和一次保存：

```bash
qatools qa ./input.xlsx \
  -s Sheet1 \
  -c A \
  -t B \
  --start-row 2 \
  -o ./artifacts/workflow_check_input.xlsx
```

检查名称、默认状态和规则入口：

| `--check` 值 | GUI 名称 | 默认 | 规则 |
|---|---|---:|---|
| `term` | 术语检查 | 开 | [术语检查](../tools/term_pair_checker/README.md) |
| `consistency` | 同 Source 不同 Target | 开 | [同 Source 不同 Target](../tools/source_consistency_checker/README.md) |
| `target-consistency` | 同 Target 不同 Source | 关 | [同 Target 不同 Source](../tools/target_consistency_checker/README.md) |
| `tag` | Tag / Placeholder | 开 | [Tag / Placeholder](../tools/tag_placeholder_checker/README.md) |
| `line-break` | 换行数量 | 开 | [换行数量](../tools/line_break_checker/README.md) |
| `number` | 数字一致性 | 开 | [数字与 URL](../tools/content_fidelity_checker/README.md) |
| `url` | URL 一致性 | 开 | [数字与 URL](../tools/content_fidelity_checker/README.md) |
| `chinese` | Target 中文 | 开 | [Target 中文](../tools/chinese_target_checker/README.md) |
| `text` | Target 文本规范 | 开 | [Target 文本规范](../tools/target_text_checker/README.md) |

一旦传入 `--check`，只运行显式选择的项目；该参数可以重复：

```bash
qatools qa ./input.xlsx -c A -t B \
  --check tag \
  --check line-break \
  --check consistency
```

例如同时运行双向一致性检查：

```bash
qatools qa ./input.xlsx -c A -t B \
  --check consistency \
  --check target-consistency
```

Target 文本规范检查支持单独选择规则；不传 `--text-rule` 时默认运行全部四项：

```bash
qatools qa ./input.xlsx -c A -t B \
  --check text \
  --text-rule abnormal-punctuation \
  --text-rule consecutive-spaces \
  --text-rule leading-trailing-spaces \
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

统一报告的 `问题处理`、`质量检查汇总` 和修订回填契约见
[workflow README](../tools/workflow/README.md)。

## PhraseLoom

`qatools phraseloom` 无参数时进入交互终端；自动化必须选择 `export` 或 `restore`。
`qatools phraseloom gui` 打开独立兼容 GUI；统一桌面入口仍为 `qatools gui`。

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

同 Source 不同 Target：

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

合并一个目录（含子目录）内所有 Excel 文件的活动工作表：

```bash
qatools merge-sheets ./excel-files \
  -o ./artifacts/merged.xlsx
```

默认只保留第一份表头。如需保留每个文件的第一行：

```bash
qatools merge-sheets ./excel-files --keep-all-headers
```

完整规则见[合并表格工具](../tools/excel_merger/README.md)。

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
| `qatools merge-sheets` | `python tools/excel_merger/merge_active_sheets.py` |
| `qatools xbench` | `python tools/xbench_report_transformer/transform_xbench_report.py` |

独立 Tk GUI 和根目录兼容脚本的定位见[仓库地图](repository-map.md#兼容边界)；它们仍受兼容性约束。

## 参数速查

以下表格补充前文示例；执行环境中的完整参数和可选值始终以 `qatools help <命令>` 为准。

| 入口 | 参数与说明 |
|---|---|
| `qa` | 必填 `input_file -c/--source-column -t/--target-column`；`-s/--sheet` 默认活动表，`--start-row` 默认 2，`-o/--output` 指定报告；`--check` 可重复 |
| `qa` 术语 | `--term-mark-style` 可重复，与 `--no-term-mark` 互斥；后者需要 `--history-tb`；历史 TB 范围用 `--history-sheet`、`--history-source-column`、`--history-target-column`、`--history-start-row` |
| `qa` Tag / 文本 | `--tag-token-type` 可重复；`--tag-angle-config` 为过滤 JSON；`--text-rule` 可重复。`numeric`、`abnormal-ellipsis` 分别为兼容别名 |
| `term-check` | 输入/范围/输出参数同检查器习惯；mark 参数名为 `--mark-style`，不是 QA 的 `--term-mark-style`；支持上述 `--history-*` 和 `--no-term-mark`，另有 `--exclusion-config` 候选排除 JSON |
| `tag-check` | 输入/范围/输出参数同检查器习惯；使用 `--token-type`、`--angle-config`，不带 QA 的 `tag-` 前缀；token 可选 `angle`、`square_color`、`brace`、`newline`、`memoq`（旧别名 `numeric`） |
| `line-break-check` / `consistency-check` / `chinese-check` | `input_file`、`-s/--sheet`、`-c/--source-column`、`-t/--target-column`、`--start-row`、`-o/--output` |
| `french-nbsp` | `input_file`、`-s/--sheet`、`-t/--target-column`、`--start-row`、`-o/--output`；`-r/--result-column` 可选，未指定则在输出副本的 Target 列修复 |
| `xbench` | `input_file`、`-s/--sheet`、`-o/--output`；输入是 Xbench 报告，不能套用 `-c/-t` |
| `batch split` | `input_file`、`--sheet`、`--batch-size`（默认 1000）、`--header-rows`（默认 1）、`--output-dir` |
| `batch restore` | 位置参数为 batch 目录或 manifest 路径；`--output` 指定复原文件 |
| `merge-sheets` | 位置参数目录，或兼容 `--folder-path`；`-o/--output`（兼容 `--output-path`）；`--keep-all-headers` 保留每份表头 |
| `phraseloom export` | `input`、`-o/--output`、`--source-col`、`--target-col`、`--context-col`、`--tag-config`（TOML）、`--group-similar`、`--split-lines/--no-split-lines` |
| `phraseloom restore` | `input` 为 Strings 工作簿；`-o/--output` 指定回填文件 |

部分兼容检查器缺少参数时会发起终端输入；脚本应显式传入参数，尤其是输入路径和列。
历史 TB 可以按表头自动识别，和 GUI 自定义别名是不同机制，见[术语规则](../tools/term_pair_checker/README.md)。

Tag 过滤示例：

```bash
qatools tag-check input.xlsx -c A -t B --token-type angle --angle-config custom_angle_tags.json
```

术语候选过滤示例：

```bash
qatools term-check input.xlsx -c A -t B --exclusion-config term_exclusions.json
```

## 扩展新命令

新增 CLI 工具时：

1. 在独立模块中保留业务逻辑和原生参数解析器。
2. 在 `qatools/cli.py` 的 `COMMANDS` 中登记命令、说明和模块。
3. 在本文件加入至少一个 `qatools <命令>` 示例。
4. 添加转发测试和工具自身测试。
5. 运行 `qatools <命令> --help`、完整 unittest 和 wheel 构建。

`python scripts/check_docs.py` 校验命令目录与注册表一致、相对文档链接有效；
`test_qatools_cli.py` 验证正式命令的帮助和转发。详细代码落点见[仓库地图](repository-map.md#扩展一个工具)。
