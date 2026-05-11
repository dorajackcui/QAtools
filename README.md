# PhraseLoom

把 Excel 本地化文件压缩成最小待翻译内容，并用历史 TM 自动预填。

核心流程：

```text
历史完成稿 -> tm_pairs
本轮源文件 + tm_pairs -> tm_prefill_pack + translator_todo
译者翻译 to_translate
to_translate -> 回填源文件
```

## 开发环境

推荐使用虚拟环境安装为可编辑包：

```bash
python -m venv .venv
.venv\Scripts\python -m pip install -e .
```

运行测试：

```bash
python -m unittest discover -v
```

## 项目结构

核心代码位于 `phraseloom/` 包中。`template_demo.py` 和 `entity_cluster_probe.py` 保留为兼容入口；新代码应优先从 `phraseloom.workflow`、`phraseloom.template_engine`、`phraseloom.excel_io`、`phraseloom.entity_cluster` 等模块导入。

安装后可以直接使用 `phraseloom` 命令；未安装时仍可使用兼容入口 `python3 template_demo.py`。

## 快速开始

进入工具目录：

```bash
cd /path/to/phrase-loom
phraseloom
```

交互菜单：

```text
Localization Workflow

1) Build TM from completed Excel
2) Prepare translator file for new source
3) Fill source from translated file
q) Quit
```

日常推荐用交互模式。简单理解：

```text
1 = 把历史完成稿做成可复用 tm_pairs
2 = 用本轮源文件 + tm_pairs 生成译者要翻译的 to_translate
3 = 译者交回 to_translate 后，把译文回填到源文件 copy
```

如果你已经有 `tm_pairs.xlsx`，不用再做第 1 步，直接选第 2 步并填入已有 `tm_pairs` 路径即可。

命令行批处理时用下面的命令。

## 1. 从历史 TM 提取 tm_pairs

历史 TM 是已经完成翻译的 Excel，必须有 source 列和 target 列。

```bash
phraseloom tm-extract '/path/to/tm.xlsx' \
  --source-col en \
  --target-col fr
```

默认输出到：

```text
/path/to/tm_l10n/tm_reusable_units.xlsx
```

`tm_pairs.xlsx` 后续可以反复复用，不需要每次重新生成。

## 2. 提取本轮待翻译内容

用历史 TM 预填本轮源文件：

```bash
phraseloom extract '/path/to/source.xlsx' \
  --source-col '英語' \
  --target-col - \
  --tm '/path/to/tm_l10n/tm_reusable_units.xlsx' \
  --no-existing-targets
```

默认输出到源文件旁边的工作目录：

```text
/path/to/source_l10n/source_tm_prefill_pack.xlsx
/path/to/source_l10n/source_translator_todo.xlsx
```

译者只需要处理：

```text
source_translator_todo.xlsx
```

只填写 `to_translate` sheet 里的 `target_unit` 列即可。

## 3. 回填译者交付的 to_translate

### 报告模式

生成检查用 workbook，不改源文件结构：

```bash
phraseloom fill '/path/to/source.xlsx' \
  --templates '/path/to/source_l10n/source_translator_todo.xlsx' \
  --source-col '英語' \
  --target-col '法语' \
  --mode report
```

### 交付模式

生成一个源文件 copy，并把结果写进 target 列：

```bash
phraseloom fill '/path/to/source.xlsx' \
  --templates '/path/to/source_l10n/source_translator_todo.xlsx' \
  --source-col '英語' \
  --target-col '法语' \
  --mode target-column
```

默认输出：

```text
/path/to/source_l10n/source_filled_result.xlsx
```

不会覆盖原始源文件。

## 输出说明

`*_reusable_units.xlsx`

历史 TM 的可复用翻译单元。

`*_tm_prefill_pack.xlsx`

过程包，包含 summary、translation_units、source_map、filled_workbook、qa_report 等详细信息。

`*_translator_todo.xlsx`

给译者的独立文件。只需要翻译 `target_unit` 为空的行。

`*_filled_result.xlsx`

回填后的交付文件，`target-column` 模式会把译文写进指定 target 列。

## 自动处理规则

工具会自动：

```text
1. 去重重复 source segment
2. 抽取可复用 template
3. 用历史 TM 预填命中的 unit
4. 纯数字/纯符号内容直接回填 source 本身
5. 保留变量，例如 {a}、{num1}、{stage1}
```

例如：

```text
Obtain 3 stars in Chapter 1
Obtain 5 stars in Chapter 2
```

会抽成：

```text
Obtain {num1} stars in Chapter {num2}
```

纯数字/符号例如下面这些不会进入待翻译表：

```text
123
...
10-20
1000%
$5.99
{0}
```

## 常看指标

在 `summary` 里重点看：

```text
total_source_rows            源文件总行数
total_translation_units      去重/抽 template 后的总 unit 数
already_filled_units         已自动填好的 unit 数
already_filled_source_rows   已自动填好的源文件行数
units_to_translate           译者需要翻译的 unit 数
source_rows_to_translate     这些 unit 覆盖的源文件行数
```

## 脱敏实测示例

使用：

```text
TM: /path/to/completed_tm.xlsx
源文件: /path/to/source.xlsx
```

生成：

```text
/path/to/completed_tm_l10n/completed_tm_reusable_units.xlsx
/path/to/source_l10n/source_translator_todo.xlsx
```

结果摘要：

```text
total_source_rows: <source row count>
total_translation_units: <unit count after dedupe/template extraction>
already_filled_units: <prefilled unit count>
already_filled_source_rows: <prefilled source row count>
units_to_translate: <unit count that still needs translation>
source_rows_to_translate: <source rows covered by untranslated units>
```
