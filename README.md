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

核心代码位于 `phraseloom/` 包中。`template_demo.py` 和 `entity_cluster_probe.py` 保留为兼容入口；新代码应优先从 `phraseloom.workflow`、`phraseloom.template_engine`、`phraseloom.excel_io`、`phraseloom.entity_workflow`、`phraseloom.entity_cluster` 等模块导入。

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
  --source-col 'source' \
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

只填写 `to_translate` sheet 里的 `target` 列即可。

## 3. 回填译者交付的 to_translate

### 报告模式

生成检查用 workbook，不改源文件结构：

```bash
phraseloom fill '/path/to/source.xlsx' \
  --templates '/path/to/source_l10n/source_translator_todo.xlsx' \
  --source-col 'source' \
  --target-col '法语' \
  --mode report
```

### 交付模式

生成一个源文件 copy，并把结果写进 target 列：

```bash
phraseloom fill '/path/to/source.xlsx' \
  --templates '/path/to/source_l10n/source_translator_todo.xlsx' \
  --source-col 'source' \
  --target-col '法语' \
  --mode target-column
```

默认输出：

```text
/path/to/source_l10n/source_filled_result.xlsx
```

不会覆盖原始源文件。

## 独立 Entity Engine

Entity engine 是 translator todo 的二次处理工具，不属于主 tag/template 流水线。推荐使用四步简化流程；下面展示的是默认输出路径：

```text
/path/to/TM_reusable_units.xlsx
-> entity-tm
   -> /path/to/TM_reusable_units_l10n/TM_reusable_units_entity_memory.xlsx

/path/to/source_translator_todo.xlsx + 可选 /path/to/TM_reusable_units_l10n/TM_reusable_units_entity_memory.xlsx
-> entity-prepare
   -> /path/to/source_translator_todo_l10n/source_translator_todo_entity_pack.xlsx

人工确认 source_translator_todo_entity_pack.xlsx 里的 entity_structures / entity_terms
-> entity-fill-pack
   -> /path/to/source_translator_todo_l10n/source_translator_todo_entity_pack_filled.xlsx

/path/to/source_translator_todo_l10n/source_translator_todo_entity_pack_filled.xlsx
-> entity-merge-pack
   -> /path/to/source_translator_todo_l10n/source_translator_todo_entity_pack_filled_merged_todo.xlsx
```

也可以用交互式菜单进入 entity workflow：

```bash
phraseloom
# 选择 4) Entity workflow

phraseloom entity
phraseloom entity-interactive
```

交互菜单只覆盖推荐的四步主流程；底层调试命令仍通过直接 CLI 子命令使用。

从预处理过的 TM workbook 中创建 entity memory：

```bash
phraseloom entity-tm '/path/to/TM_reusable_units.xlsx'
```

默认输出：

```text
/path/to/TM_reusable_units_l10n/TM_reusable_units_entity_memory.xlsx
```

准备本轮 source entity pack，并可选用 entity memory 预填：

```bash
phraseloom entity-prepare '/path/to/source_translator_todo.xlsx' \
  --tm '/path/to/TM_reusable_units_l10n/TM_reusable_units_entity_memory.xlsx'
```

默认输出：

```text
/path/to/source_translator_todo_l10n/source_translator_todo_entity_pack.xlsx
```

译员或 PM 主要处理 `source_translator_todo_entity_pack.xlsx` 里的可见 sheet：

```text
related_units
non_related_units
entity_structures
entity_terms
```

`_entity_map` 和 `_metadata` 是隐藏的内部 sheet，正常不需要编辑。

译员主要填写 `entity_structures.target_structure` 和 `entity_terms.target_entity`。
`confidence`、`risk` 会隐藏，`status` 不需要填写；目标列非空就代表可用于回填，留空就表示该结构或实体暂不通过。
`entity_structures` 和 `entity_terms` 都会带 `sample_sources` / `sample_context`，方便回看来源上下文。

把已填写的实体结构和实体词表组合回 `related_units.target`：

```bash
phraseloom entity-fill-pack '/path/to/source_translator_todo_l10n/source_translator_todo_entity_pack.xlsx'
```

默认输出：

```text
/path/to/source_translator_todo_l10n/source_translator_todo_entity_pack_filled.xlsx
```

把 `related_units` 和 `non_related_units` 合并回完整 translator todo：

```bash
phraseloom entity-merge-pack '/path/to/source_translator_todo_l10n/source_translator_todo_entity_pack_filled.xlsx'
```

默认输出：

```text
/path/to/source_translator_todo_l10n/source_translator_todo_entity_pack_filled_merged_todo.xlsx
```

最后继续使用现有 `fill` 命令，把 merged todo 回填到目标文件。

高级或调试时仍可使用底层命令：`entity-split`、`entity-extract-tm`、`entity-prefill`、`entity-fill`、`entity-merge`。

流程图可打开：

```text
docs/entity-engine-flow.html
```

## 输出说明

`*_reusable_units.xlsx`

历史 TM 的可复用翻译单元。

`*_tm_prefill_pack.xlsx`

过程包，包含 summary、translation_units、source_map、filled_workbook、qa_report 等详细信息。

`*_translator_todo.xlsx`

给译者的独立文件。只需要翻译 `target` 为空的行。

`*_filled_result.xlsx`

回填后的交付文件，`target-column` 模式会把译文写进指定 target 列。

`*_entity_memory.xlsx`

`entity-tm` 从预处理过的 `*_reusable_units.xlsx` 中创建的 entity memory。默认输出到 `*_l10n/` 工作目录。

`*_entity_pack.xlsx`

`entity-prepare` 创建的单一 entity pack，包含 `related_units`、`non_related_units`、`entity_structures`、`entity_terms`，以及隐藏的内部 sheet。

`*_entity_pack_filled.xlsx`

`entity-fill-pack` 把已填写的实体结构和实体词表组合回 `related_units.target` 后的 pack。

`*_merged_todo.xlsx`

`entity-merge-pack` 把 `related_units` 和 `non_related_units` 合并后的完整 translator todo，可继续交给现有 `fill` 命令。

高级或调试命令可能产生下面这些底层中间文件：

`*_entity_related.xlsx`

`entity-split` 拆出的实体相关 todo 子集，包含 `entity_structures`、`entity_terms` 和内部映射表。

`*_not_entity_related.xlsx`

`entity-split` 拆出的非实体相关 todo 子集，可交给普通翻译流程或其他 workflow 并行处理。

`*_entity_tm.xlsx`

`entity-extract-tm` 从预处理过的 `tm_pairs` 中抽出的 entity structure / entity term 复用表。

`*_entity_prefilled.xlsx`

`entity-prefill` 用 entity TM 预填后的实体相关 workbook。TM prefill 只填结构表和实体表，不直接写完整 todo 的 `target`。

`*_entity_filled.xlsx`

`entity-fill` 把 ready 的结构和实体组合回 `target` 后的实体相关 workbook。

## 自动处理规则

工具会自动：

```text
1. 去重重复 source segment
2. 抽取可复用 template
3. 用历史 TM 预填命中的 unit
4. 纯数字/纯符号内容直接回填 source 本身
5. 保留变量，例如 {a}、{num1}、{stage1}
```

Entity engine 额外会自动：

```text
1. 从预处理后的 TM 中创建 entity memory
2. 从预处理后的 todo 中创建单一 entity pack，内含 related_units / non_related_units
3. 可用 entity memory 预填 entity_structures / entity_terms
4. 只在结构和所有实体都有非空目标值时，组合出 related_units.target
5. 按 original_index 把 related_units 和 non_related_units 合并回完整 todo
```

底层调试命令也可以把这些步骤拆成 `entity-split`、`entity-extract-tm`、`entity-prefill`、`entity-fill`、`entity-merge` 的多个文件。

Entity engine 不会：

```text
1. 调用 tag_engine 或 template_engine
2. 直接读取原始源 Excel
3. 直接写最终交付 workbook
4. 自动翻译非 entity 内容
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
