# PhraseLoom

把 Excel 本地化文件压缩成最小待翻译内容，并用历史 TM 自动预填。

核心流程：

```text
历史完成稿 -> tm_pairs
本轮源文件 + 可选 tm_pairs -> 自包含 translator workbook
译者修改 to_translate / prefilled_units
translator workbook -> 回填后的原表格式副本
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

## 极简 GUI

启动桌面界面：

```bash
phraseloom gui
```

安装后也可以直接打开 `phraseloom-gui`。GUI 的“日常流程”只保留三个入口：

1. 准备翻译
2. 回填译文
3. 生成 TM

常用参数直接显示，输出路径、列名和 Tag 配置收在“更多选项”中。准备翻译时仍会明确显示
TM 文件选择和“使用当前 target 内容作为预填”选项；不会自动使用 TM。回填时只需选择
translator workbook。原有诊断与 Entity 命令全部放在“高级工具”页，功能与 CLI 对齐。
“准备翻译”和“生成 TM”都支持可选 Context 列；留空时会自动识别名为 `context` 的列。

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
a) Advanced tools
q) Quit
```

日常推荐用交互模式。简单理解：

```text
1 = 把历史完成稿做成可复用 tm_pairs
2 = 用本轮源文件 + 可选 tm_pairs 生成一个自包含 translator workbook
3 = 只选择译者交回的 workbook，生成回填后的原表格式 copy
```

第 2 步每次都会询问是否使用 TM，不会自动沿用上一次的 TM。检测到当前
target 列已有内容时，还会询问是否把这些内容作为 prefill。可复用模板的
最小变体数固定为 2，不再作为日常问题显示。

命令行批处理时用下面的命令。

## 1. 从历史 TM 提取 tm_pairs

历史 TM 是已经完成翻译的 Excel，必须有 source 列和 target 列。

```bash
phraseloom tm-extract '/path/to/tm.xlsx' \
  --source-col en \
  --target-col fr \
  --context-col screen_notes
```

默认输出到：

```text
/path/to/tm_l10n/tm_reusable_units.xlsx
```

`tm_pairs.xlsx` 后续可以反复复用，不需要每次重新生成。

## 2. 准备本轮 translator workbook

用历史 TM 预填本轮源文件：

```bash
phraseloom prepare '/path/to/source.xlsx' \
  --source-col 'source' \
  --target-col 'target' \
  --context-col 'context' \
  --tm '/path/to/tm_l10n/tm_reusable_units.xlsx'
```

如需把源文件当前 target 内容也作为预填，增加：

```bash
--use-existing-targets
```

默认只输出译者需要处理的自包含工作簿：

```text
/path/to/source_l10n/source_translator_todo.xlsx
```

这个工作簿内部保存了原工作簿副本以及回填所需的列配置。原工作表暂时隐藏，
译者主要处理两个可见 sheet：

- `to_translate`：尚未命中的内容，填写 `target`。
- `prefilled_units`：TM 或当前 target 已预填的内容，保持可见并允许人工修改。

回填时会同时读取两个 sheet，以人工修改后的内容为准。

## 3. 回填译者交付的 translator workbook

只需要提供译者工作簿，不再需要原文件路径、列名、回填模式或 reusable part：

```bash
phraseloom fill '/path/to/source_l10n/source_translator_todo.xlsx'
```

默认输出：

```text
/path/to/source_l10n/source_filled_result.xlsx
```

默认恢复原工作表结构、把结果写入 target 列，并移除 PhraseLoom 的内部 sheet。
不会覆盖真正的原始源文件。

只有存在 warning 或未回填行时，才会额外生成检查 workbook：

```text
/path/to/source_l10n/source_filled_result_restore_audit.xlsx
```

检查 workbook 包含 `summary` 和 `restore_warnings`，只列出 warning / unfilled
行。可以用 `--audit-output` 主动指定检查文件路径。

旧的 `extract`、`fill SOURCE.xlsx --templates ...` 和 `report` 流程继续保留为
高级兼容命令，但不再是日常推荐入口。

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
# 选择 a) Advanced tools，再选择 1) Entity workflow

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

高级 `extract` 命令生成的诊断过程包，包含 summary、translation_units、
source_map、filled_workbook、qa_report 等详细信息。日常 `prepare` 不生成它。

`*_translator_todo.xlsx`

给译者的自包含工作簿。内部保存原工作表副本和回填参数；`to_translate` 里是
需要翻译的空 target，`prefilled_units` 里是 TM 或当前 target 已预填的内容。
两个 sheet 都会参与回填，`prefilled_units` 保持可见且允许修改。

`*_filled_result.xlsx`

回填后的原表格式副本。原工作表会恢复可见，PhraseLoom 的翻译和元数据
sheet 会从交付文件中移除。

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
