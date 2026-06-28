# Excel 术语表命中检查工具

用于读取一份术语表 Excel 和一份双语检查 Excel，检查每一行 `source` 中出现的术语，是否在 `target` 中按术语表要求进行了翻译，并输出为新的 Excel 文件。

处理完成后会生成新的检查结果 Excel，并新增两个工作表：

- `术语命中问题`：列出命中术语但未按术语表翻译的问题
- `检查汇总`：列出统计信息和术语表冲突项

## 规则

- 输入为两份 Excel：
  - 术语表：`source` / `target`
  - 检查文本：`source` / `target`
- 默认匹配方式为“混合边界”
- `source` 行里只要命中术语，就要求 `target` 至少包含一次对应译法
- 复数处理：如果 `source` 命中整条 `术语+s`，且 `target` 也命中整条 `译法+s`，视为通过；其他复数形态疑似变体不进入问题报告
- 同一行同一术语重复出现时，只按“是否出现过”检查一次
- 默认忽略大小写，可通过参数或 GUI 开关改为大小写敏感
- 对英文/数字术语会按边界匹配，避免短词误击中更长单词
- `rain` 不命中 `training`
- `ACC` 不命中 `account`、`accuracy`
- `API` 可以命中 `API-key`
- `ACC` 不命中 `ACC_001`
- 重叠术语按最长优先，例如命中 `API key` 时，不再额外要求重叠区间内的 `API`
- 同一 `source` 对应多个不同 `target` 时，记为术语表冲突，不参与自动判错
- 可通过 `substring` 模式回退到旧的纯包含逻辑

## 运行方式

```bash
python3 tools/term_glossary_checker/check_terms_against_glossary.py glossary.xlsx data.xlsx \
  --glossary-source-column A --glossary-target-column B \
  --data-source-column A --data-target-column B \
  --match-mode hybrid-boundary
```

图形界面：

```bash
python3 tools/term_glossary_checker/check_terms_against_glossary_gui.py
```

GUI 现在支持：

- 术语表文件和检查文件都会自动读取工作表列表，用下拉框选择工作表
- 默认输出为新的 Excel 文件，也可手动指定输出路径
- 两份文件都会自动识别第 1 行表头中的 `source` / `target` 列并回填
- 如果未识别到列，可继续手动填写列字母作为回退
- 可勾选“使用 Codex 筛查术语误报”，在 `术语命中问题` 末尾写入 `fp_*` 辅助列

## 常用参数

- `--glossary-sheet`：术语表工作表名称
- `--glossary-source-column`：术语表 source 列
- `--glossary-target-column`：术语表 target 列
- `--data-sheet`：检查文本工作表名称
- `--data-source-column`：检查文本 source 列
- `--data-target-column`：检查文本 target 列
- `--start-row`：开始处理的行号，默认 `2`
- `--case-sensitive`：启用大小写敏感匹配
- `--match-mode`：匹配模式，可选 `hybrid-boundary` 或 `substring`
- `--codex-fp-review`：检查完成后调用本机 `codex exec` 做假阳性筛查，并在 `术语命中问题` 末尾写入 `fp_decision`、`fp_category`、`fp_confidence`、`fp_note`、`fp_by`
- `--codex-fp-sample-size`：每个同术语、期望译法、问题类型和原文/译文文本 cluster 发送给 Codex 的样本数，默认 `5`
- `--codex-model`：Codex 假阳性筛查使用的模型；不填则使用 Codex 默认模型
- `--codex-reasoning-effort`：Codex 假阳性筛查使用的 reasoning effort，默认 `high`
- `-o, --output`：输出文件路径，可选，默认生成 `<原文件名>_glossary_checked.xlsx`

## 输出说明

`术语命中问题` 工作表包含以下列：

- `行号`
- `问题类型`
- `source术语`
- `期望target术语`
- `source文本`
- `target文本`

启用 `--codex-fp-review` 或 GUI 勾选 Codex 筛查后，还会追加：

- `fp_decision`：`false_positive`、`true_issue` 或 `review`
- `fp_category`
- `fp_confidence`
- `fp_note`
- `fp_by`

`检查汇总` 工作表包含以下统计项：

- 术语表工作表 / 检查工作表
- 术语表和检查表的 source / target 列
- 开始行
- 大小写模式
- 匹配模式
- 总行数
- 命中术语行数
- 问题行数
- 问题条数
- 术语表条数
- 冲突术语数
