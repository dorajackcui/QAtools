# CLI 使用指南

这份文档面向 agent、自动化脚本、批处理命令。

仓库中的 CLI 支持“缺参后交互输入”，但对 agent 来说不稳定，因此推荐始终使用非交互方式调用。

## 通用建议

- 始终显式传入必填参数，不要依赖脚本在终端里继续提问。
- 推荐显式指定 `--sheet` 和 `--start-row`，避免工作簿默认活动工作表变化带来不确定性。
- 推荐显式指定 `-o/--output`，这样输出路径稳定，后续 agent 更容易继续处理结果文件。
- 列参数统一使用 Excel 列字母，例如 `A`、`B`、`AA`。
- 如果 CLI 与 GUI 行为看起来不同，优先记住：GUI 的“自动识别工作表 / source / target 列”属于界面增强，CLI 不会自动帮你补这些参数。
- 所有工具都会生成新的 Excel 文件，不会覆盖原文件。

## 非交互调用约定

- `tools/term_pair_checker/extract_terms_from_excel.py`
  - 必填：`input_file`、`-c/--source-column`、`-t/--target-column`
  - 常用可选：`-s/--sheet`、`--start-row`、`--mark-style`、`--exclusion-config`、`--history-tb`、`--history-sheet`、`--history-source-column`、`--history-target-column`、`--history-start-row`、`-o/--output`
- `tools/term_glossary_checker/check_terms_against_glossary.py`
  - 必填：`glossary_file`、`data_file`、`--glossary-source-column`、`--glossary-target-column`、`--data-source-column`、`--data-target-column`
  - 常用可选：`--glossary-sheet`、`--data-sheet`、`--start-row`、`--case-sensitive`、`--match-mode`、`-o/--output`
- `tools/tag_placeholder_checker/check_tags_and_placeholders.py`
  - 必填：`input_file`、`-c/--source-column`、`-t/--target-column`
  - 常用可选：`-s/--sheet`、`--start-row`、`--token-type`、`--angle-config`、`-o/--output`
- `tools/excel_line_splitter/split_excel_lines.py`
  - 必填：`input_file`、`-c/--source-column`、`-r/--result-column`
  - 常用可选：`-s/--sheet`、`--start-row`、`-o/--output`

## 推荐命令模板

### 1. 术语对检查

最稳妥的 agent 调用方式：

```bash
python3 tools/term_pair_checker/extract_terms_from_excel.py ./input.xlsx \
  -s Sheet1 \
  -c A \
  -t B \
  --start-row 2 \
  --mark-style '【】' \
  -o ./artifacts/input_term_pairs.xlsx
```

如果需要同时检查多种 tag：

```bash
python3 tools/term_pair_checker/extract_terms_from_excel.py ./input.xlsx \
  -s Sheet1 \
  -c A \
  -t B \
  --start-row 2 \
  --mark-style '【】' \
  --mark-style '[]' \
  --mark-style '<>' \
  -o ./artifacts/input_term_pairs.xlsx
```

如果要切换误判排除规则：

```bash
python3 tools/term_pair_checker/extract_terms_from_excel.py ./input.xlsx \
  -s Sheet1 \
  -c A \
  -t B \
  --start-row 2 \
  --mark-style '【】' \
  --exclusion-config ./tools/term_pair_checker/false_positive_exclusions.json \
  -o ./artifacts/input_term_pairs.xlsx
```

如果要优先复用历史 TB：

```bash
python3 tools/term_pair_checker/extract_terms_from_excel.py ./input.xlsx \
  -s Sheet1 \
  -c A \
  -t B \
  --start-row 2 \
  --mark-style '【】' \
  --history-tb ./history_tb.xlsx \
  --history-sheet Glossary \
  -o ./artifacts/input_term_pairs.xlsx
```

历史 TB 会自动识别第 1 行的 `source` / `target` 列；也兼容本工具输出里的 `source术语` / `target术语` 表头。读取历史值时会去掉支持的 mark。命中历史 source 时优先使用历史 target，未命中的 source 才按本批次第一次出现建立新增术语对。

输出结果中会新增这些工作表：

- `术语表`：包含保留 mark、无 mark 和 `术语来源` 列；来源为 `历史TB` 或 `本批次新增`
- `问题列`

标准输出会打印：

- 工作表名
- source / target 列
- mark 类型
- 术语表条目数
- 问题行数
- 输出文件路径

### 2. 术语表命中检查

推荐完整调用：

```bash
python3 tools/term_glossary_checker/check_terms_against_glossary.py \
  ./glossary.xlsx \
  ./data.xlsx \
  --glossary-sheet Glossary \
  --glossary-source-column A \
  --glossary-target-column B \
  --data-sheet Sheet1 \
  --data-source-column A \
  --data-target-column B \
  --start-row 2 \
  --match-mode hybrid-boundary \
  -o ./artifacts/data_glossary_checked.xlsx
```

如果需要大小写敏感：

```bash
python3 tools/term_glossary_checker/check_terms_against_glossary.py \
  ./glossary.xlsx \
  ./data.xlsx \
  --glossary-sheet Glossary \
  --glossary-source-column A \
  --glossary-target-column B \
  --data-sheet Sheet1 \
  --data-source-column A \
  --data-target-column B \
  --start-row 2 \
  --case-sensitive \
  --match-mode hybrid-boundary \
  -o ./artifacts/data_glossary_checked.xlsx
```

输出结果中会新增：

- `术语命中问题`
- `检查汇总`

标准输出会打印：

- 术语表工作表 / 检查工作表
- 大小写模式
- 匹配模式
- 术语表条数
- 冲突术语数
- 总行数
- 命中术语行数
- 问题行数 / 问题条数
- 输出文件路径

### 3. Excel 分行拆列

推荐完整调用：

```bash
python3 tools/excel_line_splitter/split_excel_lines.py ./input.xlsx \
  -s Sheet1 \
  -c A \
  -r B \
  --start-row 2 \
  -o ./artifacts/input_split_lines.xlsx
```

标准输出会打印：

- 工作表名
- 源列 / 结果列
- 开始行
- 写入条目数
- 输出文件路径

### 4. Tag / Placeholder 检查

推荐完整调用：

```bash
python3 tools/tag_placeholder_checker/check_tags_and_placeholders.py ./input.xlsx \
  -s Sheet1 \
  -c A \
  -t B \
  --start-row 2 \
  --token-type angle \
  --token-type brace \
  --token-type newline \
  -o ./artifacts/input_tag_placeholder_checked.xlsx
```

说明：

- `<...>` 默认不是全量检查，而是按 `tools/term_pair_checker/false_positive_exclusions.json` 中的模式识别真正需要校验的 tag
- 这份默认模式和术语配对工具共用同一个文件，因为那批被 `<>` mark 排除的内容正是 tag 检查应该关注的对象

如果只检查 `<...>` tag：

```bash
python3 tools/tag_placeholder_checker/check_tags_and_placeholders.py ./input.xlsx \
  -s Sheet1 \
  -c A \
  -t B \
  --start-row 2 \
  --token-type angle \
  --angle-config ./tools/term_pair_checker/false_positive_exclusions.json \
  -o ./artifacts/input_tag_checked.xlsx
```

输出结果中会新增：

- `标签占位问题`
- `检查汇总`

标准输出会打印：

- 检查工作表
- 检查类型
- 总行数
- 命中检查类型行数
- 含尖括号 tag 行数
- 含花括号 placeholder 行数
- 含 `\n` mark 行数
- 问题行数 / 问题条数
- 输出文件路径

## Agent 调用注意事项

- 不要把 GUI 的自动列识别能力当成 CLI 的默认能力；CLI 场景下请自己明确传列字母。
- 不要省略位置参数，否则脚本可能进入交互提问模式。
- 如果要批量处理多个文件，建议为每次运行都显式传入独立输出文件名，避免后续步骤误读旧结果。
- 如果只需要默认输出命名，也可以省略 `-o`；默认命名分别是：
  - `<原文件名>_term_pairs.xlsx`
  - `<原文件名>_glossary_checked.xlsx`
  - `<原文件名>_tag_placeholder_checked.xlsx`
  - `<原文件名>_split_lines.xlsx`
