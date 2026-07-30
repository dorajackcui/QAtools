# CLI 使用指南

这份文档面向 agent、自动化脚本、批处理命令。

仓库中的 CLI 支持“缺参后交互输入”，但对 agent 来说不稳定，因此推荐始终使用非交互方式调用。

## 通用建议

- 始终显式传入必填参数，不要依赖脚本在终端里继续提问。
- 推荐显式指定 `--sheet` 和 `--start-row`，避免工作簿默认活动工作表变化带来不确定性。
- 推荐显式指定 `-o/--output`，这样输出路径稳定，后续 agent 更容易继续处理结果文件。
- 列参数统一使用 Excel 列字母，例如 `A`、`B`、`AA`。
- 如果 CLI 与 GUI 行为看起来不同，优先记住：GUI 的“自动识别工作表 / source / target 列”属于界面增强，CLI 不会自动帮你补这些参数。
- 除特别说明外，工具会生成新的 Excel 文件，不会覆盖原文件。

## 非交互调用约定

- `tools/term_pair_checker/extract_terms_from_excel.py`
  - 必填：`input_file`、`-c/--source-column`、`-t/--target-column`
  - 常用可选：`-s/--sheet`、`--start-row`、`--mark-style`、`--no-term-mark`、`--exclusion-config`、`--history-tb`、`--history-sheet`、`--history-source-column`、`--history-target-column`、`--history-start-row`、`-o/--output`
- `tools/tag_placeholder_checker/check_tags_and_placeholders.py`
  - 必填：`input_file`、`-c/--source-column`、`-t/--target-column`
  - 常用可选：`-s/--sheet`、`--start-row`、`--token-type`、`--angle-config`、`-o/--output`
- `tools/line_break_checker/check_line_breaks.py`
  - 必填：`input_file`、`-c/--source-column`、`-t/--target-column`
  - 常用可选：`-s/--sheet`、`--start-row`、`-o/--output`
- `tools/source_consistency_checker/check_source_consistency.py`
  - 必填：`input_file`、`-c/--source-column`、`-t/--target-column`
  - 常用可选：`-s/--sheet`、`--start-row`、`-o/--output`
- `tools/chinese_target_checker/check_chinese_target.py`
  - 必填：`input_file`、`-c/--source-column`、`-t/--target-column`
  - 常用可选：`-s/--sheet`、`--start-row`、`-o/--output`
- `tools/excel_line_splitter/split_excel_lines.py`
  - 必填：`input_file`、`-c/--source-column`、`-r/--result-column`
  - 常用可选：`-s/--sheet`、`--start-row`、`-o/--output`
- `tools/french_nbsp_restorer/restore_french_nbsp.py`
  - 必填：`input_file`、`-t/--target-column`
  - 常用可选：`-s/--sheet`、`-r/--result-column`、`--start-row`、`-o/--output`
- `tools/xbench_report_transformer/transform_xbench_report.py`
  - 必填：`input_file`
  - 常用可选：`-s/--sheet`、`-o/--output`
- `python3 -m phraseloom.cli export`
  - 必填：源 Excel 路径
  - 常用可选：`-o/--output`、`--source-col`、`--target-col`、`--context-col`、`--tag-config`、`--group-similar`
- `python3 -m phraseloom.cli restore`
  - 必填：翻译完成的 Strings 工作簿
  - 常用可选：`-o/--output`

## 推荐命令模板

### 1. 术语检查

最稳妥的 agent 调用方式：

```bash
python3 tools/term_pair_checker/extract_terms_from_excel.py ./input.xlsx \
  -s Sheet1 \
  -c A \
  -t B \
  --start-row 2 \
  --mark-style '【】' \
  -o ./artifacts/term_pair_check_input.xlsx
```

如果需要同时检查多种术语 mark：

```bash
python3 tools/term_pair_checker/extract_terms_from_excel.py ./input.xlsx \
  -s Sheet1 \
  -c A \
  -t B \
  --start-row 2 \
  --mark-style '【】' \
  --mark-style '[]' \
  -o ./artifacts/term_pair_check_input.xlsx
```

如果要切换误判排除规则：

```bash
python3 tools/term_pair_checker/extract_terms_from_excel.py ./input.xlsx \
  -s Sheet1 \
  -c A \
  -t B \
  --start-row 2 \
  --mark-style '【】' \
  --exclusion-config ./custom_term_exclusions.json \
  -o ./artifacts/term_pair_check_input.xlsx
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
  -o ./artifacts/term_pair_check_input.xlsx
```

历史 TB 会自动识别第 1 行的 `source` / `target` 列；也兼容本工具输出里的 `source术语` / `target术语` 表头。读取历史值时会去掉支持的 mark。传入历史 TB 后，会先把历史 TB 全量加入检查词库，再合并本批次新增术语；命中历史 source 时使用历史 target，未命中的 source 才按本批次第一次出现建立新增术语对。输出的 `术语表` 只写本次检查文本中实际涉及的历史术语和本批次新增术语。

如果不需要从 mark 提取新术语、只想用历史 TB 检查：

```bash
python3 tools/term_pair_checker/extract_terms_from_excel.py ./input.xlsx \
  -s Sheet1 -c A -t B \
  --start-row 2 \
  --no-term-mark \
  --history-tb ./glossary.xlsx \
  --history-sheet Glossary \
  --history-source-column A \
  --history-target-column B \
  -o ./artifacts/term_pair_check_input.xlsx
```

`--no-term-mark` 必须与 `--history-tb` 一起使用。该模式替代原“术语表命中检查”，输出统一为 `术语表` 和 `问题列`。

输出结果中会新增这些工作表：

- `术语表`：包含保留 mark、无 mark 和 `术语来源` 列；来源为 `历史TB` 或 `本批次新增`
- `问题列`

标准输出会打印：

- 工作表名
- source / target 列
- mark 类型
- 术语表条目数
- 问题条数
- 输出文件路径

### 2. Excel 分行拆列

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

### 3. Tag / Placeholder 检查

推荐完整调用：

```bash
python3 tools/tag_placeholder_checker/check_tags_and_placeholders.py ./input.xlsx \
  -s Sheet1 \
  -c A \
  -t B \
  --start-row 2 \
  --token-type angle \
  --token-type square_color \
  --token-type brace \
  --token-type newline \
  --token-type memoq \
  -o ./artifacts/tag_check_input.xlsx
```

说明：

- `<...>` 默认全量作为普通 tag 检查
- 方括号 color tag 会检查 `[color=...]` 和 `[/color]`
- memoQ tag 会按 `{n}`、`{n>`、`<n}` 单独检查，例如 `{1}{2>Glace du Néant<3}` 会提取为 `{1}`、`{2>`、`<3}`，不会作为普通 `{...}` placeholder 检查

如果只检查 `<...>` tag：

```bash
python3 tools/tag_placeholder_checker/check_tags_and_placeholders.py ./input.xlsx \
  -s Sheet1 \
  -c A \
  -t B \
  --start-row 2 \
  --token-type angle \
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
- 含方括号 color tag 行数
- 含花括号 placeholder 行数
- 含 `\n` mark 行数
- 含 memoQ tag 行数
- 问题行数 / 问题条数
- 输出文件路径

### 4. 换行数量检查

```bash
python3 tools/line_break_checker/check_line_breaks.py ./input.xlsx \
  -s Sheet1 \
  -c A \
  -t B \
  --start-row 2 \
  -o ./artifacts/line_break_check_input.xlsx
```

输出结果会新增 `换行数量问题` 工作表，列出问题行号、双边换行数、数量差和原文。该工具检查 Excel 单元格中的真实换行；文本形式的 `\n` 不算真实换行。

### 5. 同源译文一致性

```bash
python3 tools/source_consistency_checker/check_source_consistency.py ./input.xlsx \
  -s Sheet1 \
  -c A \
  -t B \
  --start-row 2 \
  -o ./artifacts/source_consistency_check_input.xlsx
```

输出结果会新增 `同源译文不一致` 工作表，列出每个不一致 source 组中的全部原始行。source 和 target 均按单元格文本精确比较；空 source 跳过，空 target 参与比较。

### 6. Target 中文检查

生成新的检查结果文件：

```bash
python3 tools/chinese_target_checker/check_chinese_target.py ./input.xlsx \
  -s Sheet1 \
  -c A \
  -t B \
  --start-row 2 \
  -o ./artifacts/input_chinese_target_checked.xlsx
```

说明：

- 原数据工作表保持不变，命中行写入独立的 `Target中文问题` 工作表。
- 问题表列为 `行号 / source原文 / target原文 / 问题描述 / 命中字符`。
- 检查范围包含汉字、CJK 标点、全角标点和常见中文排版符号，例如 `【】（）`、`，。！？`、`《》“”‘’·`；单字符 `—` 和 `…` 放行。
- 普通 ASCII 标点和全角英数不会单独触发标记。
- 已有的 `Target中文问题` 或旧版 `中文检查问题` 工作表会在运行时重建或移除，避免残留。
- 不传 `-o/--output` 时，默认生成 `target_chinese_check_<原文件名>`。

标准输出会打印：

- 工作表名
- source / target 列
- 开始行
- 处理行数
- 含中文行数
- 输出文件路径

### 7. 法语 NBSP 恢复

直接修复 target 列：

```bash
python3 tools/french_nbsp_restorer/restore_french_nbsp.py ./input.xlsx \
  -s Sheet1 \
  -t B \
  --start-row 2 \
  -o ./artifacts/input_french_nbsp_restored.xlsx
```

如果希望保留原 target，并把修复后的完整译文写入另一列：

```bash
python3 tools/french_nbsp_restorer/restore_french_nbsp.py ./input.xlsx \
  -s Sheet1 \
  -t B \
  -r C \
  --start-row 2 \
  -o ./artifacts/input_french_nbsp_restored.xlsx
```

说明：

- 会恢复 `;`、`:`、`?`、`!` 前的 NBSP。
- 会恢复 `«` 后和 `»` 前的 NBSP。
- 指定 `-r/--result-column` 后，不需要修复的 target 也会复制到结果列。
- 不会改写 URL 内标点和 `12:30` 这类时间冒号。

标准输出会打印：

- 工作表名
- target 列 / 结果列
- 开始行
- 处理行数
- 修复行数
- 输出文件路径

### 8. Xbench QA Report 转换

推荐完整调用：

```bash
python3 tools/xbench_report_transformer/transform_xbench_report.py ./Xbench_QA_Report.xlsx \
  -s "Xbench QA" \
  -o ./xbench_flat.xlsx
```

输出列固定为：

```text
文件名, key, source, target, QA问题
```

说明：

- `QA问题` 使用 `源术语 -> 目标术语：问题类型` 格式。
- 同一组内多个问题用中文分号 `；` 合并。
- `Metadata` 第一行作为 key，第二行作为文件名；没有 key 时按文件名+source 或 source 降级聚类。
- 默认输出文件名为 `xbench_transform_<原文件名>`。
- 工具会生成新的结果 Excel，不会覆盖原始 Xbench 报告；如果 `-o` 指向输入文件本身，会报错。

标准输出会打印：

- 工作表名
- 读取明细数
- 输出行数
- 输出文件路径

### 9. PhraseLoom Strings 工作流

导出待翻译 Strings：

```bash
python3 -m phraseloom.cli export ./source.xlsx \
  --source-col source \
  --target-col target \
  --context-col context \
  -o ./source_strings.xlsx
```

如需按相似结构调整显示顺序，增加 `--group-similar`。该选项只分组和排序，不会合并普通相似句或生成译文。

翻译完成后回填：

```bash
python3 -m phraseloom.cli restore ./source_strings.xlsx \
  -o ./source_translated.xlsx
```

Strings 工作簿内嵌原工作簿、行映射和 Tag 规则，因此回填时不需要再次传入原始 Excel。详细格式及校验规则见 [`../phraseloom/README.md`](../phraseloom/README.md)。

## Agent 调用注意事项

- 不要把 GUI 的自动列识别能力当成 CLI 的默认能力；CLI 场景下请自己明确传列字母。
- 不要省略位置参数，否则脚本可能进入交互提问模式。
- 如果要批量处理多个文件，建议为每次运行都显式传入独立输出文件名，避免后续步骤误读旧结果。
- 如果只需要默认输出命名，也可以省略 `-o`；默认命名分别是：
  - `term_pair_check_<原文件名>`
  - `tag_check_<原文件名>`
  - `target_chinese_check_<原文件名>`
  - `<原文件名>_split_lines.xlsx`
  - `<原文件名>_french_nbsp_restored.xlsx`
  - `xbench_transform_<原文件名>`
