# Toolshub

这个仓库现在是一个轻量的 Excel `toolshub`，每个工具都有独立目录、独立说明，同时继续共用根目录依赖。

如果你是通过 agent / 脚本来调用这些工具，建议优先查看独立的 CLI 文档：[CLI 使用指南](docs/cli-usage.md)。

## 当前工具

### 1. 术语对检查

- 目录：`tools/term_pair_checker`
- 用途：从 Excel 的 `source` / `target` 列提取术语并检查是否对齐
- tag 支持：`【】`、`[]`、`<>`，且可在 GUI 中多选组合检查
- 检查规则：`术语表` 保留 tag，实际术语检查会忽略 tag，并回溯整表未标注出现
- 复数处理：回扫时双边整条 `术语+s` / `译法+s` 直接放行；其他复数签名疑似命中会在问题简述标注“疑似复数变体”
- 历史 TB：可选选择历史 TB；选择后会用“历史 TB 全量 + 本批次新增 TB”一起检查
- 误判排除：默认通过 `tools/term_pair_checker/false_positive_exclusions.json` 排除 `</>`、`<color=...>` 这类伪标签
- 输出增强：结果会给出合并后的 `术语表`（只包含本次检查文本中涉及的术语，并包含保留 mark、无 mark 和术语来源列），以及带原文上下文的 `问题列`
- GUI 增强：自动读取工作表列表，并尝试自动识别本批次和历史 TB 的 `source` / `target` 列
- 输出方式：生成新的结果 Excel，不覆盖原文件
- CLI：

```bash
python3 tools/term_pair_checker/extract_terms_from_excel.py input.xlsx \
  -s Sheet1 \
  -c A \
  -t B \
  --start-row 2 \
  --mark-style '【】' \
  --history-tb history_tb.xlsx \
  -o output_term_pairs.xlsx
```

- 兼容旧入口：

```bash
python3 extract_terms_from_excel.py input.xlsx -c A -t B
```

- GUI：

```bash
python3 tools/term_pair_checker/extract_terms_gui.py
```

详情见 `tools/term_pair_checker/README.md`。

### 2. 术语表命中检查

- 目录：`tools/term_glossary_checker`
- 用途：使用术语表检查双语 Excel 中的 source 是否按约定 target 进行了翻译
- 复数处理：双边整条 `术语+s` / `译法+s` 直接放行；其他复数签名疑似命中会在问题类型标注“疑似复数变体”
- GUI 增强：术语表文件和检查文件都支持工作表下拉选择与 `source` / `target` 列自动识别
- 输出方式：生成新的结果 Excel，不覆盖原文件
- CLI：

```bash
python3 tools/term_glossary_checker/check_terms_against_glossary.py glossary.xlsx data.xlsx \
  --glossary-sheet Glossary \
  --glossary-source-column A --glossary-target-column B \
  --data-sheet Sheet1 \
  --data-source-column A --data-target-column B
```

- GUI：

```bash
python3 tools/term_glossary_checker/check_terms_against_glossary_gui.py
```

详情见 `tools/term_glossary_checker/README.md`。

### 3. Excel 分行拆列

- 目录：`tools/excel_line_splitter`
- 用途：把指定列中带回车的单元格内容拆开，并连续写入结果列
- GUI 增强：自动读取工作表列表并用下拉框选择
- 输出方式：生成新的结果 Excel，不覆盖原文件
- CLI：

```bash
python3 tools/excel_line_splitter/split_excel_lines.py input.xlsx \
  -s Sheet1 \
  -c A \
  -r B \
  --start-row 2 \
  -o output_split_lines.xlsx
```

- GUI：

```bash
python3 tools/excel_line_splitter/split_excel_lines_gui.py
```

详情见 `tools/excel_line_splitter/README.md`。

### 4. Tag / Placeholder 检查

- 目录：`tools/tag_placeholder_checker`
- 用途：逐行检查双语 Excel 中 `source` / `target` 的 `<...>`、`{...}`、`\n` 和数字 tag 是否一致
- 检查类型：支持 `<...>` tag、`{...}` placeholder、`\n` mark 与 `{n}` / `{n>` / `<n}` 数字 tag，可单独或组合检查
- `<...>` 识别：默认直接复用 `tools/term_pair_checker/false_positive_exclusions.json` 识别真正需要校验的 tag，避免维护两份规则
- GUI 增强：自动读取工作表列表，并尝试自动识别 `source` / `target` 列
- 输出方式：生成新的结果 Excel，不覆盖原文件
- CLI：

```bash
python3 tools/tag_placeholder_checker/check_tags_and_placeholders.py input.xlsx \
  -s Sheet1 \
  -c A \
  -t B \
  --start-row 2 \
  --token-type angle \
  --token-type brace \
  --token-type newline \
  --token-type numeric
```

- GUI：

```bash
python3 tools/tag_placeholder_checker/check_tags_and_placeholders_gui.py
```

详情见 `tools/tag_placeholder_checker/README.md`。

## 依赖安装

```bash
pip install -r requirements.txt
```

## 文档导航

- Agent / 脚本调用请看：[CLI 使用指南](docs/cli-usage.md)
- 术语对检查详细说明：`tools/term_pair_checker/README.md`
- 术语表命中检查详细说明：`tools/term_glossary_checker/README.md`
- Tag / Placeholder 检查详细说明：`tools/tag_placeholder_checker/README.md`
- Excel 分行拆列详细说明：`tools/excel_line_splitter/README.md`

## 统一 GUI 入口

```bash
python3 toolshub_gui.py
```

会打开一个统一窗口，使用标签页管理这些工具；Workflow 编排页也支持给术语对检查选择历史 TB；原有各自的 GUI 入口仍然保留。
