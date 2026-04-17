# Toolshub

这个仓库现在是一个轻量的 Excel `toolshub`，每个工具都有独立目录、独立说明，同时继续共用根目录依赖。

## 当前工具

### 1. 术语对检查

- 目录：`tools/term_pair_checker`
- 用途：从 Excel 的 `source` / `target` 列提取术语并检查是否对齐
- tag 支持：`【】`、`[]`、`<>`，且可在 GUI 中多选组合检查
- 检查规则：`术语表` 保留 tag，实际术语检查会忽略 tag，并回溯整表未标注出现
- 误判排除：默认通过 `tools/term_pair_checker/false_positive_exclusions.json` 排除 `</>`、`<color=...>` 这类伪标签
- 输出增强：结果会同时给出保留 mark 的 `术语表`、无 mark 的 `术语表（无mark）`，以及带原文上下文的 `问题列`
- GUI 增强：自动读取工作表列表，并尝试自动识别 `source` / `target` 列
- 输出方式：生成新的结果 Excel，不覆盖原文件
- CLI：

```bash
python3 tools/term_pair_checker/extract_terms_from_excel.py input.xlsx -c A -t B
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
- GUI 增强：术语表文件和检查文件都支持工作表下拉选择与 `source` / `target` 列自动识别
- 输出方式：生成新的结果 Excel，不覆盖原文件
- CLI：

```bash
python3 tools/term_glossary_checker/check_terms_against_glossary.py glossary.xlsx data.xlsx \
  --glossary-source-column A --glossary-target-column B \
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
python3 tools/excel_line_splitter/split_excel_lines.py input.xlsx -c A -r B --start-row 2
```

- GUI：

```bash
python3 tools/excel_line_splitter/split_excel_lines_gui.py
```

详情见 `tools/excel_line_splitter/README.md`。

## 依赖安装

```bash
pip install -r requirements.txt
```

## 统一 GUI 入口

```bash
python3 toolshub_gui.py
```

会打开一个统一窗口，使用标签页管理这三个工具；原有各自的 GUI 入口仍然保留。
