# Toolshub

这个仓库现在是一个轻量的 Excel `toolshub`，每个工具都有独立目录、独立说明，同时继续共用根目录依赖。

如果你是通过 agent / 脚本来调用这些工具，建议优先查看独立的 CLI 文档：[CLI 使用指南](docs/cli-usage.md)。

## 当前工具

### 1. 术语对检查

- 目录：`tools/term_pair_checker`
- 用途：从 Excel 的 `source` / `target` 列提取术语并检查是否对齐
- tag 支持：`【】`、`[]`、`<>`，且可在 GUI 中多选组合检查
- 检查规则：`术语表` 保留 tag，实际术语检查会忽略 tag，并回溯整表未标注出现
- 复数处理：回扫时双边整条 `术语+s` / `译法+s` 直接放行；其他复数形态疑似变体不进入问题报告
- 历史 TB：可选选择历史 TB；选择后会用“历史 TB 全量 + 本批次新增 TB”一起检查
- 误判排除：默认通过 `tools/term_pair_checker/false_positive_exclusions.json` 排除 `</>`、`<color=...>` 这类伪标签
- Codex 假阳性筛查：可选开启；检查完成后按术语、期望译法、问题类型和原文/译文文本组成的问题 cluster 调用本机 `codex exec`，在 `问题列` 追加 `fp_*` 辅助列
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
- 复数处理：双边整条 `术语+s` / `译法+s` 直接放行；其他复数形态疑似变体不进入问题报告
- Codex 假阳性筛查：可选开启；检查完成后按术语、期望译法、问题类型和原文/译文文本组成的问题 cluster 调用本机 `codex exec`，在 `术语命中问题` 追加 `fp_*` 辅助列
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
- 用途：逐行检查双语 Excel 中 `source` / `target` 的 `<...>`、`[color=...]` / `[/color]`、`{...}`、`\n` 和数字 tag 是否一致
- 检查类型：支持 `<...>` tag、`[color=...]` / `[/color]` tag、`{...}` placeholder、`\n` mark 与 `{n}` / `{n>` / `<n}` 数字 tag，可单独或组合检查
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

### 5. Target 中文检查

- 目录：`tools/chinese_target_checker`
- 用途：检查 Excel `target` 列是否包含中文字符或中文/全角标点
- 输出方式：默认直接修改原文件，在 `target` 右侧新增或复用 `中文检查` 标记列，命中行标记为 `含中文`
- 不额外生成问题工作表；如果工作簿里已有旧的 `中文检查问题` 工作表，运行时会移除
- CLI：

```bash
python3 tools/chinese_target_checker/check_chinese_target.py input.xlsx \
  -s Sheet1 \
  -t B \
  --start-row 2
```

- GUI：

```bash
python3 tools/chinese_target_checker/check_chinese_target_gui.py
```

详情见 `tools/chinese_target_checker/README.md`。

### 6. 法语 NBSP 恢复

- 目录：`tools/french_nbsp_restorer`
- 用途：恢复 Excel target 列中的法语 non-breaking space（NBSP）
- 恢复规则：`;`、`:`、`?`、`!` 前，以及 `«` / `»` 内侧
- 保护规则：不会改写 URL 内标点和 `12:30` 这类时间冒号
- 输出方式：生成新的结果 Excel，不覆盖原文件
- 写入方式：默认直接修复 target 列；也可以指定结果列，写入修复后的完整 target，未改动行也会复制过去
- CLI：

```bash
python3 tools/french_nbsp_restorer/restore_french_nbsp.py input.xlsx \
  -s Sheet1 \
  -t B \
  -r C \
  --start-row 2 \
  -o output_french_nbsp_restored.xlsx
```

- GUI：

```bash
python3 tools/french_nbsp_restorer/restore_french_nbsp_gui.py
```

详情见 `tools/french_nbsp_restorer/README.md`。

### 7. LLM 术语提取

- 目录：`tools/llm_term_extractor`
- 用途：从 Excel 的 `source` 列抽取游戏术语；同行 `target` 有内容则记录已有译法并做冲突复核，空 target 行按 source-only 术语收集处理
- 默认模型：`gpt-5.3-codex-spark`，`reasoning effort` 默认 `high`
- prompt 管理：默认使用 `tools/llm_term_extractor/prompts/` 下的 `extract_terms_zh_target.md` 与 `conflict_review_zh_target.md`，支持 CLI 覆盖对应文件
- 输出方式：生成新的结果 Excel（默认文件名 `<原文件名>_llm_terms.xlsx`），不会覆盖原文件；结果只包含 `本批次术语汇总表` 和 `冲突汇总`
- CLI：

```bash
python3 tools/llm_term_extractor/extract_llm_terms.py input.xlsx \
  -s Sheet1 \
  -c A \
  -t B \
  --start-row 2 \
  --batch-size 50 \
  --codex-model gpt-5.3-codex-spark \
  --codex-reasoning-effort high \
  --history-tb history_tb.xlsx \
  --extract-prompt-file ./tools/llm_term_extractor/prompts/extract_terms_zh_target.md \
  --conflict-prompt-file ./tools/llm_term_extractor/prompts/conflict_review_zh_target.md \
  -o output_llm_terms.xlsx
```

- GUI：

```bash
python3 tools/llm_term_extractor/extract_llm_terms_gui.py
```

也可在统一入口启动：

```bash
python3 toolshub_gui.py
```

详情见 `tools/llm_term_extractor/README.md`。

### 8. Xbench QA Report 转换

- 目录：`tools/xbench_report_transformer`
- 用途：把 Xbench 导出的 QA Report 转换成 `文件名, key, source, target, QA问题` 五列表格
- 聚类：优先按 Metadata 第一行的 key 聚类；没有 key 时按文件名+source 或 source 降级聚类
- 输出方式：生成新的结果 Excel，默认文件名 `xbench_transform_<原文件名>`
- CLI：

```bash
python3 tools/xbench_report_transformer/transform_xbench_report.py Xbench_QA_Report.xlsx \
  -s "Xbench QA" \
  -o xbench_flat.xlsx
```

- GUI：

```bash
python3 tools/xbench_report_transformer/transform_xbench_report_gui.py
```

详情见 `tools/xbench_report_transformer/README.md`。

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
- Target 中文检查详细说明：`tools/chinese_target_checker/README.md`
- 法语 NBSP 恢复详细说明：`tools/french_nbsp_restorer/README.md`
- LLM 术语提取详细说明：`tools/llm_term_extractor/README.md`
- Xbench QA Report 转换详细说明：`tools/xbench_report_transformer/README.md`

## 统一 GUI 入口

```bash
python3 toolshub_gui.py
```

会打开一个统一窗口，使用标签页管理这些工具；Workflow 编排页也支持给术语对检查选择历史 TB；原有各自的 GUI 入口仍然保留。

统一入口同样包含 `LLM术语提取` 和 `Xbench QA 转换` 页面；LLM 页面支持 source-only 与 target 两种模式和历史 TB 参数，GUI 内可选择或覆盖抽取 / 冲突复核的 prompt 文件。
