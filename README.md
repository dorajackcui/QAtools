# Toolshub

这个仓库现在是一个轻量的 Excel `toolshub`，每个工具都有独立目录、独立说明，同时继续共用根目录依赖。

如果你是通过 agent / 脚本来调用这些工具，建议优先查看独立的 CLI 文档：[CLI 使用指南](docs/cli-usage.md)。

## 当前工具

### 1. 术语检查

- 目录：`tools/term_pair_checker`
- 用途：从可选 mark 提取新术语对，并结合历史 TB 检查 `source` / `target` 是否对齐
- mark 支持：`【】`、普通 `[]`，且可在 GUI 中多选组合检查；`<...>`、`{...}` 统一按 tag / placeholder 处理
- 检查规则：`术语表` 保留 tag，实际术语检查会忽略 tag，并回溯整表未标注出现
- 复数处理：回扫时双边整条 `术语+s` / `译法+s` 直接放行；其他复数形态疑似变体不进入问题报告
- 历史 TB：可选选择历史 TB；选择后会用“历史 TB 全量 + 本批次新增 TB”一起检查；不选 mark 时仅使用历史 TB，等价于原“术语表命中检查”
- 方括号处理：`[color=...]` / `[/color]` 这类格式 tag 不进入术语表
- 输出增强：结果会给出合并后的 `术语表`（只包含本次检查文本中涉及的术语，并包含保留 mark、无 mark 和术语来源列），以及带原文上下文的 `问题列`
- GUI 增强：自动读取并识别工作表及双语列；历史 TB 详情默认折叠，页面可滚动且主按钮固定在底部
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
  -o term_pair_check_output.xlsx
```

仅使用历史 TB 检查、不从 mark 提取新术语：

```bash
python3 tools/term_pair_checker/extract_terms_from_excel.py input.xlsx \
  -s Sheet1 -c A -t B \
  --no-term-mark \
  --history-tb glossary.xlsx \
  --history-sheet Glossary
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

### 2. Excel 分行拆列

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

### 3. Tag / Placeholder 检查

- 目录：`tools/tag_placeholder_checker`
- 用途：逐行检查双语 Excel 中 `source` / `target` 的 `<...>`、`[color=...]` / `[/color]`、`{...}`、`\n` 和 memoQ tag 是否一致
- 检查类型：支持 `<...>` tag、`[color=...]` / `[/color]` tag、`{...}` placeholder、`\n` mark 与 `{n}` / `{n>` / `<n}` memoQ tag；GUI 中 memoQ 与其他类型互斥，CLI 可按需显式组合
- `<...>` 识别：默认所有 `<...>` 都按普通 tag 检查；memoQ 数字 protected marker 独立为 `memoq` 类型
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
  --token-type memoq
```

- GUI：

```bash
python3 tools/tag_placeholder_checker/check_tags_and_placeholders_gui.py
```

详情见 `tools/tag_placeholder_checker/README.md`。

### 4. 换行数量检查

- 目录：`tools/line_break_checker`
- 用途：逐行比较 `source` / `target` 单元格中的真实换行数量
- 兼容换行：`LF`、`CRLF`、`CR`；`CRLF` 按一个换行计算
- 输出方式：生成新的结果 Excel，并新增 `换行数量问题` 工作表
- CLI：

```bash
python3 tools/line_break_checker/check_line_breaks.py input.xlsx \
  -s Sheet1 -c A -t B --start-row 2
```

- GUI：

```bash
python3 tools/line_break_checker/check_line_breaks_gui.py
```

详情见 `tools/line_break_checker/README.md`。

### 5. 同源译文一致性

- 目录：`tools/source_consistency_checker`
- 用途：检查完全相同的 `source` 是否对应多个不同 `target`
- 匹配规则：source 和 target 都按单元格文本精确比较；空 source 跳过，空 target 参与比较
- 输出方式：生成新的结果 Excel，并新增 `同源译文不一致` 工作表
- CLI：

```bash
python3 tools/source_consistency_checker/check_source_consistency.py input.xlsx \
  -s Sheet1 -c A -t B --start-row 2
```

- GUI：

```bash
python3 tools/source_consistency_checker/check_source_consistency_gui.py
```

详情见 `tools/source_consistency_checker/README.md`。

### 6. Target 中文检查

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

### 7. 法语 NBSP 恢复

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
- 术语检查详细说明：`tools/term_pair_checker/README.md`
- Tag / Placeholder 检查详细说明：`tools/tag_placeholder_checker/README.md`
- 换行数量检查详细说明：`tools/line_break_checker/README.md`
- 同源译文一致性检查详细说明：`tools/source_consistency_checker/README.md`
- Excel 分行拆列详细说明：`tools/excel_line_splitter/README.md`
- Target 中文检查详细说明：`tools/chinese_target_checker/README.md`
- 法语 NBSP 恢复详细说明：`tools/french_nbsp_restorer/README.md`
- Xbench QA Report 转换详细说明：`tools/xbench_report_transformer/README.md`

## 统一 GUI 入口

```bash
python3 toolshub_gui.py
```

会打开一个统一窗口管理这些工具。“一键质量检查”页默认执行“质量检查”板块的全部项目：术语检查、Tag 检查、换行数量检查、同源译文一致性和 Target 中文检查，并统一写入一个结果 Excel；每项检查仍可单独取消。workflow 输出中的术语问题表命名为 `术语问题`，不保留 Tag 检查原有的详细 `检查汇总`，而是新增仅含“检查项”和“问题行数”的 `质量检查汇总`。术语和 Tag 的详细设置默认折叠，Tag 可明确切换“常规 Tag / memoQ Tag”模式。该页面也支持给术语检查选择历史 TB，术语检查未选择 mark 时会仅使用历史 TB 做全表命中检查。
