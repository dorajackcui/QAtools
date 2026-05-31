# LLM 术语提取

用于从 Excel 提取术语并形成可导入的术语表候选。可直接从 `source` 与 `target` 列联合抽取，也可在 `target` 留空时只收集 source 术语。工具会对同源术语多译法做冲突复核，并按工作表输出分类结果。

## 两种模式

- **target 模式**：提供 `-t/--target-column` 时，按每行 `source` + `target` 组合抽取术语。模型在 `source` 中识别术语候选，同时保留该行出现的 `target` 表达。
- **source-only 模式**：`target` 列不传（或留空）时，仅依据 `source` 抽取术语。常用于已有未标记术语、无 target 的场景快速建池。

## Prompt 文件

Prompt 目录：`tools/llm_term_extractor/prompts/`

- 默认抽取 prompt：`extract_terms_zh_target.md`
- 默认冲突复核 prompt：`conflict_review_zh_target.md`

CLI 可覆盖：

- `--extract-prompt-file`：替换抽取 prompt
- `--conflict-prompt-file`：替换冲突复核 prompt

默认模型：`gpt-5.3-codex-spark`，可用 `--codex-model` 覆盖；默认 `reasoning effort` 为 `high`，可用 `--codex-reasoning-effort` 覆盖。

## 运行方式（CLI）

```bash
python3 tools/llm_term_extractor/extract_llm_terms.py ./input.xlsx \
  -s Sheet1 \
  -c A \
  -t B \
  --start-row 2 \
  --batch-size 50 \
  -o ./artifacts/input_llm_terms.xlsx
```

```bash
python3 tools/llm_term_extractor/extract_llm_terms.py ./source_only.xlsx \
  -s Sheet1 \
  -c A \
  --start-row 2 \
  --batch-size 50 \
  -o ./artifacts/source_only_llm_terms.xlsx
```

### 历史 TB 参数

- `--history-tb`：历史术语表文件路径。存在则加载后与本次抽取结果对比。
- `--history-sheet`：历史 TB 工作表名，不传默认优先用 `术语表`，否则用当前活动表。
- `--history-source-column` / `--history-target-column`：历史 TB 的列可传字母或表头名（如 `A` / `B` / `source` / `source术语` 等）。
- `--history-start-row`：历史 TB 起始数据行，默认 `2`（第一行为表头）。

### 示例：带历史 TB 与 prompt 覆盖

```bash
python3 tools/llm_term_extractor/extract_llm_terms.py ./input.xlsx \
  -s Sheet1 \
  -c A -t B \
  --start-row 2 \
  --batch-size 50 \
  --history-tb ./history_tb.xlsx \
  --history-sheet Glossary \
  --history-source-column source \
  --history-target-column target \
  --extract-prompt-file ./prompts/custom_extract.md \
  --conflict-prompt-file ./prompts/custom_conflict.md \
  --codex-model gpt-5.3-codex-spark \
  --dump-prompts-dir ./artifacts/prompts \
  --keep-raw-codex-output \
  -o ./artifacts/input_llm_terms.xlsx
```

该工具会调用本机 `codex exec`。在单元测试里不会要求真实调用（测试通过 mock）。

## 输出工作表

- `Terms_Source_Dedup`：术语聚合结果（去重）  
  列：`source_term`, `target_terms_observed`, `row_count`, `rows`, `source_examples`, `target_examples`, `term_types`, `confidences`, `notes`, `decision`, `decision_reason`
- `Extraction_Evidence`：逐条抽取证据，含行号与源/目标文本
- `Conflicts_To_Review`：多译法复核输出，含 LLM 决策信息
- `Import_Candidate`：可直接导入候选（`decision` 为 `same` 或无冲突且能确定唯一 target）
- `Review_Before_Import`：待复核项（含 target 缺失、多译法或冲突）
- `Already_In_History`：已在历史 TB 中命中的术语
- `Summary`：统计信息与关键计数（包括 conflict/import/review/历史命中数）

## GUI

- 独立入口：

```bash
python3 tools/llm_term_extractor/extract_llm_terms_gui.py
```

- 统一入口：

```bash
python3 toolshub_gui.py
```

从统一入口选择 `LLM术语提取` 标签页即可。
