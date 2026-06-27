# LLM 术语提取

用于从 Excel 提取术语并形成可导入的术语表候选。可从 `source` 识别术语，并在同一行 `target` 有内容时记录已有译法；也可在 `target` 未指定或为空时只收集 source 术语。工具会对同源术语多译法做冲突复核，最终结果只输出本批次术语汇总和冲突汇总两张工作表。

## 两种模式

- **target / mixed 模式**：提供 `-t/--target-column` 时，模型仍由 `source` 识别术语候选；同行 `target` 有内容则记录已有译法，空 target 行按 source-only 处理。
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
  --extract-prompt-file ./tools/llm_term_extractor/prompts/extract_terms_zh_target.md \
  --conflict-prompt-file ./tools/llm_term_extractor/prompts/conflict_review_zh_target.md \
  --codex-model gpt-5.3-codex-spark \
  --dump-prompts-dir ./artifacts/prompts \
  --keep-raw-codex-output \
  -o ./artifacts/input_llm_terms.xlsx
```

该工具会调用本机 `codex exec`。在单元测试里不会要求真实调用（测试通过 mock）。

## 输出工作表

输出 workbook 只包含：

- `本批次术语汇总表`：一行一个去重后的 source 术语，合并历史术语（如传入历史 TB）与本批次新增术语。包含 `target术语`、本批次 target 观察值、术语来源、原始行号、实例原文、实例译文、类型、备注和多译法判断信息。
- `冲突汇总`：只列出需要人工确认的多译法冲突 / 复核项，包含 observed target、行号、建议统一 target、复核原因和实例原文/译文。

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
