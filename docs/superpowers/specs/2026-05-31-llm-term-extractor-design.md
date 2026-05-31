# LLM Term Extractor Design

## Inline Summary

Build a new Toolshub utility named LLM Term Extractor. It uses local `codex exec` with `gpt-5.3-codex-spark` to extract terminology from unmarked Excel `source` text. If `target` text is present, it extracts existing source-target term pairs and reviews real translation conflicts. If `target` is absent, it works as a source-only terminology collector.

## Goal

Create a reusable Excel tool that reproduces the "术语小子" workflow from the referenced Codex session:

- Decide terminology from `source`, not from `target`.
- Collect terms that need cross-context consistency: fixed names, activities, gameplay concepts, props, currencies, systems, titles, and fixed phrases.
- Exclude ordinary UI state words, ordinary action words, and complete sentences.
- Extract existing `target` terminology only when target text exists.
- Do not recommend or rewrite target translations.
- Deduplicate by source term.
- Flag real conflicts for manual review.

## User-Approved Behavior

The tool supports two row-level modes:

- Source-only mode: if no target column is selected, or a row's target cell is empty, the row contributes source terms only.
- Source-target mode: if target text is present, the row contributes source-target term observations and can participate in conflict review.

The two modes can coexist in one workbook. A file with mixed empty and non-empty target rows should produce one combined output workbook.

## Tool Location

Add a new package:

`tools/llm_term_extractor/`

This tool is separate from `tools/term_pair_checker`, because the existing term-pair checker is mark/tag driven, while this new tool uses Codex to identify unmarked terminology.

## CLI Inputs

The CLI accepts:

- `input_file`
- `-s/--sheet`
- `-c/--source-column`
- `-t/--target-column`
- `--start-row`, default `2`
- `--batch-size`, default `50`
- `--codex-model`, default `gpt-5.3-codex-spark`
- `--codex-reasoning-effort`, default `high`
- `--history-tb`
- `--history-sheet`
- `--history-source-column`
- `--history-target-column`
- `--history-start-row`, default `2`
- `--extract-prompt-file`
- `--conflict-prompt-file`
- `--dump-prompts-dir`
- `--keep-raw-codex-output`
- `-o/--output`

The default output path is `<input_stem>_llm_terms.xlsx`.

## Prompt Files

Prompt text is intentionally external and easy to edit:

- `tools/llm_term_extractor/prompts/extract_terms_zh_target.md`
- `tools/llm_term_extractor/prompts/conflict_review_zh_target.md`

The Python code loads these files and appends runtime payloads:

- mode: `source_only`, `source_target`, or `mixed`
- batch rows with row number, source text, and optional target text
- output JSON schema

The CLI can override each prompt path. When `--dump-prompts-dir` is set, every fully rendered prompt is written to disk for review. When `--keep-raw-codex-output` is set, raw Codex responses are retained in a sidecar JSONL file and optionally summarized in the workbook.

## Extraction Prompt Rules

The extraction prompt must emphasize:

- Whether a term should be collected is decided from `source`.
- Collect terms likely to need consistent translation across contexts.
- Include fixed names, event names, gameplay names, item names, currency names, system names, titles, fixed phrases, and context anchors like `招财小能手`.
- Exclude ordinary UI state text, ordinary operations, ordinary adjectives or verbs, generic labels that are unlikely to vary, and whole sentences.
- If target is present, extract the existing target expression from the same row.
- Do not propose a better translation.
- Do not add a target term when the target text does not contain one.
- Return strict JSON only.

## Extraction JSON Shape

Codex returns:

```json
{
  "rows": [
    {
      "row_index": 2,
      "terms": [
        {
          "source_term": "花艺",
          "target_term": "Art Floral",
          "term_type": "system_or_concept",
          "confidence": "high",
          "note": "固定系统/概念名，跨上下文需要一致"
        }
      ]
    }
  ]
}
```

For source-only rows, `target_term` is an empty string.

## Conflict Review

The tool first aggregates observed target terms by normalized source term. If a source term has more than one non-empty target observation, the conflict prompt reviews the group.

Conflict review must ignore:

- case changes
- obvious singular/plural changes
- ordinary grammatical form changes
- punctuation or mark-only differences

Conflict review must flag:

- substantially different target wording
- different official-looking term choices
- conceptually different translations
- project-term decisions that need manual confirmation

Example: `花艺` observed as `Art Floral` and `composition florale` should be flagged for review.

## Conflict JSON Shape

Codex returns:

```json
{
  "results": [
    {
      "source_term": "花艺",
      "decision": "conflict",
      "category": "实质译名差异",
      "confidence": "high",
      "note": "Art Floral 和 composition florale 是不同定稿风格，建议人工确认是否按语境区分"
    }
  ]
}
```

Allowed decisions:

- `same`
- `conflict`
- `review`

Only `conflict` and `review` groups go into `Conflicts_To_Review` and `Review_Before_Import`.

## Aggregation Rules

Normalize the source term for deduplication by trimming whitespace, collapsing internal whitespace, and casefolding. Preserve the first observed display form.

For each source term, keep:

- first source display
- observed target terms
- count
- row numbers
- source examples
- target examples
- term types
- confidence summary
- extraction notes
- history TB status
- conflict decision

Rows with no extracted terms are ignored.

## History TB

History TB is optional. It follows the existing Toolshub pattern:

- Prefer a sheet named `术语表`.
- Auto-detect `source` / `target`, `source术语` / `target术语`, or no-mark variants.
- If exactly two non-empty header columns exist, use them as source and target.
- Normalize source for matching.

History matching produces:

- `Already_In_History`: extracted source terms already present in history TB.
- `Import_Candidate`: clean new source-target pairs not already in history and not in conflict.
- `Review_Before_Import`: source-only terms, target-missing rows, conflict/review groups, and low-confidence extraction items.

## Output Workbook

The output workbook adds these sheets:

- `Terms_Source_Dedup`: one row per normalized source term.
- `Extraction_Evidence`: one row per extracted observation, with row number, source/target raw text, source term, target term, type, confidence, and note.
- `Conflicts_To_Review`: conflict/review evidence groups with row numbers and original text.
- `Import_Candidate`: clean new source-target pairs that can be imported after review.
- `Review_Before_Import`: rows that need a human decision before import.
- `Already_In_History`: extracted source terms found in history TB.
- `Summary`: scan counts, batch counts, model settings, prompt file paths, term counts, conflict counts, and history counts.

When `--keep-raw-codex-output` is set, raw batch responses are also written to `<output_stem>_codex_raw.jsonl`.

## GUI

Add `LlmTermExtractorApp` to `toolshub_gui.py` as a new tab named `LLM术语提取`.

The GUI supports:

- input and output file selection
- workbook sheet selection
- source column auto-detection
- optional target column auto-detection
- start row
- batch size
- model and reasoning effort fields
- prompt file selectors
- optional history TB fields
- checkboxes for dumping prompts and keeping raw Codex output

## Error Handling

If Codex exits non-zero, raise a clear runtime error that includes stderr or stdout.

If Codex returns invalid JSON, retry that batch once with the same prompt plus a strict JSON reminder. If it still fails, stop and report the batch number and output text path when available.

If a batch returns row indexes not present in the batch, ignore those rows and record a warning in `Summary`.

If no terms are extracted, still create the output workbook with headers and summary counts.

## Tests

Use dependency injection for Codex calls so tests do not require network, OpenAI credentials, or real `codex exec`.

Test coverage:

- source-only aggregation
- source-target aggregation
- mixed target and empty-target rows
- conflict group creation and review classification
- import/review/history sheet routing
- prompt file loading and rendering
- Codex command construction with `gpt-5.3-codex-spark`
- JSON parsing from plain and fenced responses
- invalid JSON retry behavior
- CLI argument parsing
- GUI metadata loading where current GUI tests can cover it without a display

## Non-Goals

- Do not rewrite translations.
- Do not generate recommended target terms.
- Do not replace the mark/tag-based term pair checker.
- Do not call the OpenAI API directly in this version; use local `codex exec`.
