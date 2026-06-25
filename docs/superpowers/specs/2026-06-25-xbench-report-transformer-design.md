# Xbench Report Transformer Design

## Inline Summary

Add a focused Toolshub utility that transforms an ApSIC Xbench QA report into a flat, row-oriented Excel workbook. The output is meant for editing and filtering: each output row represents one source entry, grouped by key when available, with all QA issues for that entry merged into a single `QA问题` cell.

## Goals

- Convert Xbench QA reports into the existing Toolshub-style columns: `文件名`, `key`, `source`, `target`, `QA问题`.
- Preserve the original Xbench report file and write a new workbook by default.
- Group multiple QA issues for the same entry into one output row.
- Prefer metadata key-based grouping, with robust fallbacks when Xbench metadata is incomplete.
- Format QA issue text as `源术语 -> 目标术语：问题类型`.
- Join multiple QA issues with the Chinese semicolon `；`.
- Follow this repo's current Python and openpyxl patterns, with a CLI-first tool and tests.

## Non-Goals

- Do not perform new QA checks.
- Do not change existing term pair, glossary, tag, or workflow tools.
- Do not write changes back to the original source localization workbook.
- Do not require rich formatting from the Xbench source workbook.
- Do not infer corrected translations or rewrite target text.

## Input Shape

The initial supported input is an `.xlsx` report exported by Xbench. The report contains:

- A worksheet such as `Xbench QA`.
- A header row with `Source`, `Target`, `Comments`, and `Metadata`.
- QA group rows such as `Key Term Mismatch (提示 / Avis)`.
- Detail rows that contain source text, target text, comments, and metadata.

The parser should locate the header row by normalized header text instead of assuming a fixed row number. Xbench may set an incorrect worksheet dimension in the file, so the implementation must load the workbook in normal mode rather than relying on read-only dimensions.

## Metadata Rules

Metadata is treated as newline-separated text.

- Two or more non-empty metadata lines: first line is `key`, second line is `文件名`; extra lines are ignored for this version.
- One line that looks like a file name: `key` is empty, `文件名` is that line.
- One line that does not look like a file name: `key` is that line, `文件名` is empty.
- Empty metadata: both `key` and `文件名` are empty.

A metadata line looks like a file name when it ends with a common file extension such as `.xlsx`, `.xls`, `.xlsm`, `.csv`, or `.txt`, case-insensitively.

## Grouping Rules

Each detail row gets a grouping key:

- If `key` exists, group by `key`.
- If `key` is missing but `文件名` exists, group by `文件名 + source`.
- If both `key` and `文件名` are missing, group by `source`.

For grouped rows:

- `文件名`, `key`, `source`, and `target` use the first non-empty value seen in the group.
- `QA问题` contains unique issue texts in encounter order.
- Multiple issue texts are joined with `；`.

## QA Issue Formatting

The current QA group row should be attached to subsequent detail rows until the next QA group row appears.

For a group title like:

```text
Key Term Mismatch (提示 / Avis)
```

The issue text is:

```text
提示 -> Avis：Key Term Mismatch
```

If the parenthesized term pair cannot be parsed, use the title as-is. If only one side of the term pair can be parsed, preserve the available side and the issue type without inventing missing content.

## Output Workbook

The output workbook contains one sheet with the fixed header:

```text
文件名, key, source, target, QA问题
```

The default output path should use a Toolshub-style prefixed file name, for example:

```text
xbench_transform_Xbench_QA_Report.xlsx
```

CLI users can override the output path.

## CLI

Add a script under a new package:

```text
tools/xbench_report_transformer/transform_xbench_report.py
```

Expected options:

- positional `input_file`
- optional `-s/--sheet`
- optional `-o/--output`

If no sheet is supplied, use the active workbook sheet.

## Error Handling

- Missing input file raises a clear file-not-found error.
- Missing selected sheet raises a clear sheet error.
- Missing required headers raises a clear header error naming the expected headers.
- Detail rows before the first QA group row are ignored unless they have source/target/metadata content, in which case `QA问题` can be empty.
- Empty reports should still write a workbook with only the header row.

## Testing

Add focused tests for:

- Metadata with two lines, one key line, one file-name line, and empty metadata.
- Grouping by key.
- Grouping by `文件名 + source` when key is missing.
- Grouping by source when metadata is empty.
- QA issue formatting from `Issue Type (source / target)`.
- Multiple issues merged with `；` in encounter order.
- Header-row detection when the Xbench table does not start at row 1.
- Process-level output workbook generation without changing the input workbook.

Run focused and full verification:

```bash
python -m unittest tests.test_xbench_report_transformer
python -m unittest discover -s tests
python -m compileall -q tools tests
git diff --check
```

## Success Criteria

- The sample Xbench report can be transformed into the five-column flat workbook.
- Duplicate key entries are collapsed into one row.
- Missing metadata cases follow the documented fallback matrix.
- Existing tests continue to pass.
- The tool has a CLI entry point and a short README matching the repo's tool layout.

## Self-Review

- Placeholder scan: no TBD/TODO placeholders remain.
- Consistency check: metadata fallback rules match grouping rules and output columns.
- Scope check: this is a single CLI-first conversion tool, not a broader workflow rewrite.
- Ambiguity check: incomplete metadata behavior and multi-issue joining are explicit.
