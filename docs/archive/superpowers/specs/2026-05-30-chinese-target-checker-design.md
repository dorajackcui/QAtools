# Chinese Target Checker Design

## Goal

Add a Toolshub utility that checks whether cells in the Excel `target` column contain Chinese characters or Chinese/fullwidth punctuation.

## User-Approved Behavior

- The tool reads one workbook sheet and scans a selected `target` column from a configurable start row.
- By default, it inserts a marker column beside the original data.
- The default result column is a newly inserted column immediately to the right of the selected `target` column; existing columns to the right shift over.
- If the column immediately to the right of `target` is already `中文检查`, the tool reuses it instead of inserting another copy.
- Rows whose target text contains Chinese characters or Chinese/fullwidth punctuation get the marker `含中文`; other rows are left blank.
- The user can specify a different result column.
- The tool does not create a separate problem sheet.
- If an old `中文检查问题` sheet exists from a previous version, the tool removes it while processing.
- By default, the tool saves changes directly into the input workbook.
- If `-o/--output` is provided, the tool saves to that path instead.
- The tool is available through both CLI and GUI, and is added to the unified `toolshub_gui.py` notebook.

## Detection Rule

Use a Unicode regex that catches common Chinese ideographs, CJK punctuation, fullwidth punctuation, and common Chinese typography marks:

`[\u3400-\u4DBF\u4E00-\u9FFF\uF900-\uFAFF\u3000-\u303F\uFF01-\uFF0F\uFF1A-\uFF20\uFF3B-\uFF40\uFF5B-\uFF65\u00B7\u2014\u2018\u2019\u201C\u201D\u2026]`

This includes examples such as `【】（）`, `，。！？`, `《》“”‘’—…·`, while excluding plain ASCII punctuation and fullwidth alphanumerics by themselves.

## Outputs

The result column header is `中文检查`.

The summary object returned from the Python API includes saved path, worksheet title, target/result columns, start row, processed row count, and matched row count.

## CLI

Create `tools/chinese_target_checker/check_chinese_target.py` with:

- positional `input_file`
- `-s/--sheet`
- `-t/--target-column`
- `-r/--result-column`
- `--start-row`
- `-o/--output`

If required CLI arguments are missing in non-interactive mode, raise a clear `ValueError`, matching the style of existing tools.

## GUI

Create `tools/chinese_target_checker/check_chinese_target_gui.py`.

The GUI follows the existing lightweight Tkinter pattern:

- Input and output file selectors
- Sheet selector with automatic workbook sheet loading
- Target column auto-detection from the first row
- Optional result column
- Start row input
- Completion dialog with processed and matched counts

## Documentation

Update:

- `README.md`
- `docs/cli-usage.md`

Add:

- `tools/chinese_target_checker/README.md`

## Tests

Add `tests/test_chinese_target_checker.py` covering:

- Text detection helper
- Default inserted result column behavior
- Custom result column behavior
- Removal of stale `中文检查问题` sheet
- Preservation of the original input workbook

Use `python3 -m unittest tests.test_chinese_target_checker` for targeted verification and include the related existing suites before completion.
