# Chinese Target Checker Design

## Goal

Add a Toolshub utility that checks whether cells in the Excel `target` column contain Chinese characters.

## User-Approved Behavior

- The tool reads one workbook sheet and scans a selected `target` column from a configurable start row.
- By default, it writes a marker beside the original data.
- The default result column is the column immediately to the right of the selected `target` column.
- Rows whose target text contains Chinese characters get the marker `含中文`; rows without Chinese text are left blank.
- The user can specify a different result column.
- The user can optionally request a problem sheet listing rows with Chinese characters.
- The output is a new workbook and does not overwrite the input file.
- The tool is available through both CLI and GUI, and is added to the unified `toolshub_gui.py` notebook.

## Detection Rule

Use a Unicode CJK range regex that catches common Chinese ideographs:

`[\u3400-\u4DBF\u4E00-\u9FFF\uF900-\uFAFF]`

This intentionally focuses on Chinese characters, not Chinese punctuation.

## Outputs

The result column header is `中文检查`.

The optional problem sheet is named `中文检查问题` and contains:

- `行号`
- `target文本`
- `中文字符`

The summary object returned from the Python API includes output path, worksheet title, target/result columns, start row, processed row count, and matched row count.

## CLI

Create `tools/chinese_target_checker/check_chinese_target.py` with:

- positional `input_file`
- `-s/--sheet`
- `-t/--target-column`
- `-r/--result-column`
- `--start-row`
- `--problem-sheet`
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
- Checkbox for creating the problem sheet
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
- Default result column behavior
- Custom result column behavior
- Optional problem sheet output
- Preservation of the original input workbook

Use `python3 -m unittest tests.test_chinese_target_checker` for targeted verification and include the related existing suites before completion.
