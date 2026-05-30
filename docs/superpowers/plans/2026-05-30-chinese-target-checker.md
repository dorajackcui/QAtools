# Chinese Target Checker Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build an Excel tool that marks rows whose `target` column contains Chinese characters, with an optional problem sheet.

**Architecture:** Add a focused `tools/chinese_target_checker` package with a pure detection helper, an Excel processor, and a Tkinter GUI. Follow existing Toolshub conventions for CLI arguments, workbook output, sheet and column handling, tests, README documentation, and unified GUI registration.

**Tech Stack:** Python 3, `openpyxl`, `tkinter`, `unittest`.

---

## File Structure

- Create `tools/chinese_target_checker/__init__.py`: package marker.
- Create `tools/chinese_target_checker/check_chinese_target.py`: CLI, text detection, workbook processing.
- Create `tools/chinese_target_checker/check_chinese_target_gui.py`: standalone Tkinter GUI frame.
- Create `tools/chinese_target_checker/README.md`: user documentation.
- Create `tests/test_chinese_target_checker.py`: focused unit and workbook tests.
- Modify `toolshub_gui.py`: import and add the new GUI tab.
- Modify `README.md`: list the new tool.
- Modify `docs/cli-usage.md`: document non-interactive usage and command template.

### Task 1: Core Detection And Workbook Processing

**Files:**
- Create: `tests/test_chinese_target_checker.py`
- Create: `tools/chinese_target_checker/__init__.py`
- Create: `tools/chinese_target_checker/check_chinese_target.py`

- [ ] **Step 1: Write the failing tests**

Create `tests/test_chinese_target_checker.py` with tests for `contains_chinese`, default adjacent result column, custom result column, optional problem sheet, and input preservation.

- [ ] **Step 2: Run tests to verify they fail**

Run: `python3 -m unittest tests.test_chinese_target_checker`

Expected: import failure because `tools.chinese_target_checker` does not exist yet.

- [ ] **Step 3: Add the minimal implementation**

Implement:

- `CHINESE_PATTERN`
- `contains_chinese(value: object) -> bool`
- `extract_chinese_characters(value: object) -> str`
- `CheckSummary`
- `normalize_column`
- `build_default_output_path`
- `process_excel`
- `parse_args`
- `prompt_if_missing`
- `main`

Use `openpyxl.load_workbook`, write `中文检查` to the result column header, write `含中文` for matched rows and `None` for non-matched rows, and create `中文检查问题` only when requested.

- [ ] **Step 4: Run tests to verify they pass**

Run: `python3 -m unittest tests.test_chinese_target_checker`

Expected: all tests pass.

### Task 2: GUI And Toolshub Integration

**Files:**
- Create: `tools/chinese_target_checker/check_chinese_target_gui.py`
- Modify: `toolshub_gui.py`

- [ ] **Step 1: Add GUI smoke tests if current test patterns support it**

Inspect `tests/test_gui_excel_selection.py` and reuse existing GUI auto-detection patterns if they can cover the new frame without requiring a display.

- [ ] **Step 2: Implement the GUI**

Create a `ChineseTargetCheckerApp(ttk.Frame)` with file selectors, sheet selector, target/result/start row fields, problem sheet checkbox, and a run button.

- [ ] **Step 3: Register the GUI in Toolshub**

Import `ChineseTargetCheckerApp` in `toolshub_gui.py`, instantiate it, and add a tab named `Target中文检查`.

- [ ] **Step 4: Run GUI-related tests**

Run: `python3 -m unittest tests.test_gui_excel_selection`

Expected: existing GUI selection tests pass.

### Task 3: Documentation And Verification

**Files:**
- Create: `tools/chinese_target_checker/README.md`
- Modify: `README.md`
- Modify: `docs/cli-usage.md`

- [ ] **Step 1: Add tool README**

Document CLI and GUI usage, default adjacent result column behavior, optional problem sheet, and output file behavior.

- [ ] **Step 2: Update root documentation**

Add the new tool to `README.md` and `docs/cli-usage.md`, including a non-interactive command template.

- [ ] **Step 3: Run focused verification**

Run: `python3 -m unittest tests.test_chinese_target_checker tests.test_gui_excel_selection`

Expected: all listed tests pass.

- [ ] **Step 4: Run related regression tests**

Run: `python3 -m unittest tests.test_tag_placeholder_checker tests.test_french_nbsp_restorer tests.test_workflow_runner tests.test_chinese_target_checker`

Expected: all listed tests pass.

- [ ] **Step 5: Verify with the user-provided workbook**

Run the CLI against `/Users/zhiyangcui/Documents/orchestra/done/task/task.xlsx` with explicit sheet, target column, start row, output path, and `--problem-sheet` after inspecting workbook metadata.

Expected: command exits successfully and prints processed and matched row counts.

## Self-Review

- Spec coverage: each requested behavior is covered by Task 1, Task 2, or Task 3.
- Placeholder scan: no deferred implementation placeholders remain.
- Type consistency: `CheckSummary`, `process_excel`, and GUI names are consistent across tasks.
