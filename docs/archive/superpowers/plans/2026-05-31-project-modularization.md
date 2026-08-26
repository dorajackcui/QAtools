# Project Modularization Implementation Plan (Archived)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Extract shared Toolshub infrastructure for history TB handling, GUI helpers, and Codex execution while preserving current tool behavior.

**Architecture:** Add narrow shared modules under `tools/` and migrate existing consumers incrementally. Keep business rules and workbook sheet formats inside each tool package. Archive historical planning documents without deleting them.

**Tech Stack:** Python 3, `openpyxl`, `tkinter`, `unittest`, local Codex CLI.

---

## File Structure

- Create `tools/history_tb.py`: shared history TB sheet/column detection and row iteration primitives.
- Create `tools/gui_common.py`: shared Tkinter helper functions for integer parsing, file dialogs, sheet lists, and source/target auto-detection.
- Create `tools/codex_runner.py`: shared local Codex subprocess runner.
- Modify `tools/term_pair_checker/extract_terms_from_excel.py`: use `tools.history_tb` primitives.
- Modify `tools/llm_term_extractor/extract_llm_terms.py`: use `tools.history_tb` and `tools.codex_runner`.
- Modify `tools/false_positive_review.py`: use `tools.codex_runner`.
- Modify `tools/term_pair_checker/extract_terms_gui.py`: fix optional history TB start-row behavior and use GUI helpers where low risk.
- Modify `tools/workflow/workflow_gui.py`: fix optional history TB start-row behavior and use GUI helpers where low risk.
- Modify `tools/llm_term_extractor/extract_llm_terms_gui.py`: use GUI helpers where low risk.
- Modify `tests/test_term_pair_checker.py`, `tests/test_llm_term_extractor.py`, `tests/test_false_positive_review.py`, `tests/test_gui_excel_selection.py`; add `tests/test_history_tb.py`, `tests/test_codex_runner.py`, and optionally `tests/test_gui_common.py`.
- Move historical Superpowers artifacts from `docs/superpowers/{plans,specs}` into `docs/archive/superpowers/{plans,specs}` except the active modularization spec/plan.

## Task 1: Shared History TB Primitives

**Files:**
- Create: `tools/history_tb.py`
- Test: `tests/test_history_tb.py`
- Later consumers: term pair checker and LLM term extractor

- [ ] **Step 1: Write failing tests**

Create `tests/test_history_tb.py` with tests for:

```python
def test_detects_nomark_columns_before_marked_columns():
    # Workbook has 术语表 with source术语/target术语 and source术语（无mark）/target术语（无mark）.
    # detect_history_tb_columns(path) returns sheet_title="术语表", source_column="C", target_column="D".

def test_detects_header_row_from_start_row():
    # Workbook has metadata row 1, headers on row 2, data starts row 3.
    # detect_history_tb_columns(path, start_row=3) uses row 2 and returns C/D.

def test_rejects_same_column_when_source_and_target_arguments_match():
    # detect_history_tb_columns(path, source_column="A", target_column="A") raises ValueError.

def test_iter_history_rows_skips_blank_source_or_target():
    # iter_history_rows(...) returns only rows where both source and target text are non-empty.
```

- [ ] **Step 2: Verify red**

Run:

```bash
python3 -m unittest tests.test_history_tb
```

Expected: failure with `ModuleNotFoundError` or missing functions.

- [ ] **Step 3: Implement `tools/history_tb.py`**

Expose:

```python
@dataclass(frozen=True)
class HistoryTbColumns:
    sheet_title: str
    source_column: str | None
    target_column: str | None

@dataclass(frozen=True)
class HistoryTbRow:
    row_index: int
    source_text: str
    target_text: str

def normalize_history_header(value: object) -> str: ...
def choose_history_worksheet(workbook, sheet: str | None, preferred_sheet: str = "术语表"): ...
def detect_history_columns_from_header_row(worksheet, header_row: int, source_column: str | None = None, target_column: str | None = None) -> tuple[str, str]: ...
def detect_history_tb_columns(history_tb_file, *, sheet=None, source_column=None, target_column=None, start_row=2, preferred_sheet="术语表") -> HistoryTbColumns: ...
def iter_history_rows(history_tb_file, *, sheet=None, source_column=None, target_column=None, start_row=2, preferred_sheet="术语表") -> tuple[str, str, str, tuple[HistoryTbRow, ...]]: ...
```

Header priority must be exact `source/target`, then no-mark variants, then marked variants, then two-column fallback. The function must reject same source/target columns.

- [ ] **Step 4: Verify green**

Run:

```bash
python3 -m unittest tests.test_history_tb
python3 -m unittest discover -s tests
```

Expected: all tests pass.

- [ ] **Step 5: Commit**

```bash
git add tools/history_tb.py tests/test_history_tb.py
git commit -m "refactor: add shared history TB helpers"
```

## Task 2: Migrate History TB Consumers And Fix Optional GUI History Bug

**Files:**
- Modify: `tools/term_pair_checker/extract_terms_from_excel.py`
- Modify: `tools/llm_term_extractor/extract_llm_terms.py`
- Modify: `tools/term_pair_checker/extract_terms_gui.py`
- Modify: `tools/workflow/workflow_gui.py`
- Modify tests: `tests/test_term_pair_checker.py`, `tests/test_llm_term_extractor.py`, `tests/test_gui_excel_selection.py`

- [ ] **Step 1: Write failing GUI regression tests**

In `tests/test_gui_excel_selection.py`, add tests:

```python
def test_term_pair_ignores_invalid_history_start_without_history_file(self):
    # history_tb_file_var is "", history_start_row_var is "not-an-int".
    # run_extraction() must call process_excel once and not show an error.

def test_workflow_ignores_invalid_history_start_without_history_file(self):
    # term_history_tb_file_var is "", term_history_start_row_var is "not-an-int".
    # run_selected_tasks() must call run_workflow once and not show an error.
```

- [ ] **Step 2: Verify red**

Run:

```bash
python3 -m unittest tests.test_gui_excel_selection
```

Expected: the new tests fail because the GUIs still parse history start rows when no history file is selected.

- [ ] **Step 3: Migrate processor history detection**

Update term pair checker:

- import `HistoryTbColumns`, `detect_history_tb_columns`, `iter_history_rows` from `tools.history_tb`
- keep `RecordedTermPair`, mark stripping, and `normalize_history_source_key` local
- implement `load_history_tb_mapping()` by iterating `HistoryTbRow`
- keep public return shape of `detect_history_tb_columns()` compatible for GUI users if needed

Update LLM term extractor:

- import shared `detect_history_tb_columns` as the public helper
- implement `load_history_tb_mapping()` by iterating `HistoryTbRow`
- keep LLM source-key normalization local

- [ ] **Step 4: Fix optional GUI history parsing**

In `tools/term_pair_checker/extract_terms_gui.py` and `tools/workflow/workflow_gui.py`:

- initialize history start row to `2`
- only parse history start row if a history file is selected
- if no history file is selected, pass history sheet/source/target as `None`

- [ ] **Step 5: Verify**

Run:

```bash
python3 -m unittest tests.test_history_tb tests.test_term_pair_checker tests.test_llm_term_extractor tests.test_gui_excel_selection tests.test_workflow_runner
python3 -m unittest discover -s tests
```

Expected: all tests pass.

- [ ] **Step 6: Commit**

```bash
git add tools/history_tb.py tools/term_pair_checker/extract_terms_from_excel.py tools/llm_term_extractor/extract_llm_terms.py tools/term_pair_checker/extract_terms_gui.py tools/workflow/workflow_gui.py tests/test_term_pair_checker.py tests/test_llm_term_extractor.py tests/test_gui_excel_selection.py tests/test_history_tb.py
git commit -m "refactor: share history TB handling"
```

## Task 3: Shared Codex Runner

**Files:**
- Create: `tools/codex_runner.py`
- Modify: `tools/false_positive_review.py`
- Modify: `tools/llm_term_extractor/codex_term_review.py`
- Test: `tests/test_codex_runner.py`, `tests/test_false_positive_review.py`, `tests/test_llm_term_extractor.py`

- [ ] **Step 1: Write failing runner tests**

Create `tests/test_codex_runner.py`:

```python
def test_build_codex_exec_command_sets_model_reasoning_and_output():
    command = build_codex_exec_command(output_path=Path("/tmp/out.txt"), model="gpt-5.3-codex-spark", reasoning_effort="high")
    assert "--model" in command
    assert "gpt-5.3-codex-spark" in command
    assert "--output-last-message" in command

def test_run_codex_exec_prompt_reads_output_last_message():
    # Patch subprocess.run, write output_path, assert prompt goes to stdin and returned text is file content.
```

- [ ] **Step 2: Verify red**

Run:

```bash
python3 -m unittest tests.test_codex_runner
```

Expected: fails because `tools.codex_runner` does not exist.

- [ ] **Step 3: Implement shared runner**

Create `tools/codex_runner.py`:

```python
def build_codex_exec_command(*, output_path, model=None, reasoning_effort="high", sandbox="read-only") -> list[str]: ...
def run_codex_exec_prompt(prompt, *, output_path, model=None, reasoning_effort="high", timeout_seconds=600, sandbox="read-only") -> str: ...
```

It must preserve current command flags: `--ask-for-approval never`, `exec`, `--ephemeral`, `--ignore-rules`, `--skip-git-repo-check`, `--sandbox read-only`, `--output-last-message`.

- [ ] **Step 4: Migrate consumers**

- `tools/false_positive_review.py` uses `run_codex_exec_prompt()` in `review_cluster_batch_with_codex`.
- `tools/llm_term_extractor/codex_term_review.py` uses the shared runner while keeping `build_codex_command()` as a backward-compatible wrapper for tests and callers.

- [ ] **Step 5: Verify**

Run:

```bash
python3 -m unittest tests.test_codex_runner tests.test_false_positive_review tests.test_llm_term_extractor
python3 -m unittest discover -s tests
```

Expected: all tests pass.

- [ ] **Step 6: Commit**

```bash
git add tools/codex_runner.py tools/false_positive_review.py tools/llm_term_extractor/codex_term_review.py tests/test_codex_runner.py tests/test_false_positive_review.py tests/test_llm_term_extractor.py
git commit -m "refactor: share Codex subprocess runner"
```

## Task 4: Small GUI Common Helpers

**Files:**
- Create: `tools/gui_common.py`
- Modify: `tools/term_pair_checker/extract_terms_gui.py`
- Modify: `tools/workflow/workflow_gui.py`
- Modify: `tools/llm_term_extractor/extract_llm_terms_gui.py`
- Test: `tests/test_gui_common.py`, `tests/test_gui_excel_selection.py`

- [ ] **Step 1: Write failing tests**

Create `tests/test_gui_common.py`:

```python
def test_parse_positive_int_accepts_blank_default():
    assert parse_positive_int("", default=2, field_name="开始行") == 2

def test_parse_positive_int_rejects_non_integer():
    with self.assertRaises(ValueError):
        parse_positive_int("x", default=2, field_name="开始行")

def test_parse_positive_int_rejects_zero():
    with self.assertRaises(ValueError):
        parse_positive_int("0", default=2, field_name="开始行")
```

- [ ] **Step 2: Verify red**

Run:

```bash
python3 -m unittest tests.test_gui_common
```

Expected: fails because `tools.gui_common` does not exist.

- [ ] **Step 3: Implement helper module**

Create `tools/gui_common.py` with:

```python
def parse_positive_int(raw_value: str, *, default: int, field_name: str) -> int:
    ...

def set_combobox_values(combobox, values: tuple[str, ...], variable, default_value: str | None) -> str:
    ...
```

Keep it small. Do not introduce a base GUI class in this phase.

- [ ] **Step 4: Adopt helpers in touched GUIs**

Use `parse_positive_int()` for start row and history start row parsing in term pair, workflow, and LLM GUIs. Keep messagebox text equivalent.

- [ ] **Step 5: Verify**

Run:

```bash
python3 -m unittest tests.test_gui_common tests.test_gui_excel_selection
python3 -m unittest discover -s tests
```

Expected: all tests pass.

- [ ] **Step 6: Commit**

```bash
git add tools/gui_common.py tools/term_pair_checker/extract_terms_gui.py tools/workflow/workflow_gui.py tools/llm_term_extractor/extract_llm_terms_gui.py tests/test_gui_common.py tests/test_gui_excel_selection.py
git commit -m "refactor: add small GUI helpers"
```

## Task 5: Archive Historical Planning Documents

**Files:**
- Move: `docs/superpowers/plans/2026-05-30-chinese-target-checker.md`
- Move: `docs/superpowers/plans/2026-05-31-llm-term-extractor.md`
- Move: `docs/superpowers/specs/2026-05-30-chinese-target-checker-design.md`
- Move: `docs/superpowers/specs/2026-05-31-llm-term-extractor-design.md`
- Create/Modify: `docs/archive/README.md`
- Modify: `README.md` or `docs/cli-usage.md` only if a navigation note is useful.

- [ ] **Step 1: Move files with git mv**

Run:

```bash
mkdir -p docs/archive/superpowers/plans docs/archive/superpowers/specs
git mv docs/superpowers/plans/2026-05-30-chinese-target-checker.md docs/archive/superpowers/plans/
git mv docs/superpowers/plans/2026-05-31-llm-term-extractor.md docs/archive/superpowers/plans/
git mv docs/superpowers/specs/2026-05-30-chinese-target-checker-design.md docs/archive/superpowers/specs/
git mv docs/superpowers/specs/2026-05-31-llm-term-extractor-design.md docs/archive/superpowers/specs/
```

Keep the active modularization spec and plan in `docs/superpowers/`.

- [ ] **Step 2: Add archive README**

Create `docs/archive/README.md` explaining:

```markdown
# Archive

This directory keeps historical implementation plans and design notes that are useful for provenance but not part of the active user documentation.
```

- [ ] **Step 3: Verify docs paths and tests**

Run:

```bash
python3 -m unittest discover -s tests
git diff --check
```

Expected: all tests pass; no whitespace issues.

- [ ] **Step 4: Commit**

```bash
git add docs
git commit -m "docs: archive historical implementation plans"
```

## Task 6: Final Verification And Review

**Files:**
- No intended source changes unless review finds issues.

- [ ] **Step 1: Run final verification**

Run:

```bash
python3 -m unittest discover -s tests
python3 -m compileall -q tools tests
git diff --check main...HEAD
git status --short --branch
```

Expected:

- unit tests pass
- compileall has no output and exit code 0
- diff check has no output
- branch is clean

- [ ] **Step 2: Request final code review**

Ask a reviewer to compare the branch against `main` and check:

- no behavior changes to public CLI/workbook formats
- shared modules are small and useful
- old tests still cover migrated behavior
- optional history GUI bug is fixed
- archive move is appropriate

- [ ] **Step 3: Fix review issues if any**

For each issue:

1. verify it against the code
2. write/update a failing test if behavior-related
3. implement fix
4. run focused test
5. commit

- [ ] **Step 4: Final status**

Run final verification commands again and report results.

## Self-Review

- Spec coverage: every spec goal maps to a task.
- Placeholder scan: no TBD/TODO/fill-in-later text remains.
- Type consistency: shared helper names are stable across tasks.
- Scope control: no large algorithm split is included in this phase.
