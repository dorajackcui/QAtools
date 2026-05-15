# Entity TUI Menu Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a terminal menu for the four-step entity workflow, reachable from the top-level interactive menu and direct `phraseloom entity` aliases.

**Architecture:** Keep the TUI in `phraseloom.interactive`, where the existing template workflow menu already lives. Reuse the existing entity workflow functions and default output path helpers so interactive and direct CLI behavior stay consistent.

**Tech Stack:** Python 3.11, `argparse`, `openpyxl`, `unittest`, existing PhraseLoom workflow modules.

---

## File Structure

- Modify `phraseloom/interactive.py`: add an entity submenu, entity step handlers, entity stat display helpers, and exports.
- Modify `phraseloom/cli.py`: route `phraseloom entity` and `phraseloom entity-interactive` to the new entity submenu and include the direct interactive aliases in top-level help.
- Modify `tests/test_template_workflow.py`: update the existing menu assertion to include the entity option.
- Modify `tests/test_entity_workflow.py`: add interactive entity menu and workbook behavior tests beside the existing entity CLI tests.
- Modify `README.md`: document the new TUI entry points briefly in the entity workflow section.

## Task 1: Top-Level Entity Menu Entry

**Files:**
- Modify: `tests/test_template_workflow.py`
- Modify: `phraseloom/interactive.py`

- [ ] **Step 1: Write the failing top-level menu test**

In `tests/test_template_workflow.py`, update `test_interactive_menu_shows_three_step_workflow` so it expects the entity menu option:

```python
self.assertIn("4) Entity workflow", menu)
```

Keep the existing assertions for steps 1, 2, and 3.

- [ ] **Step 2: Run the focused test and verify RED**

Run:

```powershell
python -m unittest tests.test_template_workflow.TemplateDemoTests.test_interactive_menu_shows_three_step_workflow
```

Expected: FAIL because `4) Entity workflow` is not printed yet.

- [ ] **Step 3: Implement the top-level menu option**

In `phraseloom/interactive.py`, update `run_interactive()` to print the new option and dispatch it:

```python
def run_interactive() -> int:
    print("Localization Workflow")
    print()
    print("1) Build TM from completed Excel")
    print("2) Prepare translator file for new source")
    print("3) Fill source from translated file")
    print("4) Entity workflow")
    print("q) Quit")

    action = _prompt_text("Choose step", default="2").lower()
    if action in {"q", "quit", "exit"}:
        print("Bye.")
        return 0
    if action in {"1", "tm", "tm-extract", "extract-tm", "build"}:
        return _interactive_tm_extract()
    if action in {"2", "extract", "prepare", "p"}:
        return _interactive_extract()
    if action in {"3", "fill", "f"}:
        return _interactive_fill()
    if action in {"4", "entity", "entity-workflow", "e"}:
        return run_entity_interactive(back_returns_to_main=True)

    print(f"Unknown step: {action}")
    return 2
```

Add a temporary minimal `run_entity_interactive()` below `_interactive_fill()` so Task 1 can pass before the full submenu exists:

```python
def run_entity_interactive(*, back_returns_to_main: bool = False) -> int:
    print("Entity Workflow")
    print()
    print("1) Build entity memory from TM reusable units")
    print("2) Prepare source entity pack")
    print("3) Fill completed entity pack")
    print("4) Merge filled entity pack back to translator todo")
    print("b) Back")
    print("q) Quit")

    action = _prompt_text("Choose entity step", default="2").lower()
    if action in {"q", "quit", "exit", "b", "back"}:
        print("Bye.")
        return 0
    print(f"Unknown entity step: {action}")
    return 2
```

- [ ] **Step 4: Run the focused test and verify GREEN**

Run:

```powershell
python -m unittest tests.test_template_workflow.TemplateDemoTests.test_interactive_menu_shows_three_step_workflow
```

Expected: PASS.

## Task 2: Direct Interactive Entity Aliases

**Files:**
- Modify: `tests/test_entity_workflow.py`
- Modify: `phraseloom/cli.py`
- Modify: `phraseloom/interactive.py`

- [ ] **Step 1: Write failing alias tests**

In `tests/test_entity_workflow.py`, add a new test method to `EntityPackWorkflowCliTests`:

```python
def test_entity_interactive_aliases_open_entity_menu(self):
    from contextlib import redirect_stdout
    from io import StringIO
    from unittest.mock import patch

    from phraseloom.cli import _dispatch

    for command in ("entity", "entity-interactive"):
        with self.subTest(command=command):
            stream = StringIO()
            with patch("builtins.input", side_effect=["q"]), redirect_stdout(stream):
                self.assertEqual(_dispatch([command]), 0)

            menu = stream.getvalue()
            self.assertIn("Entity Workflow", menu)
            self.assertIn("1) Build entity memory from TM reusable units", menu)
            self.assertIn("4) Merge filled entity pack back to translator todo", menu)
```

Add a top-level routing test:

```python
def test_top_level_interactive_option_opens_entity_menu(self):
    from contextlib import redirect_stdout
    from io import StringIO
    from unittest.mock import patch

    from phraseloom.cli import _dispatch

    stream = StringIO()
    with patch("builtins.input", side_effect=["4", "q"]), redirect_stdout(stream):
        self.assertEqual(_dispatch([]), 0)

    menu = stream.getvalue()
    self.assertIn("4) Entity workflow", menu)
    self.assertIn("Entity Workflow", menu)
```

- [ ] **Step 2: Run the alias tests and verify RED**

Run:

```powershell
python -m unittest tests.test_entity_workflow.EntityPackWorkflowCliTests.test_entity_interactive_aliases_open_entity_menu tests.test_entity_workflow.EntityPackWorkflowCliTests.test_top_level_interactive_option_opens_entity_menu
```

Expected: FAIL because `_dispatch(["entity"])` currently falls through to legacy parsing.

- [ ] **Step 3: Implement alias routing**

In `phraseloom/cli.py`, update the interactive import:

```python
from .interactive import (
    _normalize_optional_column,
    run_entity_interactive,
    run_interactive,
)
```

Then update `_dispatch()` before the existing `entity-tm` branch:

```python
if argv[0] in {"entity", "entity-interactive"}:
    return run_entity_interactive()
```

Update `_print_top_level_help()` under `Interactive:`:

```python
print("  phraseloom entity")
print("  phraseloom entity-interactive")
```

In `phraseloom/interactive.py`, add `run_entity_interactive` to `__all__`.

- [ ] **Step 4: Run the alias tests and verify GREEN**

Run:

```powershell
python -m unittest tests.test_entity_workflow.EntityPackWorkflowCliTests.test_entity_interactive_aliases_open_entity_menu tests.test_entity_workflow.EntityPackWorkflowCliTests.test_top_level_interactive_option_opens_entity_menu
```

Expected: PASS.

## Task 3: Entity Step Handlers

**Files:**
- Modify: `tests/test_entity_workflow.py`
- Modify: `phraseloom/interactive.py`

- [ ] **Step 1: Write failing tests for the four entity steps**

In `tests/test_entity_workflow.py`, add these methods to `EntityPackWorkflowCliTests`.

For step 1:

```python
def test_entity_interactive_step_1_writes_memory_workbook(self):
    from contextlib import redirect_stdout
    from io import StringIO
    from unittest.mock import patch

    from phraseloom.cli import _dispatch

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        tm_pairs_path = tmp_path / "tm_pairs.xlsx"
        output_path = tmp_path / "entity_memory.xlsx"
        _write_tm_pairs_workbook(tm_pairs_path)

        answers = iter(["1", str(tm_pairs_path), str(output_path), "2"])
        with patch("builtins.input", side_effect=lambda _="": next(answers)), redirect_stdout(StringIO()):
            self.assertEqual(_dispatch(["entity"]), 0)

        self.assertTrue(output_path.exists())
        self.assertEqual(
            _headers(output_path, schema.ENTITY_STRUCTURES_SHEET)[0],
            schema.STRUCTURE_ID_COLUMN,
        )
```

For step 2:

```python
def test_entity_interactive_step_2_writes_pack_with_optional_memory(self):
    from contextlib import redirect_stdout
    from io import StringIO
    from unittest.mock import patch

    from phraseloom.cli import _dispatch

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        todo_path = tmp_path / "source_translator_todo.xlsx"
        memory_path = tmp_path / "entity_memory.xlsx"
        output_path = tmp_path / "entity_pack.xlsx"
        _write_todo_workbook(
            todo_path,
            [
                {
                    schema.UNIT_ID_COLUMN: "U0001",
                    schema.UNIT_TYPE_COLUMN: "segment",
                    schema.SOURCE_UNIT_COLUMN: "Squirtle launched an attack and dealt damage.",
                    schema.TARGET_UNIT_COLUMN: None,
                },
                {
                    schema.UNIT_ID_COLUMN: "U0002",
                    schema.UNIT_TYPE_COLUMN: "segment",
                    schema.SOURCE_UNIT_COLUMN: "Login failed.",
                    schema.TARGET_UNIT_COLUMN: None,
                },
                {
                    schema.UNIT_ID_COLUMN: "U0003",
                    schema.UNIT_TYPE_COLUMN: "segment",
                    schema.SOURCE_UNIT_COLUMN: "Pikachu launched an attack and dealt damage.",
                    schema.TARGET_UNIT_COLUMN: None,
                },
            ],
        )
        _write_entity_tm_workbook(memory_path)

        answers = iter(["2", str(todo_path), str(memory_path), str(output_path), "2"])
        with patch("builtins.input", side_effect=lambda _="": next(answers)), redirect_stdout(StringIO()):
            self.assertEqual(_dispatch(["entity"]), 0)

        self.assertTrue(output_path.exists())
        structures = _rows_by_header(output_path, schema.ENTITY_STRUCTURES_SHEET)
        self.assertEqual(
            structures[0][schema.TARGET_STRUCTURE_COLUMN],
            "{entity1} launched a localized attack and dealt localized damage.",
        )
```

For step 3:

```python
def test_entity_interactive_step_3_fills_entity_pack(self):
    from contextlib import redirect_stdout
    from io import StringIO
    from unittest.mock import patch

    from phraseloom.cli import _dispatch
    from phraseloom.entity_workflow import prepare_entity_pack_workbook

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        todo_path = tmp_path / "source_translator_todo.xlsx"
        pack_path = tmp_path / "entity_pack.xlsx"
        output_path = tmp_path / "entity_pack_filled.xlsx"
        _write_todo_workbook(
            todo_path,
            [
                {
                    schema.UNIT_ID_COLUMN: "U0001",
                    schema.UNIT_TYPE_COLUMN: "segment",
                    schema.SOURCE_UNIT_COLUMN: "Squirtle launched an attack and dealt damage.",
                    schema.TARGET_UNIT_COLUMN: None,
                },
                {
                    schema.UNIT_ID_COLUMN: "U0002",
                    schema.UNIT_TYPE_COLUMN: "segment",
                    schema.SOURCE_UNIT_COLUMN: "Pikachu launched an attack and dealt damage.",
                    schema.TARGET_UNIT_COLUMN: None,
                },
            ],
        )
        prepare_entity_pack_workbook(todo_path, pack_path, min_group_size=2)
        _complete_entity_tables(
            pack_path,
            term_targets={"Squirtle": "Carapuce", "Pikachu": "Pikachu"},
        )

        answers = iter(["3", str(pack_path), str(output_path)])
        with patch("builtins.input", side_effect=lambda _="": next(answers)), redirect_stdout(StringIO()):
            self.assertEqual(_dispatch(["entity"]), 0)

        related_rows = _rows_by_header(output_path, schema.RELATED_UNITS_SHEET)
        self.assertEqual(
            [row[schema.TARGET_UNIT_COLUMN] for row in related_rows],
            [
                "Carapuce launched a localized attack and dealt localized damage.",
                "Pikachu launched a localized attack and dealt localized damage.",
            ],
        )
```

For step 4:

```python
def test_entity_interactive_step_4_merges_filled_entity_pack(self):
    from contextlib import redirect_stdout
    from io import StringIO
    from unittest.mock import patch

    from phraseloom.cli import _dispatch
    from phraseloom.entity_workflow import fill_entity_pack_workbook, prepare_entity_pack_workbook

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        todo_path = tmp_path / "source_translator_todo.xlsx"
        pack_path = tmp_path / "entity_pack.xlsx"
        filled_path = tmp_path / "entity_pack_filled.xlsx"
        output_path = tmp_path / "merged_todo.xlsx"
        _write_todo_workbook(
            todo_path,
            [
                {
                    schema.UNIT_ID_COLUMN: "U0001",
                    schema.UNIT_TYPE_COLUMN: "segment",
                    schema.SOURCE_UNIT_COLUMN: "Squirtle launched an attack and dealt damage.",
                    schema.TARGET_UNIT_COLUMN: None,
                },
                {
                    schema.UNIT_ID_COLUMN: "U0002",
                    schema.UNIT_TYPE_COLUMN: "segment",
                    schema.SOURCE_UNIT_COLUMN: "Login failed.",
                    schema.TARGET_UNIT_COLUMN: None,
                },
                {
                    schema.UNIT_ID_COLUMN: "U0003",
                    schema.UNIT_TYPE_COLUMN: "segment",
                    schema.SOURCE_UNIT_COLUMN: "Pikachu launched an attack and dealt damage.",
                    schema.TARGET_UNIT_COLUMN: None,
                },
            ],
        )
        prepare_entity_pack_workbook(todo_path, pack_path, min_group_size=2)
        _complete_entity_tables(
            pack_path,
            term_targets={"Squirtle": "Carapuce", "Pikachu": "Pikachu"},
        )
        _set_first_non_related_target(pack_path, "Login failed localized.")
        fill_entity_pack_workbook(pack_path, filled_path)

        answers = iter(["4", str(filled_path), str(output_path)])
        with patch("builtins.input", side_effect=lambda _="": next(answers)), redirect_stdout(StringIO()):
            self.assertEqual(_dispatch(["entity"]), 0)

        merged_rows = _rows_by_header(output_path, schema.TO_TRANSLATE_SHEET)
        self.assertEqual(
            [row[schema.TARGET_UNIT_COLUMN] for row in merged_rows],
            [
                "Carapuce launched a localized attack and dealt localized damage.",
                "Login failed localized.",
                "Pikachu launched a localized attack and dealt localized damage.",
            ],
        )
```

- [ ] **Step 2: Run the four step tests and verify RED**

Run:

```powershell
python -m unittest tests.test_entity_workflow.EntityPackWorkflowCliTests.test_entity_interactive_step_1_writes_memory_workbook tests.test_entity_workflow.EntityPackWorkflowCliTests.test_entity_interactive_step_2_writes_pack_with_optional_memory tests.test_entity_workflow.EntityPackWorkflowCliTests.test_entity_interactive_step_3_fills_entity_pack tests.test_entity_workflow.EntityPackWorkflowCliTests.test_entity_interactive_step_4_merges_filled_entity_pack
```

Expected: FAIL because the temporary entity menu does not dispatch steps 1 through 4.

- [ ] **Step 3: Implement the entity step handlers**

In `phraseloom/interactive.py`, import entity workflow functions and default path helpers:

```python
from .entity_workflow import (
    default_entity_filled_pack_output_path,
    default_entity_memory_output_path,
    default_entity_merged_todo_output_path,
    default_entity_pack_output_path,
    extract_entity_memory_workbook,
    fill_entity_pack_workbook,
    merge_entity_pack_workbook,
    prepare_entity_pack_workbook,
)
```

Replace `run_entity_interactive()` with:

```python
def run_entity_interactive(*, back_returns_to_main: bool = False) -> int:
    print("Entity Workflow")
    print()
    print("1) Build entity memory from TM reusable units")
    print("2) Prepare source entity pack")
    print("3) Fill completed entity pack")
    print("4) Merge filled entity pack back to translator todo")
    print("b) Back")
    print("q) Quit")

    action = _prompt_text("Choose entity step", default="2").lower()
    if action in {"q", "quit", "exit"}:
        print("Bye.")
        return 0
    if action in {"b", "back"}:
        if back_returns_to_main:
            return run_interactive()
        print("Bye.")
        return 0
    if action in {"1", "tm", "entity-tm", "memory"}:
        return _interactive_entity_tm()
    if action in {"2", "prepare", "pack"}:
        return _interactive_entity_prepare()
    if action in {"3", "fill", "fill-pack"}:
        return _interactive_entity_fill_pack()
    if action in {"4", "merge", "merge-pack"}:
        return _interactive_entity_merge_pack()

    print(f"Unknown entity step: {action}")
    return 2
```

Add the four handlers:

```python
def _interactive_entity_tm() -> int:
    input_path = _user_path(_prompt_text("TM reusable units path", required=True))
    output_path = _user_path(
        _prompt_text(
            "Output entity memory workbook",
            default=str(default_entity_memory_output_path(input_path)),
        )
    )
    min_group_size = _prompt_int(
        "Minimum variants for a reusable entity structure",
        default=3,
    )
    stats = extract_entity_memory_workbook(
        input_path,
        output_path,
        min_group_size=min_group_size,
    )
    _display_entity_extract_tm_stats(stats)
    return 0


def _interactive_entity_prepare() -> int:
    input_path = _user_path(_prompt_text("Translator todo path", required=True))
    tm_text = _prompt_text("Entity memory path (- for none)", default="-")
    tm_path = _user_path(tm_text) if _normalize_optional_column(tm_text) is not None else None
    output_path = _user_path(
        _prompt_text(
            "Output entity pack workbook",
            default=str(default_entity_pack_output_path(input_path)),
        )
    )
    min_group_size = _prompt_int(
        "Minimum variants for a reusable entity structure",
        default=3,
    )
    stats = prepare_entity_pack_workbook(
        input_path,
        output_path,
        tm_path=tm_path,
        min_group_size=min_group_size,
    )
    _display_entity_prepare_stats(stats)
    return 0


def _interactive_entity_fill_pack() -> int:
    input_path = _user_path(_prompt_text("Source entity pack path", required=True))
    output_path = _user_path(
        _prompt_text(
            "Output filled entity pack workbook",
            default=str(default_entity_filled_pack_output_path(input_path)),
        )
    )
    stats = fill_entity_pack_workbook(input_path, output_path)
    _display_entity_fill_stats(stats)
    return 0


def _interactive_entity_merge_pack() -> int:
    input_path = _user_path(_prompt_text("Filled source entity pack path", required=True))
    output_path = _user_path(
        _prompt_text(
            "Output merged translator todo workbook",
            default=str(default_entity_merged_todo_output_path(input_path)),
        )
    )
    stats = merge_entity_pack_workbook(input_path, output_path)
    _display_entity_merge_stats(stats)
    return 0
```

Add stats display helpers matching `phraseloom.cli`:

```python
def _display_entity_prepare_stats(stats: dict[str, int | str]) -> None:
    print(f"Wrote: {stats['output_path']}")
    print(f"Related units: {stats['related_unit_count']}")
    print(f"Non-entity units: {stats['non_entity_unit_count']}")
    print(f"Entity structures: {stats['entity_structure_count']}")
    print(f"Entity terms: {stats['entity_term_count']}")
    print(f"Prefilled structures: {stats['prefilled_structure_count']}")
    print(f"Prefilled terms: {stats['prefilled_term_count']}")


def _display_entity_extract_tm_stats(stats: dict[str, int | str]) -> None:
    print(f"Wrote: {stats['output_path']}")
    print(f"Entity structures: {stats['entity_structure_count']}")
    print(f"Entity terms: {stats['entity_term_count']}")


def _display_entity_fill_stats(stats: dict[str, int | str]) -> None:
    print(f"Wrote: {stats['output_path']}")
    print(f"Filled entity units: {stats['filled_entity_unit_count']}")


def _display_entity_merge_stats(stats: dict[str, int | str]) -> None:
    print(f"Wrote: {stats['output_path']}")
    print(f"Merged units: {stats['merged_unit_count']}")
```

Export the new handlers and display helpers in `__all__`.

- [ ] **Step 4: Run the four step tests and verify GREEN**

Run:

```powershell
python -m unittest tests.test_entity_workflow.EntityPackWorkflowCliTests.test_entity_interactive_step_1_writes_memory_workbook tests.test_entity_workflow.EntityPackWorkflowCliTests.test_entity_interactive_step_2_writes_pack_with_optional_memory tests.test_entity_workflow.EntityPackWorkflowCliTests.test_entity_interactive_step_3_fills_entity_pack tests.test_entity_workflow.EntityPackWorkflowCliTests.test_entity_interactive_step_4_merges_filled_entity_pack
```

Expected: PASS.

## Task 4: README And Full Verification

**Files:**
- Modify: `README.md`

- [ ] **Step 1: Update README with TUI entry points**

In `README.md`, in the entity engine section before the first `phraseloom entity-tm` command example, add:

```markdown
也可以用交互式菜单进入 entity workflow：

```bash
phraseloom
# 选择 4) Entity workflow

phraseloom entity
phraseloom entity-interactive
```

交互菜单只覆盖推荐的四步主流程；底层调试命令仍通过直接 CLI 子命令使用。
```

- [ ] **Step 2: Run entity and template focused tests**

Run:

```powershell
python -m unittest tests.test_template_workflow.TemplateDemoTests.test_interactive_menu_shows_three_step_workflow tests.test_entity_workflow.EntityPackWorkflowCliTests
```

Expected: PASS.

- [ ] **Step 3: Run the full test suite**

Run:

```powershell
python -m unittest
```

Expected: PASS.

- [ ] **Step 4: Inspect the final diff**

Run:

```powershell
git diff -- phraseloom/interactive.py phraseloom/cli.py tests/test_template_workflow.py tests/test_entity_workflow.py README.md
```

Expected: diff only contains entity TUI menu, alias routing, tests, and README documentation.

- [ ] **Step 5: Commit the implementation**

Run:

```powershell
git add phraseloom/interactive.py phraseloom/cli.py tests/test_template_workflow.py tests/test_entity_workflow.py README.md
git commit -m "feat: add entity workflow tui"
```

Expected: commit succeeds.

## Self-Review

Spec coverage:

- Top-level `phraseloom` menu entry is covered by Task 1.
- Direct `phraseloom entity` and `phraseloom entity-interactive` aliases are covered by Task 2.
- The four-step entity menu and behavior are covered by Task 3.
- Advanced/debug tools are intentionally excluded from the interactive menu.
- README documentation is covered by Task 4.

Placeholder scan: this plan contains no placeholder tasks or undefined behavior.

Type consistency: the plan uses existing entity workflow function names and stat keys already present in `phraseloom.cli`.
