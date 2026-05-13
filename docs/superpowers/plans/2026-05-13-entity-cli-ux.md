# Entity CLI UX Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a simplified four-command entity UX that produces one reusable entity memory workbook and one source entity pack workbook.

**Architecture:** Keep the existing low-level entity commands intact and add a compact pack workflow on top of `phraseloom.entity_workflow`. New APIs will write/read `related_units`, `non_related_units`, `entity_structures`, `entity_terms`, and hidden `_entity_map` sheets while reusing the current clustering, prefill, fill, and merge helpers. CLI additions stay in `phraseloom.cli`, with README documenting the user-facing workflow.

**Tech Stack:** Python 3.11+, openpyxl, unittest, existing PhraseLoom workbook schema and entity cluster strategy.

---

## File Structure

- Modify `phraseloom/workbook_schema.py`: add sheet constants for the compact entity pack.
- Modify `phraseloom/entity_workflow.py`: add default output path helpers, high-level pack APIs, pack sheet read/write helpers, hidden technical columns/sheets, and compatibility sheet lookup helpers.
- Modify `phraseloom/cli.py`: add `entity-tm`, `entity-prepare`, `entity-fill-pack`, and `entity-merge-pack` commands while keeping existing low-level commands.
- Modify `tests/test_entity_workflow.py`: add focused tests for the four new UX operations and CLI dispatch.
- Modify `README.md`: describe the simplified workflow first and keep the low-level commands as advanced/debug commands.

## Implementation Notes

The compact pack should keep `original_index` in `related_units` and `non_related_units` as a hidden first column. This preserves deterministic merge behavior without making users manage the column.

The hidden `_entity_map` sheet should use the existing `ENTITY_SOURCE_MAP_COLUMNS` layout. It only needs entity-related rows for fill; merge order comes from the hidden `original_index` column in both visible unit sheets.

Existing low-level workbooks should continue to work with old APIs. The new pack fill helper should also accept old low-level entity workbooks by looking for `to_translate` and `entity_source_map` when `related_units` and `_entity_map` are not present.

## Tasks

### Task 1: Add Entity Memory Command With L10n Default Path

**Files:**
- Modify: `tests/test_entity_workflow.py`
- Modify: `phraseloom/entity_workflow.py`
- Modify: `phraseloom/cli.py`

- [ ] **Step 1: Write the failing CLI test**

Add this test class near the existing `EntityWorkflowCliTests` in `tests/test_entity_workflow.py`:

```python
class EntityPackWorkflowCliTests(unittest.TestCase):
    def test_entity_tm_cli_writes_memory_workbook_to_l10n(self):
        from phraseloom.cli import _dispatch

        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            tm_pairs_path = tmp_path / "tm_pairs.xlsx"
            _write_tm_pairs_workbook(tm_pairs_path)

            self.assertEqual(
                _dispatch(
                    [
                        "entity-tm",
                        str(tm_pairs_path),
                        "--min-group-size",
                        "2",
                    ]
                ),
                0,
            )

            entity_memory_path = (
                tmp_path
                / "tm_pairs_l10n"
                / "tm_pairs_entity_memory.xlsx"
            )
            self.assertTrue(entity_memory_path.exists())
            self.assertEqual(
                _headers(entity_memory_path, schema.ENTITY_STRUCTURES_SHEET),
                [
                    schema.STRUCTURE_ID_COLUMN,
                    schema.SOURCE_STRUCTURE_COLUMN,
                    schema.TARGET_STRUCTURE_COLUMN,
                    schema.COVERAGE_COUNT_COLUMN,
                    schema.CONFIDENCE_COLUMN,
                    schema.RISK_COLUMN,
                    schema.STATUS_COLUMN,
                    schema.SAMPLE_SOURCES_COLUMN,
                    schema.ROW_NUMBERS_COLUMN,
                    schema.WARNING_COLUMN,
                ],
            )
            self.assertEqual(
                _headers(entity_memory_path, schema.ENTITY_TERMS_SHEET),
                [
                    schema.TERM_ID_COLUMN,
                    schema.SOURCE_ENTITY_COLUMN,
                    schema.TARGET_ENTITY_COLUMN,
                    schema.OCCURRENCE_COUNT_COLUMN,
                    schema.STRUCTURE_IDS_COLUMN,
                    schema.STATUS_COLUMN,
                    schema.WARNING_COLUMN,
                ],
            )
```

- [ ] **Step 2: Run the test to verify RED**

Run:

```bash
python3 -m unittest tests.test_entity_workflow.EntityPackWorkflowCliTests.test_entity_tm_cli_writes_memory_workbook_to_l10n -v
```

Expected: FAIL because `_dispatch(["entity-tm", ...])` is not registered and no `*_entity_memory.xlsx` default exists.

- [ ] **Step 3: Add default path helpers and memory API**

In `phraseloom/entity_workflow.py`, add these helpers after `ClusterProbeStrategy`:

```python
def _default_entity_work_dir(input_path: Path) -> Path:
    return input_path.parent / f"{input_path.stem}_l10n"


def default_entity_memory_output_path(input_path: str | Path) -> Path:
    input_path = Path(input_path)
    return _default_entity_work_dir(input_path) / f"{input_path.stem}_entity_memory.xlsx"


def default_entity_pack_output_path(input_path: str | Path) -> Path:
    input_path = Path(input_path)
    return _default_entity_work_dir(input_path) / f"{input_path.stem}_entity_pack.xlsx"


def default_entity_filled_pack_output_path(input_path: str | Path) -> Path:
    input_path = Path(input_path)
    return input_path.with_name(f"{input_path.stem}_filled.xlsx")


def default_entity_merged_todo_output_path(input_path: str | Path) -> Path:
    input_path = Path(input_path)
    return input_path.with_name(f"{input_path.stem}_merged_todo.xlsx")


def extract_entity_memory_workbook(
    input_path: str | Path,
    output_path: str | Path,
    *,
    min_group_size: int = 3,
    strategy: EntityExtractionStrategy | None = None,
) -> dict[str, int | str]:
    return extract_entity_tm_workbook(
        input_path,
        output_path,
        min_group_size=min_group_size,
        strategy=strategy,
    )
```

Add the new names to `__all__`:

```python
    "default_entity_filled_pack_output_path",
    "default_entity_memory_output_path",
    "default_entity_merged_todo_output_path",
    "default_entity_pack_output_path",
    "extract_entity_memory_workbook",
```

- [ ] **Step 4: Register the `entity-tm` CLI command**

In `phraseloom/cli.py`, extend the entity import:

```python
from .entity_workflow import (
    default_entity_memory_output_path,
    extract_entity_memory_workbook,
    extract_entity_tm_workbook,
    fill_entity_workbook,
    merge_entity_workbooks,
    prefill_entity_workbook,
    split_entity_workbook,
)
```

In `_dispatch`, add the high-level command before the low-level entity commands:

```python
    if argv[0] == "entity-tm":
        return _main_entity_tm(argv[1:])
```

Add the command implementation before `_main_entity_split`:

```python
def _main_entity_tm(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Build reusable entity memory from a TM reusable-units workbook."
    )
    parser.add_argument("input", type=Path, help="TM reusable units .xlsx file")
    parser.add_argument("-o", "--output", type=Path, help="Entity memory output .xlsx")
    parser.add_argument("--min-group-size", type=int, default=3)
    args = parser.parse_args(argv)
    output = args.output or default_entity_memory_output_path(args.input)
    stats = extract_entity_memory_workbook(
        args.input,
        output,
        min_group_size=args.min_group_size,
    )
    _print_entity_extract_tm_stats(stats)
    return 0
```

Add `_main_entity_tm` to `__all__`.

- [ ] **Step 5: Run the test to verify GREEN**

Run:

```bash
python3 -m unittest tests.test_entity_workflow.EntityPackWorkflowCliTests.test_entity_tm_cli_writes_memory_workbook_to_l10n -v
```

Expected: PASS.

- [ ] **Step 6: Commit**

Run:

```bash
git add phraseloom/entity_workflow.py phraseloom/cli.py tests/test_entity_workflow.py
git commit -m "feat: add entity memory cli"
```

### Task 2: Prepare One Source Entity Pack With Optional TM Prefill

**Files:**
- Modify: `phraseloom/workbook_schema.py`
- Modify: `tests/test_entity_workflow.py`
- Modify: `phraseloom/entity_workflow.py`
- Modify: `phraseloom/cli.py`

- [ ] **Step 1: Write the failing pack preparation test**

Add this helper near `_headers` in `tests/test_entity_workflow.py`:

```python
def _sheet_state(path: Path, sheet_name: str) -> str:
    wb = load_workbook(path, data_only=True)
    try:
        return wb[sheet_name].sheet_state
    finally:
        wb.close()
```

Add this test to `EntityPackWorkflowCliTests`:

```python
    def test_entity_prepare_cli_writes_single_pack_with_prefill(self):
        from phraseloom.cli import _dispatch

        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            todo_path = tmp_path / "source_translator_todo.xlsx"
            entity_memory_path = tmp_path / "tm_entity_memory.xlsx"
            _write_todo_workbook(
                todo_path,
                [
                    {
                        schema.UNIT_ID_COLUMN: "U0001",
                        schema.UNIT_TYPE_COLUMN: "segment",
                        schema.SOURCE_UNIT_COLUMN: "Squirtle launched an attack and dealt damage.",
                        schema.TARGET_UNIT_COLUMN: None,
                        schema.COVERAGE_COUNT_COLUMN: 1,
                    },
                    {
                        schema.UNIT_ID_COLUMN: "U0002",
                        schema.UNIT_TYPE_COLUMN: "segment",
                        schema.SOURCE_UNIT_COLUMN: "Login failed.",
                        schema.TARGET_UNIT_COLUMN: None,
                        schema.COVERAGE_COUNT_COLUMN: 1,
                    },
                    {
                        schema.UNIT_ID_COLUMN: "U0003",
                        schema.UNIT_TYPE_COLUMN: "segment",
                        schema.SOURCE_UNIT_COLUMN: "Pikachu launched an attack and dealt damage.",
                        schema.TARGET_UNIT_COLUMN: None,
                        schema.COVERAGE_COUNT_COLUMN: 1,
                    },
                ],
            )
            _write_entity_tm_workbook(entity_memory_path)

            self.assertEqual(
                _dispatch(
                    [
                        "entity-prepare",
                        str(todo_path),
                        "--tm",
                        str(entity_memory_path),
                        "--min-group-size",
                        "2",
                    ]
                ),
                0,
            )

            pack_path = (
                tmp_path
                / "source_translator_todo_l10n"
                / "source_translator_todo_entity_pack.xlsx"
            )
            self.assertTrue(pack_path.exists())
            wb = load_workbook(pack_path, data_only=True)
            try:
                self.assertEqual(
                    wb.sheetnames[:5],
                    [
                        schema.RELATED_UNITS_SHEET,
                        schema.NON_RELATED_UNITS_SHEET,
                        schema.ENTITY_STRUCTURES_SHEET,
                        schema.ENTITY_TERMS_SHEET,
                        schema.ENTITY_MAP_SHEET,
                    ],
                )
            finally:
                wb.close()

            self.assertEqual(_sheet_state(pack_path, schema.ENTITY_MAP_SHEET), "hidden")
            self.assertEqual(
                _sheet_state(pack_path, schema.METADATA_SHEET),
                "hidden",
            )
            related_rows = _rows_by_header(pack_path, schema.RELATED_UNITS_SHEET)
            non_related_rows = _rows_by_header(pack_path, schema.NON_RELATED_UNITS_SHEET)
            self.assertEqual(
                [row[schema.UNIT_ID_COLUMN] for row in related_rows],
                ["U0001", "U0003"],
            )
            self.assertEqual(
                [row[schema.UNIT_ID_COLUMN] for row in non_related_rows],
                ["U0002"],
            )
            structures = _rows_by_header(pack_path, schema.ENTITY_STRUCTURES_SHEET)
            self.assertEqual(
                structures[0][schema.TARGET_STRUCTURE_COLUMN],
                "{entity1} launched a localized attack and dealt localized damage.",
            )
            terms = _rows_by_header(pack_path, schema.ENTITY_TERMS_SHEET)
            self.assertEqual(
                {
                    row[schema.SOURCE_ENTITY_COLUMN]: row[schema.TARGET_ENTITY_COLUMN]
                    for row in terms
                },
                {"Pikachu": "Pikachu", "Squirtle": "Carapuce"},
            )
```

- [ ] **Step 2: Run the test to verify RED**

Run:

```bash
python3 -m unittest tests.test_entity_workflow.EntityPackWorkflowCliTests.test_entity_prepare_cli_writes_single_pack_with_prefill -v
```

Expected: FAIL because `RELATED_UNITS_SHEET`, `NON_RELATED_UNITS_SHEET`, `ENTITY_MAP_SHEET`, `entity-prepare`, and pack writing do not exist.

- [ ] **Step 3: Add compact pack sheet constants**

In `phraseloom/workbook_schema.py`, add these sheet constants after `ENTITY_SOURCE_MAP_SHEET`:

```python
RELATED_UNITS_SHEET = "related_units"
NON_RELATED_UNITS_SHEET = "non_related_units"
ENTITY_MAP_SHEET = "_entity_map"
```

- [ ] **Step 4: Add pack prefill and writing helpers**

In `phraseloom/entity_workflow.py`, extract the current prefill logic into a workbook helper and make `prefill_entity_workbook()` call it:

```python
def _prefill_entity_sheets(wb, tm_path: Path) -> tuple[int, int]:
    structure_prefills = _load_unique_prefills(
        tm_path,
        schema.ENTITY_STRUCTURES_SHEET,
        schema.SOURCE_STRUCTURE_COLUMN,
        schema.TARGET_STRUCTURE_COLUMN,
    )
    term_prefills = _load_unique_prefills(
        tm_path,
        schema.ENTITY_TERMS_SHEET,
        schema.SOURCE_ENTITY_COLUMN,
        schema.TARGET_ENTITY_COLUMN,
    )
    prefilled_structure_count = _apply_prefills(
        wb[schema.ENTITY_STRUCTURES_SHEET],
        source_column=schema.SOURCE_STRUCTURE_COLUMN,
        target_column=schema.TARGET_STRUCTURE_COLUMN,
        prefills=structure_prefills,
        ready_on_prefill=False,
    )
    prefilled_term_count = _apply_prefills(
        wb[schema.ENTITY_TERMS_SHEET],
        source_column=schema.SOURCE_ENTITY_COLUMN,
        target_column=schema.TARGET_ENTITY_COLUMN,
        prefills=term_prefills,
        ready_on_prefill=True,
    )
    return prefilled_structure_count, prefilled_term_count
```

Replace the body inside `prefill_entity_workbook()` after `wb = load_workbook(entity_input_path)` with:

```python
        prefilled_structure_count, prefilled_term_count = _prefill_entity_sheets(
            wb,
            tm_path,
        )
        _save_workbook(wb, output_path)
```

Add these helpers near the existing workbook writers:

```python
def prepare_entity_pack_workbook(
    input_path: str | Path,
    output_path: str | Path,
    *,
    tm_path: str | Path | None = None,
    min_group_size: int = 3,
    strategy: EntityExtractionStrategy | None = None,
) -> dict[str, int | str]:
    input_path = Path(input_path)
    output_path = Path(output_path)
    rows, headers = _read_unit_rows(input_path, schema.TO_TRANSLATE_SHEET)
    active_strategy = strategy or ClusterProbeStrategy(min_group_size=min_group_size)
    clusters = active_strategy.find_clusters(
        [(row.source_unit, row.target_unit) for row in rows]
    )
    entity_rows, structures, terms, source_map = _build_entity_split(rows, clusters)
    entity_indices = {row.original_index for row in entity_rows}
    non_entity_rows = [row for row in rows if row.original_index not in entity_indices]
    wb = _build_entity_pack_workbook(
        input_path,
        headers,
        entity_rows,
        non_entity_rows,
        structures,
        terms,
        source_map,
    )
    prefilled_structure_count = 0
    prefilled_term_count = 0
    if tm_path is not None:
        prefilled_structure_count, prefilled_term_count = _prefill_entity_sheets(
            wb,
            Path(tm_path),
        )
    _save_workbook(wb, output_path)
    return {
        "related_unit_count": len(entity_rows),
        "non_related_unit_count": len(non_entity_rows),
        "entity_structure_count": len(structures),
        "entity_term_count": len(terms),
        "prefilled_structure_count": prefilled_structure_count,
        "prefilled_term_count": prefilled_term_count,
        "output_path": str(output_path),
    }


def _build_entity_pack_workbook(
    input_path: Path,
    headers: list[str],
    entity_rows: list[UnitRow],
    non_entity_rows: list[UnitRow],
    structures: list[dict[str, object]],
    terms: list[dict[str, object]],
    source_map: list[dict[str, object]],
):
    wb = Workbook()
    related_ws = wb.active
    related_ws.title = schema.RELATED_UNITS_SHEET
    _append_unit_rows(related_ws, headers, entity_rows, include_original_index=True)
    _hide_original_index_column(related_ws)
    non_related_ws = wb.create_sheet(schema.NON_RELATED_UNITS_SHEET)
    _append_unit_rows(non_related_ws, headers, non_entity_rows, include_original_index=True)
    _hide_original_index_column(non_related_ws)
    _append_dict_sheet(wb.create_sheet(schema.ENTITY_STRUCTURES_SHEET), ENTITY_STRUCTURE_COLUMNS, structures)
    _append_dict_sheet(wb.create_sheet(schema.ENTITY_TERMS_SHEET), ENTITY_TERM_COLUMNS, terms)
    entity_map_ws = wb.create_sheet(schema.ENTITY_MAP_SHEET)
    _append_dict_sheet(entity_map_ws, ENTITY_SOURCE_MAP_COLUMNS, source_map)
    entity_map_ws.sheet_state = "hidden"
    _copy_support_sheets(input_path, wb, exclude={schema.TO_TRANSLATE_SHEET})
    if schema.METADATA_SHEET not in wb.sheetnames:
        metadata_ws = wb.create_sheet(schema.METADATA_SHEET)
        metadata_ws.append(schema.METADATA_COLUMNS)
        metadata_ws.append([schema.SCHEMA_VERSION_KEY, schema.SCHEMA_VERSION])
    wb[schema.METADATA_SHEET].sheet_state = "hidden"
    return wb


def _hide_original_index_column(ws) -> None:
    headers = _header_values(ws)
    if headers and headers[0] == schema.ORIGINAL_INDEX_COLUMN:
        ws.column_dimensions["A"].hidden = True
```

Add `prepare_entity_pack_workbook` to `__all__`.

- [ ] **Step 5: Register the `entity-prepare` CLI command**

In `phraseloom/cli.py`, extend the entity import:

```python
from .entity_workflow import (
    default_entity_memory_output_path,
    default_entity_pack_output_path,
    extract_entity_memory_workbook,
    extract_entity_tm_workbook,
    fill_entity_workbook,
    merge_entity_workbooks,
    prepare_entity_pack_workbook,
    prefill_entity_workbook,
    split_entity_workbook,
)
```

In `_dispatch`, add:

```python
    if argv[0] == "entity-prepare":
        return _main_entity_prepare(argv[1:])
```

Add the command and stats printer:

```python
def _main_entity_prepare(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Prepare one source entity pack with related and non-related units."
    )
    parser.add_argument("input", type=Path, help="Translator todo .xlsx file")
    parser.add_argument("-o", "--output", type=Path, help="Entity pack output .xlsx")
    parser.add_argument("--tm", type=Path, help="Entity memory workbook used to prefill the pack")
    parser.add_argument("--min-group-size", type=int, default=3)
    args = parser.parse_args(argv)
    output = args.output or default_entity_pack_output_path(args.input)
    stats = prepare_entity_pack_workbook(
        args.input,
        output,
        tm_path=args.tm,
        min_group_size=args.min_group_size,
    )
    _print_entity_prepare_stats(stats)
    return 0


def _print_entity_prepare_stats(stats: dict[str, int | str]) -> None:
    print(f"Wrote: {stats['output_path']}")
    print(f"Related units: {stats['related_unit_count']}")
    print(f"Non-related units: {stats['non_related_unit_count']}")
    print(f"Entity structures: {stats['entity_structure_count']}")
    print(f"Entity terms: {stats['entity_term_count']}")
    print(f"Prefilled structures: {stats['prefilled_structure_count']}")
    print(f"Prefilled terms: {stats['prefilled_term_count']}")
```

Add `_main_entity_prepare` and `_print_entity_prepare_stats` to `__all__`.

- [ ] **Step 6: Run the test to verify GREEN**

Run:

```bash
python3 -m unittest tests.test_entity_workflow.EntityPackWorkflowCliTests.test_entity_prepare_cli_writes_single_pack_with_prefill -v
```

Expected: PASS.

- [ ] **Step 7: Run existing low-level entity tests**

Run:

```bash
python3 -m unittest tests.test_entity_workflow.EntityWorkflowTests tests.test_entity_workflow.EntityWorkflowCliTests -v
```

Expected: PASS. This confirms the refactored prefill logic did not break existing commands.

- [ ] **Step 8: Commit**

Run:

```bash
git add phraseloom/workbook_schema.py phraseloom/entity_workflow.py phraseloom/cli.py tests/test_entity_workflow.py
git commit -m "feat: prepare compact entity pack"
```

### Task 3: Fill Related Units Inside A Pack

**Files:**
- Modify: `tests/test_entity_workflow.py`
- Modify: `phraseloom/entity_workflow.py`
- Modify: `phraseloom/cli.py`

- [ ] **Step 1: Add a helper for completing pack entity tables**

Add this helper near `_complete_entity_workbook` in `tests/test_entity_workflow.py`:

```python
def _complete_entity_tables(
    path: Path,
    *,
    term_targets: dict[str, str],
) -> None:
    wb = load_workbook(path)
    try:
        structures = wb[schema.ENTITY_STRUCTURES_SHEET]
        structure_headers = [cell.value for cell in structures[1]]
        structures.cell(
            row=2,
            column=structure_headers.index(schema.TARGET_STRUCTURE_COLUMN) + 1,
        ).value = "{entity1} launched a localized attack and dealt localized damage."
        structures.cell(
            row=2,
            column=structure_headers.index(schema.STATUS_COLUMN) + 1,
        ).value = "ready"

        terms = wb[schema.ENTITY_TERMS_SHEET]
        term_headers = [cell.value for cell in terms[1]]
        source_index = term_headers.index(schema.SOURCE_ENTITY_COLUMN) + 1
        target_index = term_headers.index(schema.TARGET_ENTITY_COLUMN) + 1
        status_index = term_headers.index(schema.STATUS_COLUMN) + 1
        for row_number in range(2, terms.max_row + 1):
            source_entity = terms.cell(row=row_number, column=source_index).value
            target_entity = term_targets.get(str(source_entity))
            if target_entity:
                terms.cell(row=row_number, column=target_index).value = target_entity
                terms.cell(row=row_number, column=status_index).value = "ready"
        wb.save(path)
    finally:
        wb.close()
```

- [ ] **Step 2: Write the failing fill-pack test**

Add this test to `EntityPackWorkflowCliTests`:

```python
    def test_entity_fill_pack_cli_writes_targets_to_related_units(self):
        from phraseloom.entity_workflow import prepare_entity_pack_workbook
        from phraseloom.cli import _dispatch

        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            todo_path = tmp_path / "source_translator_todo.xlsx"
            pack_path = tmp_path / "source_entity_pack.xlsx"
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
            prepare_entity_pack_workbook(
                todo_path,
                pack_path,
                min_group_size=2,
            )
            _complete_entity_tables(
                pack_path,
                term_targets={"Squirtle": "Carapuce", "Pikachu": "Pikachu"},
            )

            self.assertEqual(_dispatch(["entity-fill-pack", str(pack_path)]), 0)

            filled_pack_path = tmp_path / "source_entity_pack_filled.xlsx"
            self.assertTrue(filled_pack_path.exists())
            related_rows = _rows_by_header(filled_pack_path, schema.RELATED_UNITS_SHEET)
            self.assertEqual(
                [row[schema.TARGET_UNIT_COLUMN] for row in related_rows],
                [
                    "Carapuce launched a localized attack and dealt localized damage.",
                    "Pikachu launched a localized attack and dealt localized damage.",
                ],
            )
            map_rows = _rows_by_header(filled_pack_path, schema.ENTITY_MAP_SHEET)
            self.assertEqual(
                [row[schema.FILL_STATUS_COLUMN] for row in map_rows],
                ["filled", "filled"],
            )
```

- [ ] **Step 3: Run the test to verify RED**

Run:

```bash
python3 -m unittest tests.test_entity_workflow.EntityPackWorkflowCliTests.test_entity_fill_pack_cli_writes_targets_to_related_units -v
```

Expected: FAIL because `entity-fill-pack` and `fill_entity_pack_workbook()` do not exist.

- [ ] **Step 4: Add compatibility sheet lookup and pack fill API**

In `phraseloom/entity_workflow.py`, add these helpers near `_one_based_columns`:

```python
def _worksheet_by_name(wb, names: list[str]):
    for name in names:
        if name in wb.sheetnames:
            return wb[name]
    expected = " or ".join(names)
    raise WorkflowError(f"Workbook is missing required sheet: {expected}")
```

Add the public fill API near `fill_entity_workbook()`:

```python
def fill_entity_pack_workbook(
    pack_input_path: str | Path,
    output_path: str | Path,
) -> dict[str, int | str]:
    pack_input_path = Path(pack_input_path)
    output_path = Path(output_path)
    wb = load_workbook(pack_input_path)
    try:
        structures = _load_rows_by_key(
            wb[schema.ENTITY_STRUCTURES_SHEET],
            schema.STRUCTURE_ID_COLUMN,
        )
        terms = _load_rows_by_key(
            wb[schema.ENTITY_TERMS_SHEET],
            schema.SOURCE_ENTITY_COLUMN,
        )
        related_ws = _worksheet_by_name(
            wb,
            [schema.RELATED_UNITS_SHEET, schema.TO_TRANSLATE_SHEET],
        )
        entity_map_ws = _worksheet_by_name(
            wb,
            [schema.ENTITY_MAP_SHEET, schema.ENTITY_SOURCE_MAP_SHEET],
        )
        related_by_original_index = _todo_rows_by_original_index(related_ws)
        filled_count = _fill_source_map_rows(
            entity_map_ws,
            related_ws,
            related_by_original_index,
            structures,
            terms,
        )
        _save_workbook(wb, output_path)
    finally:
        wb.close()

    return {
        "filled_entity_unit_count": filled_count,
        "output_path": str(output_path),
    }
```

Add `fill_entity_pack_workbook` to `__all__`.

- [ ] **Step 5: Register the `entity-fill-pack` CLI command**

In `phraseloom/cli.py`, extend the entity import:

```python
from .entity_workflow import (
    default_entity_filled_pack_output_path,
    default_entity_memory_output_path,
    default_entity_pack_output_path,
    extract_entity_memory_workbook,
    extract_entity_tm_workbook,
    fill_entity_pack_workbook,
    fill_entity_workbook,
    merge_entity_workbooks,
    prepare_entity_pack_workbook,
    prefill_entity_workbook,
    split_entity_workbook,
)
```

In `_dispatch`, add:

```python
    if argv[0] == "entity-fill-pack":
        return _main_entity_fill_pack(argv[1:])
```

Add the command:

```python
def _main_entity_fill_pack(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Fill ready entity structures and terms back into related_units."
    )
    parser.add_argument("input", type=Path, help="Source entity pack .xlsx file")
    parser.add_argument("-o", "--output", type=Path, help="Filled entity pack output .xlsx")
    parser.add_argument(
        "--in-place",
        action="store_true",
        help="Update the input pack instead of writing a new file",
    )
    args = parser.parse_args(argv)
    output = args.input if args.in_place else (
        args.output or default_entity_filled_pack_output_path(args.input)
    )
    stats = fill_entity_pack_workbook(args.input, output)
    _print_entity_fill_stats(stats)
    return 0
```

Add `_main_entity_fill_pack` to `__all__`.

- [ ] **Step 6: Run the test to verify GREEN**

Run:

```bash
python3 -m unittest tests.test_entity_workflow.EntityPackWorkflowCliTests.test_entity_fill_pack_cli_writes_targets_to_related_units -v
```

Expected: PASS.

- [ ] **Step 7: Run existing fill tests**

Run:

```bash
python3 -m unittest tests.test_entity_workflow.EntityWorkflowTests.test_fill_writes_ready_entity_targets tests.test_entity_workflow.EntityWorkflowTests.test_fill_blocks_when_entity_translation_is_missing -v
```

Expected: PASS. This keeps the old `entity-fill` behavior intact.

- [ ] **Step 8: Commit**

Run:

```bash
git add phraseloom/entity_workflow.py phraseloom/cli.py tests/test_entity_workflow.py
git commit -m "feat: fill compact entity pack"
```

### Task 4: Merge A Filled Pack Back Into A Normal Todo Workbook

**Files:**
- Modify: `tests/test_entity_workflow.py`
- Modify: `phraseloom/entity_workflow.py`
- Modify: `phraseloom/cli.py`

- [ ] **Step 1: Write the failing merge-pack test**

Add this helper near `_set_first_todo_target`:

```python
def _set_first_non_related_target(path: Path, target: str) -> None:
    wb = load_workbook(path)
    try:
        ws = wb[schema.NON_RELATED_UNITS_SHEET]
        headers = [cell.value for cell in ws[1]]
        ws.cell(
            row=2,
            column=headers.index(schema.TARGET_UNIT_COLUMN) + 1,
        ).value = target
        wb.save(path)
    finally:
        wb.close()
```

Add this test to `EntityPackWorkflowCliTests`:

```python
    def test_entity_merge_pack_cli_restores_full_todo_order(self):
        from phraseloom.entity_workflow import (
            fill_entity_pack_workbook,
            prepare_entity_pack_workbook,
        )
        from phraseloom.cli import _dispatch

        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            todo_path = tmp_path / "source_translator_todo.xlsx"
            pack_path = tmp_path / "source_entity_pack.xlsx"
            filled_pack_path = tmp_path / "source_entity_pack_filled.xlsx"
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
            prepare_entity_pack_workbook(
                todo_path,
                pack_path,
                min_group_size=2,
            )
            _complete_entity_tables(
                pack_path,
                term_targets={"Squirtle": "Carapuce", "Pikachu": "Pikachu"},
            )
            _set_first_non_related_target(pack_path, "Login failed localized.")
            fill_entity_pack_workbook(pack_path, filled_pack_path)

            self.assertEqual(_dispatch(["entity-merge-pack", str(filled_pack_path)]), 0)

            merged_path = tmp_path / "source_entity_pack_filled_merged_todo.xlsx"
            self.assertTrue(merged_path.exists())
            self.assertNotIn(
                schema.ORIGINAL_INDEX_COLUMN,
                _headers(merged_path, schema.TO_TRANSLATE_SHEET),
            )
            merged_rows = _rows_by_header(merged_path, schema.TO_TRANSLATE_SHEET)
            self.assertEqual(
                [row[schema.UNIT_ID_COLUMN] for row in merged_rows],
                ["U0001", "U0002", "U0003"],
            )
            self.assertEqual(
                [row[schema.TARGET_UNIT_COLUMN] for row in merged_rows],
                [
                    "Carapuce launched a localized attack and dealt localized damage.",
                    "Login failed localized.",
                    "Pikachu launched a localized attack and dealt localized damage.",
                ],
            )
```

- [ ] **Step 2: Run the test to verify RED**

Run:

```bash
python3 -m unittest tests.test_entity_workflow.EntityPackWorkflowCliTests.test_entity_merge_pack_cli_restores_full_todo_order -v
```

Expected: FAIL because `entity-merge-pack` and `merge_entity_pack_workbook()` do not exist.

- [ ] **Step 3: Add pack row reading and merge API**

In `phraseloom/entity_workflow.py`, add this helper near `_read_unit_rows`:

```python
def _read_pack_unit_rows(
    path: Path,
    preferred_sheet_name: str,
    fallback_sheet_name: str | None = None,
) -> tuple[list[UnitRow], list[str]]:
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet_name = preferred_sheet_name
        if sheet_name not in wb.sheetnames:
            if fallback_sheet_name is None or fallback_sheet_name not in wb.sheetnames:
                raise WorkflowError(f"Workbook is missing required sheet: {preferred_sheet_name}")
            sheet_name = fallback_sheet_name
        ws = wb[sheet_name]
        headers = _header_values(ws)
        rows: list[UnitRow] = []
        for sequence_index, row in enumerate(
            ws.iter_rows(min_row=2, values_only=True),
            start=1,
        ):
            values = {
                header: row[index] if index < len(row) else None
                for index, header in enumerate(headers)
                if header
            }
            original_index = int(
                values.get(schema.ORIGINAL_INDEX_COLUMN) or sequence_index
            )
            if values.get(schema.SOURCE_UNIT_COLUMN):
                rows.append(UnitRow(original_index, values))
        return rows, headers
    finally:
        wb.close()
```

Add the public merge API near `merge_entity_workbooks()`:

```python
def merge_entity_pack_workbook(
    pack_path: str | Path,
    output_path: str | Path,
) -> dict[str, int | str]:
    pack_path = Path(pack_path)
    output_path = Path(output_path)
    related_rows, related_headers = _read_pack_unit_rows(
        pack_path,
        schema.RELATED_UNITS_SHEET,
        schema.TO_TRANSLATE_SHEET,
    )
    non_related_rows, non_related_headers = _read_pack_unit_rows(
        pack_path,
        schema.NON_RELATED_UNITS_SHEET,
    )
    headers = [
        header
        for header in (related_headers or non_related_headers)
        if header != schema.ORIGINAL_INDEX_COLUMN
    ]
    merged_rows = _merge_unit_rows(related_rows + non_related_rows)

    wb = Workbook()
    ws = wb.active
    ws.title = schema.TO_TRANSLATE_SHEET
    ws.append(headers)
    for row in merged_rows:
        ws.append([row.values.get(header) for header in headers])
    _copy_support_sheets(
        pack_path,
        wb,
        exclude={
            schema.RELATED_UNITS_SHEET,
            schema.NON_RELATED_UNITS_SHEET,
            schema.ENTITY_STRUCTURES_SHEET,
            schema.ENTITY_TERMS_SHEET,
            schema.ENTITY_MAP_SHEET,
            schema.TO_TRANSLATE_SHEET,
            schema.ENTITY_SOURCE_MAP_SHEET,
        },
    )
    _save_workbook(wb, output_path)
    return {
        "merged_unit_count": len(merged_rows),
        "output_path": str(output_path),
    }
```

Add `merge_entity_pack_workbook` to `__all__`.

- [ ] **Step 4: Register the `entity-merge-pack` CLI command**

In `phraseloom/cli.py`, extend the entity import:

```python
from .entity_workflow import (
    default_entity_filled_pack_output_path,
    default_entity_memory_output_path,
    default_entity_merged_todo_output_path,
    default_entity_pack_output_path,
    extract_entity_memory_workbook,
    extract_entity_tm_workbook,
    fill_entity_pack_workbook,
    fill_entity_workbook,
    merge_entity_pack_workbook,
    merge_entity_workbooks,
    prepare_entity_pack_workbook,
    prefill_entity_workbook,
    split_entity_workbook,
)
```

In `_dispatch`, add:

```python
    if argv[0] == "entity-merge-pack":
        return _main_entity_merge_pack(argv[1:])
```

Add the command:

```python
def _main_entity_merge_pack(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Merge related_units and non_related_units into a normal translator todo."
    )
    parser.add_argument("input", type=Path, help="Filled source entity pack .xlsx file")
    parser.add_argument("-o", "--output", type=Path, help="Merged todo output .xlsx")
    args = parser.parse_args(argv)
    output = args.output or default_entity_merged_todo_output_path(args.input)
    stats = merge_entity_pack_workbook(args.input, output)
    _print_entity_merge_stats(stats)
    return 0
```

Add `_main_entity_merge_pack` to `__all__`.

- [ ] **Step 5: Run the test to verify GREEN**

Run:

```bash
python3 -m unittest tests.test_entity_workflow.EntityPackWorkflowCliTests.test_entity_merge_pack_cli_restores_full_todo_order -v
```

Expected: PASS.

- [ ] **Step 6: Run the full new pack CLI test class**

Run:

```bash
python3 -m unittest tests.test_entity_workflow.EntityPackWorkflowCliTests -v
```

Expected: PASS.

- [ ] **Step 7: Commit**

Run:

```bash
git add phraseloom/entity_workflow.py phraseloom/cli.py tests/test_entity_workflow.py
git commit -m "feat: merge compact entity pack"
```

### Task 5: Update Help, README, And Run Full Verification

**Files:**
- Modify: `phraseloom/cli.py`
- Modify: `README.md`
- Modify: `tests/test_entity_workflow.py`

- [ ] **Step 1: Write the failing help test**

Add this test to `EntityPackWorkflowCliTests`:

```python
    def test_top_level_help_lists_simplified_entity_commands(self):
        from contextlib import redirect_stdout
        from io import StringIO

        from phraseloom.cli import _dispatch

        stream = StringIO()
        with redirect_stdout(stream):
            self.assertEqual(_dispatch(["--help"]), 0)

        help_text = stream.getvalue()
        self.assertIn("phraseloom entity-tm TM_REUSABLE_UNITS.xlsx", help_text)
        self.assertIn("phraseloom entity-prepare TRANSLATOR_WORKBOOK.xlsx", help_text)
        self.assertIn("phraseloom entity-fill-pack ENTITY_PACK.xlsx", help_text)
        self.assertIn("phraseloom entity-merge-pack FILLED_ENTITY_PACK.xlsx", help_text)
```

- [ ] **Step 2: Run the help test to verify RED**

Run:

```bash
python3 -m unittest tests.test_entity_workflow.EntityPackWorkflowCliTests.test_top_level_help_lists_simplified_entity_commands -v
```

Expected: FAIL because top-level help does not list the new simplified entity commands yet.

- [ ] **Step 3: Update top-level help**

In `_print_top_level_help()` in `phraseloom/cli.py`, keep the current template commands and add a simplified entity section before the low-level entity commands:

```python
    print("Entity workflow:")
    print("  phraseloom entity-tm TM_REUSABLE_UNITS.xlsx [options]")
    print("  phraseloom entity-prepare TRANSLATOR_WORKBOOK.xlsx [options]")
    print("  phraseloom entity-fill-pack ENTITY_PACK.xlsx [options]")
    print("  phraseloom entity-merge-pack FILLED_ENTITY_PACK.xlsx [options]")
    print()
    print("Advanced entity commands:")
    print("  phraseloom entity-split TRANSLATOR_WORKBOOK.xlsx [options]")
    print("  phraseloom entity-extract-tm TM_REUSABLE_UNITS.xlsx [options]")
    print("  phraseloom entity-prefill ENTITY.xlsx --tm ENTITY_TM.xlsx [options]")
    print("  phraseloom entity-fill ENTITY.xlsx [options]")
    print("  phraseloom entity-merge --entity ENTITY.xlsx --non-entity NON_ENTITY.xlsx [options]")
```

Make sure the old low-level lines are not duplicated in the generic `Commands:` section.

- [ ] **Step 4: Run the help test to verify GREEN**

Run:

```bash
python3 -m unittest tests.test_entity_workflow.EntityPackWorkflowCliTests.test_top_level_help_lists_simplified_entity_commands -v
```

Expected: PASS.

- [ ] **Step 5: Update README entity section**

In `README.md`, replace the current "独立 Entity Engine" command walkthrough with the simplified workflow first:

````markdown
## 独立 Entity Engine

Entity engine 是 translator todo 的二次处理工具，不属于主 tag/template 流水线。推荐使用四步简化流程：

```text
TM_reusable_units.xlsx
-> entity-tm
   -> TM_entity_memory.xlsx

source_translator_todo.xlsx + 可选 TM_entity_memory.xlsx
-> entity-prepare
   -> source_entity_pack.xlsx

人工确认 source_entity_pack.xlsx 里的 entity_structures / entity_terms
-> entity-fill-pack
   -> source_entity_pack_filled.xlsx

source_entity_pack_filled.xlsx
-> entity-merge-pack
   -> source_merged_todo.xlsx
```

从预处理过的 TM workbook 中创建 entity memory：

```bash
phraseloom entity-tm '/path/to/TM_reusable_units.xlsx'
```

准备本轮 source entity pack，并可选用 entity memory 预填：

```bash
phraseloom entity-prepare '/path/to/source_translator_todo.xlsx' \
  --tm '/path/to/TM_entity_memory.xlsx'
```

译员或 PM 主要处理 `source_entity_pack.xlsx` 里的可见 sheet：

```text
related_units
non_related_units
entity_structures
entity_terms
```

`_entity_map` 和 `_metadata` 是隐藏的内部 sheet，正常不需要编辑。

把 ready 的实体结构和实体词表组合回 `related_units.target_unit`：

```bash
phraseloom entity-fill-pack '/path/to/source_entity_pack.xlsx'
```

把 `related_units` 和 `non_related_units` 合并回完整 translator todo：

```bash
phraseloom entity-merge-pack '/path/to/source_entity_pack_filled.xlsx'
```

最后继续使用现有 `fill` 命令，把 merged todo 回填到目标文件。
````

Keep a short "高级/调试命令" paragraph below it listing the existing low-level commands:

```markdown
高级或调试时仍可使用底层命令：`entity-split`、`entity-extract-tm`、`entity-prefill`、`entity-fill`、`entity-merge`。
```

- [ ] **Step 6: Run focused entity tests**

Run:

```bash
python3 -m unittest tests.test_entity_workflow -v
```

Expected: PASS.

- [ ] **Step 7: Run full project tests**

Run:

```bash
python3 -m unittest discover -v
```

Expected: PASS.

- [ ] **Step 8: Commit**

Run:

```bash
git add phraseloom/cli.py README.md tests/test_entity_workflow.py
git commit -m "docs: explain compact entity workflow"
```

## Self-Review

Spec coverage:

- Task 1 covers `entity-tm` and the TM entity memory workbook default.
- Task 2 covers `entity-prepare`, one source entity pack, visible `related_units`, `non_related_units`, `entity_structures`, `entity_terms`, hidden `_entity_map`, hidden `_metadata`, and optional TM prefill.
- Task 3 covers `entity-fill-pack`, safe new-file default output, `--in-place`, and writing generated targets to `related_units`.
- Task 4 covers `entity-merge-pack`, normal merged `to_translate` output, preserved order, and non-related translations.
- Task 5 covers help text, README UX docs, and full verification.

Placeholder scan:

- No incomplete steps remain. Every implementation step names concrete files, functions, code, commands, and expected outcomes.

Type consistency:

- Public API names are `extract_entity_memory_workbook`, `prepare_entity_pack_workbook`, `fill_entity_pack_workbook`, and `merge_entity_pack_workbook`.
- Default helper names are `default_entity_memory_output_path`, `default_entity_pack_output_path`, `default_entity_filled_pack_output_path`, and `default_entity_merged_todo_output_path`.
- Sheet constants are `RELATED_UNITS_SHEET`, `NON_RELATED_UNITS_SHEET`, and `ENTITY_MAP_SHEET`.
