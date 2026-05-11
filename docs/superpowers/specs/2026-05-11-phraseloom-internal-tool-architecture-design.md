# PhraseLoom Internal Tool Architecture Design

Date: 2026-05-11

## Goal

Turn PhraseLoom from a working script prototype into a stable internal Python CLI tool. The first refactor focuses on package structure, clear module boundaries, dependency declaration, testability, and compatibility with the current user workflow.

This phase does not add new localization capabilities. `entity_cluster` remains an experimental discovery module, and future steps such as tag preprocessing are reserved as extension points after the core architecture is stable.

## Current Project Shape

The repository currently contains:

- `template_demo.py`: main workflow, template parsing, Excel I/O, workbook writing, CLI commands, and interactive prompts in one file.
- `entity_cluster_probe.py`: experimental entity pattern discovery for reusable sentence structures.
- `test_template_demo.py`: workflow and CLI-adjacent tests for the current main script.
- `test_entity_cluster_probe.py`: tests for entity cluster discovery.
- `README.md`: Chinese user guide for the current three-step localization workflow.

The current commands and behavior are useful and should remain compatible:

- `tm-extract`: completed Excel workbook to reusable `tm_pairs`.
- `extract`: new source workbook plus optional `tm_pairs` to `template_pack` and `to_translate`.
- `fill`: translated `to_translate` workbook back to report workbook or target-column output copy.

## Recommended Approach

Use a gradual package refactor while preserving the current behavior. The code should move into a `phraseloom` package, but existing workflow semantics and workbook formats should keep working through regression tests.

This approach is preferred over a minimal cleanup because the main script has already crossed the size where responsibilities are easy to reason about. It is also preferred over a full rewrite because the current business logic is covered by tests and should be moved carefully.

## Target Structure

```text
phraseloom/
  __init__.py
  cli.py
  interactive.py
  models.py
  template_engine.py
  workflow.py
  excel_io.py
  workbook_schema.py
  entity_cluster.py
tests/
  test_template_workflow.py
  test_entity_cluster.py
pyproject.toml
README.md
```

### Module Responsibilities

`models.py`

Defines shared data structures such as `TemplateMatch`, `RowItem`, `TranslationUnit`, `EntityOccurrence`, and `EntityCluster`.

`template_engine.py`

Owns pure string/template behavior:

- source template parsing
- variable key naming
- target template inference
- target template application
- non-translatable segment detection if it remains purely text-based

This module must not know about Excel files or CLI arguments.

`workbook_schema.py`

Defines sheet names, column names, and workbook schema versions. This centralizes values such as `translation_units`, `to_translate`, `tm_pairs`, `target_unit`, and `source_unit` so format changes are visible and testable.

`excel_io.py`

Owns Excel-specific read/write behavior:

- source row loading
- column resolution
- workbook sheet parsing
- output workbook creation
- sheet styling
- default output path helpers if they are mostly filesystem concerns

This module should not decide which rows become templates or segments. It receives workflow results and serializes them.

`workflow.py`

Owns business orchestration:

- `extract_tm_pairs`
- `prepare_translation`
- `fill_translation`
- compatibility aliases for current public functions if needed during migration

It coordinates template parsing, unit building, TM/template loading, workbook writing, and statistics.

`cli.py`

Owns `argparse` command definitions and maps CLI arguments to `workflow.py`. It should preserve current command semantics:

- `tm-extract` and `extract-tm`
- `extract`
- `fill`
- legacy positional mode if we choose to keep it during the transition

`interactive.py`

Owns the three-step prompt flow:

1. Build TM from completed Excel
2. Prepare translator file for new source
3. Fill source from translated file

It calls `workflow.py` and contains no workbook logic.

`entity_cluster.py`

Moves the current experimental entity cluster logic out of `entity_cluster_probe.py`. It remains separate from the main workflow and keeps its own tests. Future integration can happen after the stable package boundary is in place.

## Data Flow

### TM Extraction

```text
Completed Excel
  -> excel_io.read_source_rows
  -> workflow.build_translation_units
  -> excel_io.write_tm_pairs
```

### Translation Preparation

```text
Source Excel + optional tm_pairs
  -> excel_io.read_source_rows
  -> excel_io.load_translated_units
  -> workflow.build_translation_units
  -> excel_io.write_template_pack
  -> excel_io.write_to_translate
```

### Fill

```text
Source Excel + translated to_translate
  -> excel_io.read_source_rows
  -> excel_io.load_translated_units
  -> workflow.build_fill_context
  -> excel_io.write_report or excel_io.write_target_column_copy
```

The public library API should be small and stable:

```python
from phraseloom.workflow import extract_tm_pairs, prepare_translation, fill_translation
```

## Dependency Direction

Dependencies should flow inward from interfaces to core logic:

```text
cli / interactive
  -> workflow
    -> template_engine
    -> excel_io
    -> workbook_schema
    -> models
```

`template_engine` and `models` should remain pure and fast to test. `excel_io` can depend on `openpyxl`. `cli` and `interactive` should not be imported by core modules.

## Error Handling

Introduce structured exceptions:

```text
PhraseLoomError
  ConfigError
  WorkbookFormatError
  ColumnNotFoundError
  TranslationUnitLoadError
  WorkflowError
```

Rules:

- Missing columns should report the requested column and available headers.
- Missing workbook sheets or required columns should raise a schema error instead of silently returning an empty result.
- CLI should catch `PhraseLoomError` and print concise, actionable messages.
- Library APIs should raise structured exceptions so tests and future integrations can handle them.
- The tool should continue writing outputs into `*_l10n/` by default and should not overwrite the original source workbook.

Example CLI error:

```text
Column 'fr' not found in header row.
Available columns: source, target, note
```

## Workbook Compatibility

The first refactor should read existing workbooks written by the current scripts. New workbooks should include a schema version in a summary or metadata sheet while keeping existing sheet and column names intact.

Compatibility priorities:

1. Current `tm_pairs` workbooks can prefill new extraction runs.
2. Current `to_translate` workbooks can be used by `fill`.
3. Current `template_pack` workbooks remain readable where the workflow already supports them.
4. New files keep the same translator-facing columns unless a later spec deliberately changes them.

## Extension Points

Do not build a plugin system in this phase. Define simple internal seams for later workflows:

```python
class Preprocessor:
    def process_rows(self, rows: list[RowItem]) -> list[RowItem]:
        ...

class UnitDiscovery:
    def discover(self, rows: list[RowItem]) -> list[DiscoveredPattern]:
        ...
```

Future use:

```text
read Excel
  -> optional tag preprocess
  -> template extraction
  -> optional entity cluster discovery
  -> translation units
```

This keeps the first refactor focused while making later tag and entity workflows easier to add without reshaping the package again.

## Testing Strategy

The first implementation plan should protect existing behavior before moving code:

- Move tests into `tests/`.
- Keep current scenario coverage for TM extraction, translation preparation, fill, output paths, interactive flow, and entity clustering.
- Add focused tests for `template_engine` pure functions.
- Add tests for column resolution and workbook schema errors.
- Add CLI smoke tests for command help.
- Keep entity clustering tests separate and label the module experimental.

The preferred verification command after a Python environment is available:

```bash
python -m unittest discover -v
```

If project tooling switches to pytest later, that should be a separate decision with equivalent coverage.

## Packaging And Tooling

Add `pyproject.toml` with:

- project name: `phraseloom`
- Python version constraint
- dependency on `openpyxl`
- console script entry point, for example `phraseloom = phraseloom.cli:main`
- test configuration if needed

The repository should also gain a clear development setup section in `README.md`, including how to create a virtual environment, install dependencies, and run tests.

## Migration Plan

The implementation should be incremental:

1. Add package scaffold, `pyproject.toml`, and test layout.
2. Move data models and pure template functions first.
3. Move Excel reading/writing and workbook schema constants.
4. Move workflow orchestration while preserving existing public function names through wrappers if needed.
5. Move CLI and interactive prompt logic.
6. Move entity cluster logic into `phraseloom.entity_cluster`.
7. Update tests and README after each migration step.
8. Keep script-level compatibility shims until tests and documentation confirm the package entry point is ready.

## Out Of Scope For This Phase

- A GUI, web app, or desktop app.
- New tag preprocessing behavior.
- Integrating entity clusters into the main translation workflow.
- Replacing `openpyxl`.
- Changing translator-facing workbook columns beyond adding non-breaking metadata.
- Publishing to a package index.

## Acceptance Criteria

- The project has a `phraseloom` package with clear module responsibilities.
- Current core commands keep working through the new CLI entry point.
- Existing tests are migrated and passing in the configured Python environment.
- Workbook output remains compatible with the current documented workflow.
- README explains setup, CLI usage, and test execution.
- Experimental entity cluster logic is isolated from the main workflow.
