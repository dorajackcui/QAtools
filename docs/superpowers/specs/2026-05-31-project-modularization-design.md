# Project Modularization Design

## Inline Summary

Refactor Toolshub in a low-risk phase so the project stays modular, iterable, and easier to extend. This phase extracts shared infrastructure, fixes two optional-history GUI edge cases, and archives historical agent planning documents. It must not change existing workbook formats, CLI flags, tool behavior, or sheet names.

## Goals

- Reduce repeated Excel history TB detection code.
- Reduce repeated Tkinter GUI plumbing.
- Reduce repeated Codex subprocess plumbing.
- Keep each tool's domain logic in its own package.
- Preserve all public CLI entry points and backward-compatible wrappers.
- Keep existing tests passing and add regression tests for moved behavior.
- Move historical Superpowers plan/spec artifacts out of the active docs flow while retaining them in the repo.

## Non-Goals

- Do not rewrite the core term-pair, glossary, tag, or LLM extraction algorithms.
- Do not change output workbook sheet names or column names.
- Do not change existing CLI argument names.
- Do not change prompt content except through future tool-specific work.
- Do not introduce a new GUI framework.
- Do not delete archived planning context.

## Architecture

Add small shared modules under `tools/`:

- `tools/history_tb.py`: common history TB workbook helpers.
- `tools/gui_common.py`: small GUI helper functions/classes for file selection, integer parsing, sheet loading, and column auto-detection.
- `tools/codex_runner.py`: common Codex subprocess execution with output-last-message handling.

The tools keep their current package layout. Shared modules expose narrow APIs; individual tools remain responsible for their business rules and workbook output.

## History TB Module

`tools/history_tb.py` should own:

- workbook path resolution
- default sheet selection, preferring `术语表`
- header normalization
- source/target header detection
- two-column fallback
- same-column rejection
- row loading with configurable `start_row`

It should support both consumers:

- Term pair checker needs stripped-mark `RecordedTermPair` construction to remain local to `tools/term_pair_checker`.
- LLM term extractor needs simple `dict[source_key, target]` mapping to remain local to `tools/llm_term_extractor`.

The common module should therefore expose column/sheet detection primitives and row iteration, not tool-specific term-pair objects.

## GUI Common Module

`tools/gui_common.py` should keep the API deliberately small:

- `parse_positive_int(value, field_name, default)`
- `choose_open_file(...)`
- `choose_save_file(...)`
- `set_sheet_choices(...)`
- `detect_and_set_source_target_columns(...)`

Existing GUI classes can adopt these helpers gradually. This phase should update only the history TB optional parsing bug and the highest-duplication sheet/column paths where doing so is low risk.

## Codex Runner Module

`tools/codex_runner.py` should own:

- command construction for local `codex exec`
- `--ask-for-approval never`
- `model_reasoning_effort`
- `--output-last-message`
- temporary output file handling
- timeout and non-zero exit error messages

Tool-specific code should still own:

- prompt templates
- schema text
- JSON parsing into domain dataclasses
- retry policy details if tool-specific

## Bug Fixes

Fix optional history TB validation in:

- `tools/term_pair_checker/extract_terms_gui.py`
- `tools/workflow/workflow_gui.py`

If no history TB file is selected, those GUIs must not reject an invalid history start-row field and must pass `None` history sheet/source/target values into the processor.

## Documentation Archive

Move historical implementation artifacts from:

- `docs/superpowers/plans/`
- `docs/superpowers/specs/`

to:

- `docs/archive/superpowers/plans/`
- `docs/archive/superpowers/specs/`

The files should remain tracked. The active docs should keep user-facing guidance in `README.md`, `docs/cli-usage.md`, and per-tool README files.

## Testing

Add or update tests for:

- shared history TB detection primitives
- term pair checker history TB behavior after adopting shared helpers
- LLM term extractor history TB behavior after adopting shared helpers
- GUI optional-history start-row behavior in term pair checker and workflow
- Codex runner command construction and output reading

Run full regression:

```bash
python3 -m unittest discover -s tests
python3 -m compileall -q tools tests
git diff --check
```

## Migration Strategy

This is a behavior-preserving refactor. Migrate one shared layer at a time:

1. Add tests around existing behavior.
2. Introduce the shared module.
3. Move one consumer to the shared module.
4. Run focused and full tests.
5. Commit.

## Success Criteria

- All existing tests pass.
- New shared modules are covered by tests.
- No user-facing CLI or workbook output changes.
- Optional history TB GUI bug is fixed in both affected GUIs.
- Historical plan/spec files are archived.
- Code paths for history TB and Codex subprocess execution have less duplication.

## Self-Review

- Placeholder scan: no TBD/TODO placeholders remain.
- Scope check: this spec is one implementation phase and intentionally avoids large algorithm rewrites.
- Ambiguity check: public behavior preservation is explicit; shared modules are narrow and not a framework rewrite.
