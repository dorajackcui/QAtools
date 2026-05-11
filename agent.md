# PhraseLoom Agent Entry

This file is the fast onboarding note for a new Codex/agent session in this
repository. Read it before asking the user to re-explain the project.

## Project Purpose

PhraseLoom is an internal Excel localization workflow tool. It reduces a source
Excel file into reusable translation units, uses historical TM data to prefill
matching units, creates a translator-facing todo workbook, and fills translated
units back into a target column or report workbook.

Main workflow:

1. Completed historical Excel -> reusable TM workbook.
2. New source Excel + reusable TM workbook -> TM prefill pack + translator todo.
3. Translator fills `target_unit` in the todo workbook.
4. Todo workbook -> filled source workbook copy or report workbook.

Protected-token extraction is integrated into the main workflow as an internal
pre-template layer. It serializes recognized tags and every complete raw `{...}`
placeholder into translator-facing protected tokens such as `{1>`, `<2}`, and
`{3}`, runs the normal template/TM flow on that serialized text, then restores
raw spans during fill. Entity clustering exists as an experimental side module
and is not integrated into the main localization workflow yet.

## First Files To Read

- `README.md`: user-facing workflow, commands, and output file meanings.
- `phraseloom/workflow.py`: business orchestration for extract, TM extraction,
  and fill.
- `phraseloom/excel_io.py`: Excel reading/writing, default output paths, schema
  metadata, and workbook lifecycle handling.
- `phraseloom/template_engine.py`: pure template parsing/inference/application.
- `phraseloom/tag_engine.py`: pure tag serialization, validation, tag-only
  detection, and raw tag restoration.
- `phraseloom/workbook_schema.py`: centralized sheet names, column names, and
  schema version constants.
- `phraseloom/errors.py`: user-facing structured exceptions.
- `phraseloom/cli.py`: argparse CLI and command dispatch.
- `phraseloom/interactive.py`: interactive three-step prompt flow.
- `phraseloom/entity_cluster.py`: public facade for the experimental entity
  cluster probe.
- `tests/test_template_workflow.py`: main regression tests for workflow, CLI,
  tag integration, compatibility shims, schema errors, and default paths.
- `tests/test_tag_engine.py`: focused pure tag extraction and restoration tests.
- `tests/test_tag_workflow_testfiles.py`: tag workflow tests that create
  isolated fixture workbooks under ignored `testfiles/` temp directories.
- `tests/test_entity_cluster.py`: entity cluster regression tests.
- `docs/superpowers/specs/2026-05-11-phraseloom-internal-tool-architecture-design.md`:
  architecture design context.
- `docs/superpowers/plans/2026-05-11-phraseloom-internal-tool-architecture.md`:
  implementation plan that shaped the current package structure.
- `docs/superpowers/specs/2026-05-11-tag-extractor-design.md`: tag extractor
  design, placeholder contract, and workflow integration.
- `docs/superpowers/plans/2026-05-11-tag-extractor.md`: implementation plan for
  the tag extraction integration.

## Package Shape

Use the package modules for new code:

- `phraseloom.workflow`: public workflow API.
- `phraseloom.tag_engine`: pure tag extraction, validation, and restoration.
- `phraseloom.template_engine`: pure string/template logic.
- `phraseloom.excel_io`: Excel I/O and default path helpers.
- `phraseloom.workbook_schema`: workbook contract constants.
- `phraseloom.cli`: console entry point.
- `phraseloom.interactive`: prompt-based workflow.
- `phraseloom.entity_cluster`: experimental entity clustering API.

Top-level `template_demo.py` and `entity_cluster_probe.py` are compatibility
shims. Keep them working unless the user explicitly asks to remove legacy entry
points.

## Current Default Output Names

Default outputs are created beside the input under `<input_stem>_l10n/`.

- TM extraction: `<stem>_reusable_units.xlsx`
- TM prefill process pack: `<stem>_tm_prefill_pack.xlsx`
- Translator todo workbook: `<stem>_translator_todo.xlsx`
- Filled delivery workbook: `<stem>_filled_result.xlsx`
- Legacy result workbook: `<stem>_phraseloom_result.xlsx`

Do not change workbook sheet names casually; existing workbooks must remain
readable. Sheet names such as `tm_pairs`, `to_translate`, `translation_units`,
`prefilled_units`, `source_map`, and `filled_workbook` are part of the workbook
compatibility contract.

## Commands

On this Windows machine, Python may not be visible inside the default sandbox.
If `python` or `py -3` fails in the sandbox, use an escalated Windows command
for Python work. The user has explicitly allowed this pattern for Python-related
project testing.

Run tests:

```powershell
py -3 -m unittest discover -v
```

Run the three-step workflow against local sample files:

```powershell
py -3 -m phraseloom.cli tm-extract testfiles\TM.xlsx --source-col source --target-col target

py -3 -m phraseloom.cli extract testfiles\for_test.xlsx --source-col source --target-col target --tm testfiles\TM_l10n\TM_reusable_units.xlsx --no-existing-targets

py -3 -m phraseloom.cli fill testfiles\for_test.xlsx --templates testfiles\for_test_l10n\for_test_translator_todo.xlsx --source-col source --target-col target --mode target-column
```

Known sample stats from `testfiles/`:

- `TM.xlsx`: 20,532 source/target rows.
- `for_test.xlsx`: 45,045 source rows and initially empty target rows.
- TM extraction currently produces 19,402 reusable units.
- Extract with TM prefill currently fills 1,333 units covering 4,571 rows.

## Important Engineering Notes

- Keep `template_engine.py` pure: no Excel, CLI, or filesystem logic there.
- Keep `tag_engine.py` pure: no Excel, CLI, or filesystem logic there.
- Protected tokens are reserved as `{N>`, `<N}`, and `{N}`. The protected-token
  extractor owns those tokens; template parsing must preserve them and must not
  include their numbers in normal template variables.
- Protected-only units should auto-fill with `target_unit_source = "tag_only"`
  for workbook compatibility.
- Fill order matters: apply template variables first, validate tag placeholders,
  then restore known raw tags. Tag mismatch warnings do not block writing the
  target; downstream QA can do stricter checks.
- Keep orchestration in `workflow.py`; keep Excel serialization in `excel_io.py`.
- Use `workbook_schema.py` constants instead of hardcoding sheet/column names in
  new I/O code.
- Raise `PhraseLoomError` subclasses for user-facing failures. CLI should catch
  them and print concise actionable messages.
- Always close `openpyxl` workbooks after loading. Windows locks `.xlsx` files
  until `wb.close()` is called.
- The default test command should discover 48 tests at the time of this note.
- `testfiles/` contains local sample workbooks and generated outputs. Treat it as
  local test data; do not commit large generated workbook outputs.
- `.worktrees/` is ignored for local git worktrees. Do not commit generated
  worktree contents.

## Recent Useful Commits

- `60d2d1d Normalize tagged example units`
- `bc799d1 Isolate testfiles tag fixtures`
- `71fd41e Restore tags after template fill`
- `221b5e9 Serialize row tags before template parsing`
- `d5aec33 Add pure tag extraction engine`

## Good Next Steps

Likely future work:

- Integrate entity cluster discovery into the main workflow as an optional step.
- Add a stricter optional tag QA/reporting pass for ordering and malformed
  placeholder structure if needed.
- Extend `tag_engine.py` with additional conservative tag rules as real samples
  reveal them.
- Improve README encoding/content if the displayed Chinese text appears garbled.
- Add more real-sample tests once the desired golden outputs are stable.
