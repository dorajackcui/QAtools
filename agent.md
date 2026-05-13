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
pre-template layer. It serializes allowed tags and every complete raw `{...}`
placeholder into translator-facing protected tokens such as `{1>`, `<2}`, and
`{3}`, runs the normal template/TM flow on that serialized text, then restores
raw spans during fill. Tag extraction is governed by `phraseloom/tag_rules.toml`.
The default allowlist protects formatting tags such as `color`, `size`, `img`,
`br`, `i`, `u`, `outline`, and `c`; unknown angle-bracket labels such as
`<Activate>` stay translatable text.

Entity engine exists as an independent second-pass workflow over already
preprocessed todo/TM workbooks. It does not call `tag_engine` or
`template_engine`, does not read the original source workbook, and does not
write the final delivery workbook directly. It splits a translator todo into
entity-related and non-entity workbooks, extracts entity structures/terms from a
preprocessed `tm_pairs` workbook, prefills entity structures/terms, fills ready
entity rows back into the entity-related todo, and merges the two todo branches
back into a complete translator todo. The merged todo then goes through the
existing `fill` command.

## First Files To Read

- `README.md`: user-facing workflow, commands, and output file meanings.
- `phraseloom/workflow.py`: business orchestration for extract, TM extraction,
  and fill.
- `phraseloom/excel_io.py`: Excel reading/writing, default output paths, schema
  metadata, and workbook lifecycle handling.
- `phraseloom/template_engine.py`: pure template parsing/inference/application.
- `phraseloom/tag_engine.py`: pure protected-token serialization, validation,
  protected-only detection, and raw span restoration.
- `phraseloom/workbook_schema.py`: centralized sheet names, column names, and
  schema version constants.
- `phraseloom/errors.py`: user-facing structured exceptions.
- `phraseloom/cli.py`: argparse CLI and command dispatch.
- `phraseloom/interactive.py`: interactive three-step prompt flow.
- `phraseloom/entity_workflow.py`: independent entity workflow for splitting
  preprocessed todo workbooks, extracting entity TM, prefill/fill, and merge.
- `phraseloom/entity_cluster.py`: public facade for the experimental entity
  cluster probe used by the first entity extraction strategy.
- `tests/test_template_workflow.py`: main regression tests for workflow, CLI,
  tag integration, compatibility shims, schema errors, and default paths.
- `tests/test_tag_engine.py`: focused pure tag extraction and restoration tests.
- `tests/test_tag_workflow_testfiles.py`: tag workflow tests that create
  isolated fixture workbooks under ignored `testfiles/` temp directories.
- `tests/test_entity_workflow.py`: entity workflow regression tests for split,
  TM extraction, prefill, fill, merge, and CLI dispatch.
- `tests/test_entity_cluster.py`: entity cluster regression tests.
- `docs/entity-engine-flow.html`: visual overview of the independent entity
  engine workflow.
- `docs/superpowers/specs/2026-05-12-entity-engine-design.md`: current entity
  engine design and boundaries.
- `docs/superpowers/plans/2026-05-12-entity-engine-workflow.md`: implementation
  plan for the first entity workflow skeleton.
- `docs/superpowers/specs/2026-05-11-phraseloom-internal-tool-architecture-design.md`:
  architecture design context.
- `docs/superpowers/plans/2026-05-11-phraseloom-internal-tool-architecture.md`:
  implementation plan that shaped the current package structure.
- `docs/superpowers/specs/2026-05-11-protected-token-design.md`: current
  protected-token contract and workflow integration.
- `docs/superpowers/plans/2026-05-11-protected-token.md`: implementation plan
  for the protected-token contract.
- `docs/superpowers/specs/2026-05-11-tag-extractor-design.md`: historical tag
  extractor design that the protected-token contract superseded.

## Package Shape

Use the package modules for new code:

- `phraseloom.workflow`: public workflow API.
- `phraseloom.tag_engine`: pure tag extraction, validation, and restoration.
- `phraseloom.template_engine`: pure string/template logic.
- `phraseloom.excel_io`: Excel I/O and default path helpers.
- `phraseloom.workbook_schema`: workbook contract constants.
- `phraseloom.cli`: console entry point.
- `phraseloom.interactive`: prompt-based workflow.
- `phraseloom.entity_workflow`: independent entity workflow API over
  preprocessed todo/TM workbooks.
- `phraseloom.entity_cluster`: experimental entity clustering API and
  compatibility facade for the cluster probe strategy.

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
- Entity split outputs: `<todo_stem>_entity_related.xlsx` and
  `<todo_stem>_not_entity_related.xlsx`
- Entity TM output: `<tm_stem>_entity_tm.xlsx`
- Entity prefill output: `<entity_stem>_prefilled.xlsx`
- Entity fill output: `<entity_stem>_filled.xlsx`
- Entity merge output: `<entity_stem>_merged_todo.xlsx`

Do not change workbook sheet names casually; existing workbooks must remain
readable. Sheet names such as `tm_pairs`, `to_translate`, `translation_units`,
`prefilled_units`, `source_map`, `filled_workbook`, `entity_structures`,
`entity_terms`, and `entity_source_map` are part of the workbook compatibility
contract.

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

Run the independent entity workflow against local sample files:

```powershell
py -3 -m phraseloom.cli entity-split testfiles\entity_test\target.xlsx --entity-output testfiles\entity_test\entity_run\target_entity_related.xlsx --non-entity-output testfiles\entity_test\entity_run\target_not_entity_related.xlsx

py -3 -m phraseloom.cli entity-extract-tm testfiles\entity_test\TM_reusable_units.xlsx -o testfiles\entity_test\entity_run\TM_entity_tm.xlsx

py -3 -m phraseloom.cli entity-prefill testfiles\entity_test\entity_run\target_entity_related.xlsx --tm testfiles\entity_test\entity_run\TM_entity_tm.xlsx -o testfiles\entity_test\entity_run\target_entity_prefilled.xlsx

py -3 -m phraseloom.cli entity-fill testfiles\entity_test\entity_run\target_entity_prefilled.xlsx -o testfiles\entity_test\entity_run\target_entity_filled.xlsx

py -3 -m phraseloom.cli entity-merge --entity testfiles\entity_test\entity_run\target_entity_filled.xlsx --non-entity testfiles\entity_test\entity_run\target_not_entity_related.xlsx -o testfiles\entity_test\entity_run\target_merged_todo.xlsx
```

Known sample stats from `testfiles/`:

- `TM.xlsx`: 20,532 source/target rows.
- `for_test.xlsx`: 45,045 source rows and initially empty target rows.
- TM extraction currently produces 19,402 reusable units.
- Extract with TM prefill currently fills 1,333 units covering 4,571 rows.
- `testfiles/entity_test/target.xlsx` is a preprocessed translator todo
  workbook used for entity workflow testing.
- `testfiles/entity_test/TM_reusable_units.xlsx` is a preprocessed `tm_pairs`
  workbook used for entity TM extraction testing.
- Recent entity sample run: target split produced 1,621 entity units, 16,598
  non-entity units, 144 entity structures, and 911 entity terms; entity TM
  extraction produced 121 structures and 587 terms; prefill filled 4 structures
  and 31 terms; entity fill produced 0 ready rows until manual review marks
  structures/terms as `ready`.

## Important Engineering Notes

- Keep `template_engine.py` pure: no Excel, CLI, or filesystem logic there.
- Keep `tag_engine.py` pure: no Excel, CLI, or filesystem logic there.
- Protected tokens are reserved as `{N>`, `<N}`, and `{N}`. The protected-token
  extractor owns those tokens; template parsing must preserve them and must not
  include their numbers in normal template variables.
- Tag extraction is governed by `phraseloom/tag_rules.toml`. The default
  allowlist protects formatting tags such as `color`, `size`, `img`, `br`, `i`,
  `u`, `outline`, and `c`; unknown angle-bracket labels such as `<Activate>`
  stay translatable text.
- Protected-only units should auto-fill with `target_unit_source = "tag_only"`
  for workbook compatibility.
- Fill order matters: apply template variables first, validate protected tokens,
  then restore known raw spans. Protected-token mismatch warnings do not block
  writing the target; downstream QA can do stricter checks.
- Entity workflow is deliberately decoupled from the main tag/template flow.
  It consumes preprocessed `to_translate` and `tm_pairs` workbooks, writes
  entity workbooks, and finally emits a merged translator todo for the existing
  fill command.
- Entity TM prefill writes only `entity_structures.target_structure` and
  `entity_terms.target_entity`; it does not directly write full todo
  `target_unit` values.
- Entity fill writes `target_unit` only when the structure and every referenced
  entity term are translated and marked `ready`.
- Entity extraction is strategy-based. The first strategy wraps the existing
  cluster probe logic, but future strategies should plug into
  `phraseloom.entity_workflow` without reshaping split/prefill/fill/merge.
- Keep orchestration in `workflow.py`; keep Excel serialization in `excel_io.py`.
- Use `workbook_schema.py` constants instead of hardcoding sheet/column names in
  new I/O code.
- Raise `PhraseLoomError` subclasses for user-facing failures. CLI should catch
  them and print concise actionable messages.
- Always close `openpyxl` workbooks after loading. Windows locks `.xlsx` files
  until `wb.close()` is called.
- The default test count changes as regression coverage grows; rely on the
  command result rather than a hardcoded count.
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

- Simplify the visible entity workbook tables now that the first entity engine
  skeleton is in place; keep internal mapping fields hidden where possible.
- Add a stricter optional protected-token QA/reporting pass for ordering and
  malformed token structure if needed.
- Extend `tag_engine.py` with additional conservative tag rules as real samples
  reveal them.
- Improve README encoding/content if the displayed Chinese text appears garbled.
- Add more real-sample tests once the desired golden outputs are stable.
