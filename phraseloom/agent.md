# PhraseLoom Agent Entry

Read this file before changing the project.

## Product

PhraseLoom is a deterministic Excel Strings preprocessing and restore tool.
There are only two business operations:

1. `export`: clean untranslated Source rows into one self-contained Strings
   workbook.
2. `restore`: write translated Strings back into the embedded original workbook.

The product does not contain TM, translation prefill, terminology extraction,
Entity workflows, or generated translations.

## Current Flow

Export performs this sequence:

1. Read the first worksheet and ignore empty Source rows.
2. Protect configured Tags and complete `{...}` placeholders.
3. Move rows with an existing Target to the visible `completed` sheet.
4. Auto-complete numeric-only, symbol-only, and protected-only rows.
5. Deduplicate identical pending strings.
6. Merge eligible numeric/color/sequence variants into templates.
7. Optionally group structurally similar cleaned units for display ordering.
8. Embed the original workbook, row mapping, and Tag rules.

Restore expands row-specific template values, restores raw Tags/placeholders,
writes every mapped Target, preserves existing Targets, and restores original
sheet visibility and formatting.

## Modules

- `phraseloom/strings_workflow.py`: export/restore business orchestration.
- `phraseloom/strings_package.py`: Strings workbook format, metadata, and styles.
- `phraseloom/workbook_io.py`: source-row, header, column, and metadata reads.
- `phraseloom/cleaning.py`: exact deduplication and numeric template compression.
- `phraseloom/string_cluster.py`: optional similarity grouping only.
- `phraseloom/tag_engine.py`: protected-token extraction, validation, restoration.
- `phraseloom/tag_rules.py` and `tag_rules.toml`: configurable protected spans.
- `phraseloom/template_engine.py`: numeric/color/sequence template parsing and
  row-value expansion.
- `phraseloom/cli.py`, `interactive.py`, `gui.py`: the three user interfaces.
- `phraseloom/workbook_schema.py`: current Strings workbook constants only.

Keep these boundaries narrow. Do not add translation inference, memory, or
multi-stage workbook workflows to the Strings path.

## Workbook Contract

Visible sheets:

- `strings`: rows needing translation.
- `completed`: existing Targets and automatic passthrough rows.

Hidden workflow sheets:

- `_strings_map`: original row and variable mapping.
- `_metadata`: workbook kind, original state, columns, grouping, and Tag rules.
- all original worksheets during translation.

Protected tokens use `{N>`, `<N}`, and `{N}`. Raw `{0}` is source syntax and is
converted to a protected token before numeric template parsing.

## Commands

```powershell
qatools phraseloom gui
qatools phraseloom export source.xlsx
qatools phraseloom export source.xlsx --group-similar
qatools phraseloom restore source_strings.xlsx
python -m unittest discover -s tests/phraseloom_tests -v
python -m unittest discover -s tests/phraseloom_tests -p test_strings_workflow_e2e.py -v
```

On this machine, if the project virtual environment is stale, use the bundled
workspace Python reported by Codex.

## Engineering Rules

- Preserve the current Strings workbook schema unless an explicit migration is
  requested.
- Always close `openpyxl` workbooks; Windows otherwise keeps files locked.
- Keep `tag_engine.py`, `template_engine.py`, and `string_cluster.py` free of
  Excel I/O.
- Similarity grouping may change group/order only; it must never merge ordinary
  strings or create Targets.
- Run the focused end-to-end test and the full suite after workflow changes.
- `testfiles/` is ignored local data; do not commit generated workbooks.
