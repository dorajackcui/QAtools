# Entity CLI UX Design

## Purpose

This change simplifies the entity engine from a sequence of low-level workbook
commands into a four-step user workflow that mirrors the main PhraseLoom
template workflow.

The existing entity engine capabilities stay intact. The UX layer reduces the
number of intermediate workbooks a PM or translator needs to understand, makes
human-edited sheets obvious, and hides implementation-only mapping sheets.

## User Workflow

The workflow has four user-facing operations:

```text
1. Build entity memory from TM
2. Prepare a source entity pack
3. Fill related units from completed entity tables
4. Merge related and non-related units back into a full todo
```

These operations map to the existing five low-level commands:

```text
entity-extract-tm -> build entity memory
entity-split + optional entity-prefill -> prepare source entity pack
entity-fill -> fill related units
entity-merge -> merge full todo
```

## Workbook Model

The simplified UX uses two main intermediate workbooks.

### TM Entity Memory Workbook

Default output:

```text
<tm_stem>_entity_memory.xlsx
```

Default location:

```text
<tm_stem>_l10n/
```

Visible sheets:

```text
entity_structures
entity_terms
```

This workbook is reusable. It is generated from a preprocessed
`*_reusable_units.xlsx` workbook and can be passed into source pack preparation
to prefill matching structures and terms.

### Source Entity Pack Workbook

Default output:

```text
<todo_stem>_entity_pack.xlsx
```

Default location:

```text
<todo_stem>_l10n/
```

Visible sheets:

```text
related_units
non_related_units
entity_structures
entity_terms
```

Hidden sheets:

```text
_entity_map
_metadata
```

`related_units` contains the todo rows that belong to discovered entity
structures. `non_related_units` contains the remaining todo rows. The two entity
tables are the primary human editing surface.

`_entity_map` contains merge and reconstruction details such as original row
order, structure IDs, extracted source entities, preview targets, fill status,
and warnings. Users should not need to edit this sheet during normal work.

## Sheet Responsibilities

`related_units`:

```text
unit_id
unit_type
source_unit
target_unit
sample_sources
context
row_number
coverage_count
variables
warning
translator_note
```

This sheet is the entity-related subset of the normal `to_translate` sheet.
Step 3 writes generated `target_unit` values here when the corresponding
structure and terms are ready.

`non_related_units`:

```text
unit_id
unit_type
source_unit
target_unit
sample_sources
context
row_number
coverage_count
variables
warning
translator_note
```

This sheet is the non-entity subset of the normal `to_translate` sheet. Users
can translate these rows normally while entity rows are handled through the
entity tables.

`entity_structures`:

```text
structure_id
source_structure
target_structure
coverage_count
confidence
risk
status
sample_sources
row_numbers
warning
```

Users fill `target_structure` and set `status` to `ready` when a structure is
approved. Only `ready` structures can fill related units.

`entity_terms`:

```text
term_id
source_entity
target_entity
occurrence_count
structure_ids
status
warning
```

Users fill `target_entity` and set `status` to `ready` when a term is approved.
Terms prefilled from entity memory may be marked ready automatically when the
match is unambiguous.

`_entity_map`:

```text
original_index
unit_id
unit_type
source_unit
structure_id
entities_json
preview_target
fill_status
warning
```

This hidden sheet powers step 3 and step 4. It preserves ordering and stores the
per-row entity mapping needed to reconstruct full target strings.

## CLI Shape

The high-level commands should be easy to explain in user docs:

```bash
phraseloom entity-tm TM_reusable_units.xlsx

phraseloom entity-prepare source_translator_todo.xlsx \
  --tm TM_entity_memory.xlsx

phraseloom entity-fill-pack source_entity_pack.xlsx

phraseloom entity-merge-pack source_entity_pack_filled.xlsx
```

`entity-prepare` accepts `--tm` as optional. When provided, it creates the pack
and immediately prefills `entity_structures` and `entity_terms`.

`entity-fill-pack` defaults to a new output workbook:

```text
<pack_stem>_filled.xlsx
```

It may support `--in-place` for advanced users, but safe new-file output is the
default.

`entity-merge-pack` outputs a normal PhraseLoom translator todo workbook:

```text
<pack_stem>_merged_todo.xlsx
```

That merged todo continues into the existing `fill` command.

## Compatibility

The current low-level commands remain available:

```text
entity-extract-tm
entity-split
entity-prefill
entity-fill
entity-merge
```

They are useful for debugging, regression tests, and advanced workflows. The new
commands are wrappers around the existing entity workflow behavior plus a more
compact workbook format.

Existing entity workbook sheet names stay readable where possible. The new pack
reader should accept both the new sheet names and the old low-level entity
workbook shape during the transition:

```text
related_units or to_translate
_entity_map or entity_source_map
```

## Error Handling

The high-level commands should raise user-facing `PhraseLoomError` subclasses
for common workflow mistakes:

```text
missing entity_structures sheet
missing entity_terms sheet
missing related_units sheet
missing non_related_units sheet
missing _entity_map sheet
duplicate original_index
missing original_index values during merge
unit_id or source_unit changed between mapping and visible rows
```

`entity-fill-pack` should leave blocked rows unfilled and write precise
`fill_status` and `warning` values rather than stopping the whole workbook when
one row is incomplete.

## Testing

Tests should cover the four user-facing operations:

```text
entity-tm writes one memory workbook with entity_structures and entity_terms
entity-prepare writes one source pack with related_units, non_related_units,
  entity_structures, entity_terms, hidden _entity_map, and optional TM prefills
entity-fill-pack writes generated targets to related_units in a new workbook
entity-merge-pack restores the original todo order and preserves non-related
  translations
```

Compatibility tests should confirm existing low-level commands still pass.

## Out Of Scope

This UX pass does not change entity extraction quality, add automatic
translation, call the tag/template engines from entity code, or write the final
delivery workbook directly. Final delivery still goes through the existing
`fill` command.
