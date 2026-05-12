# Entity Engine Design

## Purpose

The entity engine is an independent second-pass workflow for PhraseLoom todo
workbooks. It finds reusable entity-bearing structures inside already
preprocessed translation units, lets those structures and their entity terms be
translated or prefilled from an entity TM, and then merges completed entity
results back into a full target todo workbook.

The entity engine is not part of the tag or template pipeline. It does not read
the original source workbook, call `tag_engine`, call `template_engine`, or
write the final delivery workbook.

## Inputs And Outputs

Target input:

```text
target todo workbook
```

The workbook is assumed to have already been produced by the existing
tag/template workflow. The entity engine reads the todo workbook's
`to_translate.source_unit` values as its source text surface.

Entity split outputs:

```text
entity_related_workbook
not_entity_related_workbook
```

The entity-related workbook contains only units that belong to discovered
entity structures. The non-entity workbook contains the remaining todo rows and
preserves hidden/supporting sheets needed by the normal PhraseLoom fill flow.
Both outputs carry stable merge keys.

Entity TM input:

```text
completed entity workbook
```

The entity TM contains reusable `source_structure -> target_structure` and
`source_entity -> target_entity` rows. It may be generated from previous entity
workbooks or prepared manually after entity extraction.

Final output:

```text
merged target todo workbook
```

The merged workbook is a normal PhraseLoom todo workbook. It is passed to the
existing `fill` command to write the final target workbook.

## Workflow

```text
target todo workbook
-> entity-split
   -> entity_related_workbook
   -> not_entity_related_workbook

entity TM workbook
-> reusable entity_structures + entity_terms

entity-prefill
entity TM structures/terms -> target entity structures/terms

entity-fill
completed target structures/terms -> filled entity_related_workbook

entity-merge
filled entity_related_workbook + not_entity_related_workbook
-> merged target todo workbook
```

`entity-prefill` and `entity-fill` are intentionally different actions.
Prefill only copies reusable structure and term translations into the entity
workbook. Fill only writes a `target_unit` when the current target row's
structure and all of its entities are complete and approved.

## Workbook Contract

The entity-related workbook contains the original subset of `to_translate`
rows, with an `original_index` column used for merging, plus these sheets:

### `entity_structures`

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

`status` values:

```text
ready
review
skip
```

Only `ready` structures can be used by `entity-fill`.

### `entity_terms`

```text
term_id
source_entity
target_entity
occurrence_count
structure_ids
status
warning
```

Only `ready` terms with a non-empty `target_entity` can be used by
`entity-fill`.

### `entity_source_map`

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

`entities_json` maps placeholder names such as `entity1` to the source entity
text found in that row.

The non-entity workbook contains the remaining `to_translate` rows with
`original_index`. It does not contain entity structures or entity terms.

## Fill Rules

For each row in `entity_source_map`, `entity-fill` writes a preview and updates
the entity-related workbook's `to_translate.target_unit` only when all of these
conditions hold:

```text
structure.status = ready
target_structure is not blank
every source entity has a matching entity_terms row
every matching entity term has status = ready
every matching entity term has target_entity
unit_id and source_unit still match the entity-related to_translate row
```

If any condition fails, the row is left unfilled and receives a precise
`fill_status` and `warning`.

## Merge Rules

`entity-merge` combines the filled entity-related workbook and the non-entity
workbook.

It restores the original todo order by `original_index`. For entity rows it
uses the `target_unit` from the filled entity-related workbook. For non-entity
rows it uses the `target_unit` from the non-entity workbook.

Merge stops without writing output if:

```text
original_index is duplicated
unit_id is missing
unit_id/source_unit disagree for the same original_index
there is a gap in the expected original_index sequence
```

## Extraction Strategies

Entity structure discovery is a strategy boundary. The workflow code depends on
an extractor interface, not on one hard-coded algorithm.

The first implementation uses the existing cluster-probe approach from
`phraseloom/_entity_cluster_probe.py`. Future strategies can add different
rules, language-specific boundary handling, glossary-assisted extraction, or
model-assisted proposals without changing split, prefill, fill, or merge.

The default cluster strategy produces structures such as:

```text
{entity1} launched an attack and dealt damage.
```

or:

```text
Equip {1>}{entity1} Outfit<2}, then switch to {entity2} in demo.
```

The workflow then extracts per-row entity values from the selected source
structure and writes them to `entity_source_map`.

## Out Of Scope

The first version does not:

```text
translate non-entity rows
create a second partB translation memory
call tag_engine or template_engine
write the final delivery workbook
infer translated target structures from arbitrary bilingual TM automatically
```

Non-entity rows are preserved for other workflows and merged back later.
