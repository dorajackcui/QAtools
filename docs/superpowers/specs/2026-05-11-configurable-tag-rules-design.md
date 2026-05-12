# Configurable Tag Rules Design

Date: 2026-05-11

## Goal

Add a configuration file that defines which tag-like spans PhraseLoom protects
as structural tokens. The goal is to stop over-extracting translatable text such
as `<Activate>` while preserving the useful normalization behavior for real
formatting tags such as `<color=#123>...</>`.

The protected-token contract remains unchanged:

```text
{N>   opening protected span
<N}   closing protected span
{N}   single protected span
```

Raw `{...}` placeholders remain protected by default.

## Root Cause

The current `tag_engine` treats every HTML-like `<Name...>` span as a structural
tag. In real TM data, some angle-bracket text is translatable label text:

```text
TM source: <Activate> HP increased by {a}%
TM target: <Active> HP increased translated...
```

Current behavior protects source `<Activate>` as `{1>` but target serialization
is source-driven and only replaces exact raw spans from source metadata. The
translated target text has a different raw span, so it remains raw. This creates
an inconsistent TM pair:

```text
source_unit: {1> HP increased by {2}%
target_unit: <Active> ... {2}%
```

The fix is not to independently extract target tags. That would classify
translated label text as structure too. The fix is to restrict structural tag
extraction to configured tag names.

## Implemented Approach

Use an allowlist configuration file.

- Only configured angle tags and BBCode tags are protected.
- Unknown tag-like text remains normal translatable text.
- Raw `{...}` placeholders stay protected.
- TM source, TM target, new source rows, and fill rows all use the same tag
  rules.
- TM source and target serialize under the same active rules, with target
  serialization driven by the source row's protected metadata.
- Matching uses the protected source unit/template only.
- Target units are protected so that matched translations can restore the
  current source row's raw tags during fill.

## Config File

The default packaged config path is:

```text
phraseloom/tag_rules.toml
```

Schema:

```toml
version = 1

[angle_tags]
mode = "allowlist"
allowed = ["color", "size", "img", "br", "i", "u", "outline", "c"]

[bbcode_tags]
mode = "allowlist"
allowed = ["color", "b", "i", "u", "size"]

[raw_braces]
protect_all = true
```

Rules:

- Tag names are case-insensitive.
- Attributes do not affect matching once the tag name is allowed.
- Shorthand closes such as `</>` and `[/]` close the nearest currently protected
  open tag.
- A named close is protected only when its tag name is allowed and it matches
  the current protected stack top.
- Excluded tag-like text does not enter the protected stack and does not produce
  unpaired-tag warnings.
- Allowed but malformed tags can still produce warnings, as they do today.

The CLI accepts an optional override:

```text
phraseloom tm-extract TM.xlsx --tag-config path/to/tag_rules.toml
phraseloom extract SOURCE.xlsx --tag-config path/to/tag_rules.toml
phraseloom fill SOURCE.xlsx --templates TRANSLATOR_FILE.xlsx --tag-config path/to/tag_rules.toml
```

Interactive mode can continue using the packaged default first; adding a prompt
for custom config can be a later UX improvement.

## TM Serialization And Matching

TM prefill must serialize both TM source and TM target before building reusable
units.

```text
TM raw source
  -> extract allowed protected spans with tag rules
  -> parse source template
  -> source_unit / source_template match key

TM raw target
  -> serialize against TM source token metadata
  -> infer target template / target_unit
```

The target does not participate in the match key. It is the reusable output
associated with a source key.

New source extraction uses the same rules:

```text
new raw source
  -> extract allowed protected spans with tag rules
  -> parse source template
  -> source_unit / source_template
```

TM prefill matches:

```text
TM source_unit or source_template
==
new source_unit or source_template
```

## Fill Behavior

When a TM target unit is applied to a new source row, restore protected tokens
using the new source row's metadata, not the TM row's raw spans.

Example:

```text
TM source: <color=#123>Source</>
TM target: <color=#123>Target</>

TM source_unit: {1>Source<2}
TM target_unit: {1>Target<2}

New source: <color=#3333>Source</>
New source_unit: {1>Source<2}

Filled serialized target: {1>Target<2}
Final target: <color=#3333>Target</>
```

This is the intended normalization: matching ignores formatting attribute
differences, while fill preserves the current file's raw formatting.

Counter-example:

```text
TM source: <Activate> HP increased by {a}%
TM target: <Active> PV increased by {a}%
```

If `activate` is not in `angle_tags.allowed`:

```text
TM source_unit: <Activate> HP increased by {1}%
TM target_unit: <Active> PV increased by {1}%
```

The angle-bracket labels remain translatable text. Only `{a}` is protected.

## Config Metadata

Generated workbooks record the tag rules used to produce protected units.

Metadata keys:

```text
tag_rules_version
tag_rules_hash
tag_rules_source
```

The hash is computed from the normalized config content, not from raw
file formatting. This makes harmless TOML formatting changes less noisy.

Compatibility behavior:

- If both workbooks have tag rule hashes and they differ, loading rejects the
  workbook with a clear configuration mismatch error.
- If a workbook has no tag rule hash, it is treated as legacy and still loads.
- New generated workbooks write the metadata hash, version, and source.

This prevents using TM pairs generated under one tag rule set to fill a source
file generated under another rule set.

## Implementation Shape

The `phraseloom.tag_rules` module provides:

- `TagRules` dataclass.
- `load_tag_rules(path: Path | None) -> TagRules`.
- `default_tag_rules() -> TagRules`.
- `normalized_tag_rules_hash(rules: TagRules) -> str`.

`tag_engine` accepts rules:

```python
extract_tags(text, rules: TagRules | None = None)
```

The default handling avoids repeatedly reading the TOML file per row. Workflow
loads once and passes the rules through row reading and TM creation.

`serialize_known_tags()` remains source-metadata-driven and does not need to
independently read rules.

Because the default config is packaged inside `phraseloom`, packaging metadata
includes the TOML file as package data.

`workflow` and `excel_io` carry the loaded rules through:

- `generate_tm_pairs(..., tag_config=None)`
- `generate_workbook(..., tag_config=None)`
- `fill_target_column_workbook(..., tag_config=None)`
- `_read_source_rows(..., tag_rules=rules)`

## Warning Behavior

Warnings are surfaced through the generated unit metadata:

- Source extraction warnings are attached to generated units.
- Target serialization warnings are attached to generated units.
- TM pair workbooks surface these warnings in `tm_pairs.warning`.
- Fill continues warning on missing or extra protected tokens.

Excluded tag-like text should not emit warnings merely because it looks like a
tag. It is intentional text under the current config.

## Tests

Focused tests cover:

- Default config loads and includes expected allowed tags.
- `<color=#123>Text</>` is protected.
- `<Activate> HP increased by {a}%` keeps `<Activate>` raw and protects only
  `{a}`.
- TM source and target are serialized with the same source metadata.
- TM color attribute normalization fills a new row with the new row's raw color
  value.
- Target is not independently scanned for translated angle labels.
- Config mismatch metadata is detected.
- Existing protected-token behavior for raw `{...}` placeholders still passes.

## Out Of Scope

- Entity extractor integration.
- Order-sensitive or nesting-sensitive protected-token validation.
- Automatic discovery of allowed tag names from the data.
- Legacy conversion of old `{tN_op}` workbooks.
- Interactive custom-config prompts.
