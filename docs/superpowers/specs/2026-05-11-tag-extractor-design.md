# PhraseLoom Tag Extractor Design

Date: 2026-05-11

## Goal

Add an internal tag extraction layer to PhraseLoom so HTML-like and BBCode-like
formatting tags stop interfering with template discovery, TM matching, machine
translation, and fill-back.

The user workflow remains the same:

```text
1. Build TM from completed Excel
2. Prepare translator file for new source
3. Fill source from translated file
```

Tag handling is automatic. It does not add a new user step, a required CLI
argument, or a translator-facing sheet to manage.

## Recommended Approach

Use a conservative, extensible tag scanner before template parsing. The scanner
serializes recognized tags into protected placeholders, then the existing
template flow runs on the serialized text.

Example:

```text
raw source:
<a href="shop">VIP10 Pack</a>

tag-serialized source:
{t1_op}VIP10 Pack{t1_cl}

template source:
{t1_op}VIP{num1} Pack{t1_cl}
```

On fill, PhraseLoom reverses the order:

```text
target_unit
  -> apply template placeholders
  -> validate tag placeholders
  -> restore raw tags
  -> write Excel output
```

This keeps tag handling and template handling as two separate pure-text layers.

## Placeholder Contract

Tag placeholders use a compact reserved namespace:

```text
{t1_op}
{t1_cl}
{t2_sf}
```

The format is intentionally centralized behind helpers and constants, not
hardcoded at call sites:

```python
TAG_PLACEHOLDER_PREFIX = "t"
TAG_OPEN_SUFFIX = "op"
TAG_CLOSE_SUFFIX = "cl"
TAG_SELF_SUFFIX = "sf"
```

Helpers own generation and parsing:

```text
make_tag_placeholder(1, "op") -> "{t1_op}"
make_tag_placeholder(1, "cl") -> "{t1_cl}"
make_tag_placeholder(2, "sf") -> "{t2_sf}"
is_tag_placeholder("{t1_op}") -> True
parse_tag_placeholder("{t1_op}") -> index=1, kind="op"
```

Rules:

- The tag extractor owns this namespace.
- Template parsing must preserve tag placeholders exactly.
- Template values must not contain `t1_op`, `t1_cl`, or `t2_sf` variables.
- Raw text that already contains this namespace is a warning condition.
- If the placeholder format changes later, it should be changed in the helper
  layer and tests, not across workflow code.

## Tag Detection

Add a pure module:

```text
phraseloom/tag_engine.py
```

It owns:

- tag scanning
- tag placeholder helpers
- tag-only segment detection
- tag placeholder validation
- raw tag restoration

The first implementation should use a conservative rule list instead of a broad
"capture anything inside brackets" regex. Rules are extensible so future tag
forms can be added without changing workflow orchestration.

Initial recognized families:

- self-closing angle tags, such as `<img src="..."/>`
- angle open tags, such as `<a href="...">`
- angle close tags, such as `</a>`
- shorthand close tags, such as `</>`
- BBCode-style open tags, such as `[color=#ff0]`
- BBCode-style close tags, such as `[/]`

Unknown or suspicious tag-like text should remain raw text and produce a
warning. This reduces damage from irregular source text.

## Data Model

`tag_engine.extract_tags()` returns structured data rather than only a string:

```python
TagExtraction(
    text="{t1_op}here{t1_cl}",
    tags=[
        TagToken(index=1, kind="op", placeholder="{t1_op}", raw='<a href="x">'),
        TagToken(index=1, kind="cl", placeholder="{t1_cl}", raw="</a>"),
    ],
    warnings=[],
)
```

Open and close tokens share the same index. Self-closing tokens use their own
index.

For generic closes such as `</>` and `[/]`, the extractor pairs the close with
the nearest compatible unclosed open token when possible. If pairing is unclear,
the extractor should keep the text raw or emit a warning instead of guessing
aggressively.

## Workflow Integration

### Source Row Reading

The current row read path parses templates directly from raw source text:

```text
raw source -> parse_template(raw source) -> RowItem
```

The new path is:

```text
raw source
  -> extract_tags(raw source)
  -> parse_template(tag_serialized_source)
  -> RowItem
```

`RowItem.source` should represent the tag-serialized source used by template
and workflow logic. Raw source can stay available through the original workbook
values and optional debug columns.

### TM Extraction

TM extraction must serialize both source and target text before inferring target
templates.

Example:

```text
raw source: <a href="x">VIP10</a>
raw target: <a href="x">VIP10 Pack FR</a>

serialized source: {t1_op}VIP10{t1_cl}
serialized target: {t1_op}VIP10 Pack FR{t1_cl}

source_unit: {t1_op}VIP{num1}{t1_cl}
target_unit: {t1_op}VIP{num1} Pack FR{t1_cl}
```

Target serialization should prefer the source row's tag map when raw target tags
match the source tags. If the target has extra or missing tags, keep processing
and surface warnings.

### Translation Preparation

Translator-facing `source_unit` values use serialized tag placeholders:

```text
{t1_op}Click here{t1_cl}
{t2_sf} Get {num1} coins
```

The translator still only fills `target_unit`. No extra manual tag sheet is part
of the normal workflow.

### Fill

Fill reads the source workbook again and extracts each row's tag map at that
time. It does not rely on a row-level tag map saved in the translator workbook.

Per row:

```text
source raw
  -> extract current row tags
  -> apply target template values
  -> validate tag placeholders against source row tag map
  -> restore known tag placeholders
  -> write target
```

This keeps the translator workbook lightweight. The source workbook passed to
fill remains the source of truth for row-level raw tag attributes.

## Tag-Only Units

After serialization, some rows become pure tag placeholders:

```text
<img src="icon.png"/>
=> {t1_sf}
```

These should be auto-filled during translation unit building, alongside the
existing non-translatable numeric and symbol segments.

Conceptual rule:

```python
elif is_tag_only_segment(source_unit):
    target_unit = source_unit
    target_unit_source = "tag_only"
elif is_non_translatable_segment(source_unit):
    target_unit = source_unit
    target_unit_source = "non_translatable"
```

`is_tag_only_segment()` is conservative: after removing whitespace, the segment
must contain only tag placeholders. Mixed text remains translatable.

Auto-filled examples:

```text
{t1_sf}
{t1_op}{t1_cl}
{t1_op} {t1_cl}
```

Not auto-filled:

```text
{t1_op}Click{t1_cl}
{t1_sf} 100 coins
```

## Validation And Warnings

Tag validation checks target placeholders against the source row tag map.
Position can change. Counts and identity should match.

Expected cases:

```text
source: Click {t1_op}here{t1_cl}
target: Cliquez {t1_op}ici{t1_cl}
=> ok

target: Cliquez ici
=> warning: tag_mismatch, missing {t1_op}, {t1_cl}

target: {t1_op}Cliquez ici{t1_cl}{t2_sf}
=> warning: tag_mismatch, extra {t2_sf}
```

PhraseLoom should still write the target even when tag validation warns. The
main workflow remains permissive because a downstream tag QA process will do a
stricter check.

Known placeholders are restored. Unknown placeholders remain visible in output
so the downstream checker can catch them.

## Workbook UX

The user-facing workflow stays simple:

- no new user step
- no new required command argument
- no manual tag extraction command
- no translator-managed tag map sheet
- pure tag units stay out of `to_translate`

`to_translate` remains centered on the same task:

```text
source_unit
target_unit
warning
translator_note
```

The process workbook can include lightweight debug information where useful,
especially in `source_map`:

```text
raw_source
tagged_source
tag_warning
```

This is for troubleshooting, not a new translator responsibility.

`qa_report` should add aggregate counts such as:

```text
tag_mismatch_rows
tag_warning_units
tag_only_units
```

## Error Handling

Tag problems are warnings, not hard failures, for this feature's first version.

Warnings should cover:

- raw source contains reserved tag placeholder namespace
- suspicious tag-like text was not serialized
- close tag has no clear open partner
- target is missing source tag placeholders
- target contains extra tag placeholders
- target contains unknown tag placeholders

Warnings should be attached to units or source rows wherever the current
workbook output can display them without increasing user workload.

## Testing Strategy

Add focused pure tests for `tag_engine.py`:

- placeholder helper generation and parsing
- source extraction for angle and BBCode tags
- paired open/close serialization
- self-closing serialization
- generic close pairing
- suspicious text preserved with warning
- namespace conflict warning
- tag-only segment detection
- tag restore with reordered placeholders
- tag mismatch reporting

Add workflow tests:

- source tags do not prevent numeric template extraction
- TM extraction serializes source and target tags consistently
- pure tag units auto-fill and do not enter `to_translate`
- fill applies template values before restoring tags
- fill writes target even when tag mismatch warning is present
- summary and QA counts include tag-related rows

## Out Of Scope

- Adding a separate user-visible tag-cleaning command.
- Making tag validation block fill output.
- Building a full HTML or BBCode parser.
- Supporting pre-feature workbooks that predate this design.
- Making translators manage row-level tag maps manually.
- Solving all malformed source text cases in the first version.

## Acceptance Criteria

- Tag extraction is enabled internally by default in all three existing workflow
  steps.
- User steps and required CLI arguments remain unchanged.
- Tag placeholders use the centralized `{tN_op}` / `{tN_cl}` / `{tN_sf}`
  contract.
- Template parsing does not convert tag placeholders into template variables.
- Pure tag-only units are auto-filled.
- Fill restores raw tags after template application.
- Tag mismatch writes the target and records warnings.
- Tests cover tag engine behavior and end-to-end workflow behavior.
