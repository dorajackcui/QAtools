# PhraseLoom Strings Workflow Redesign (Archived)

## Product Definition

PhraseLoom is a deterministic Excel strings preprocessing and restore tool.

It exports untranslated strings into one clean, self-contained workbook,
retains the established Translation Unit cleaning, optionally places
structurally similar cleaned units together, and restores completed
translations to a copy of the original workbook.

The primary workflow does not use translation memory, infer translations, build
terminology, or compose targets from extracted values. The established template
compression remains preprocessing: one translated template is expanded back to
the original row values during restore.

## User Workflow

```text
Source Excel
  -> Export Strings
  -> Translate the strings sheet
  -> Restore Translations
  -> Translated copy of the original Excel
```

The GUI and interactive terminal expose only these two actions.

## Export Rules

Rows are processed in this order:

1. Ignore rows with an empty Source.
2. Replace protected Tags, brace placeholders, and literal `\n` / `\r` markers
   with Protected Tokens.
3. Treat a non-empty Target as completed; exclude it from user work and preserve
   it in the completed reference section.
4. Treat numeric-only, symbol-only, and Tag-only Source values as automatic
   passthrough rows: exclude them from user work and use the original Source as
   Target during restore.
5. Collect the remaining rows as pending strings.
6. Run the established Translation Unit cleaning:
   - deduplicate exact Source values;
   - compress eligible numeric and sequence variants into templates such as
     `通用补偿器LV{num1}`.
7. By default, order cleaned units by their first original source row.
8. If similarity grouping is explicitly enabled, cluster the cleaned
   Translation Units by structural similarity.

Completed rows are evaluated independently. A completed translation for one
row is never copied to another pending row with the same Source.

## Similarity Grouping

Similarity grouping is an optional, default-off display-order feature.

It may use internal generalized patterns to identify related Source values, but
it only assigns a group ID and display order. It does not extract term/structure
tables, request partial translations, compose Targets, or create extra workflow
stages.

Every cleaned Translation Unit remains an independent translation row.
Variants intentionally compressed into the same established template share one
translation row; ordinary similar strings do not.

When grouping is disabled, the `group` column is empty and units remain in
first-source-row order. When enabled, all unclustered units remain first,
followed by clustered units. Each cluster is contiguous; clusters and their
members are ordered by their earliest original source row. Unclustered strings
remain valid and are restored through exactly the same mapping mechanism.

## Strings Workbook

The workbook has two visible workflow sheets:

1. `strings` contains only content that requires user translation.
2. `completed` contains rows with an existing Target and automatic passthrough
   rows.

Visible `strings` columns:

```text
group
source
target
sample_sources
context
occurrences
```

`string_id` is present but hidden. The original workbook, row mapping, metadata,
column configuration, and Tag rules are embedded as hidden content so restore
requires no second source file.

The Source column is immutable workflow data. It may contain template variables
or Protected Tokens produced by cleaning. Restore rejects a workbook when a
visible Source or embedded original Source was changed.

`sample_sources` exposes the first original raw Source represented by the
cleaned unit, while `context` carries the configured reference context from
that same original row. This preserves the established paired-reference
semantics. Both are translator reference data and are not used as restore keys.

Visible `completed` columns:

```text
status
source
target
context
```

`existing_target` rows preserve their original Target. `auto_passthrough` rows
show the Source copied into Target. The sheet is a review section and is not a
translation input surface.

## Restore Rules

Restore:

- writes each translated string to every mapped pending row;
- expands template variables using each original row's hidden values;
- restores Protected Tokens to the original Tags, placeholders, and literal
  `\n` / `\r` markers;
- never touches rows skipped because they already had a Target;
- writes each automatic passthrough Source into its empty Target;
- restores original sheet visibility and active-sheet state;
- preserves the original workbook structure and formatting;
- checks protected Tags, variables, placeholders, and literal markers;
- writes a separate restore-issues workbook only when a Target is empty or a
  protected-content warning exists.

Numeric-only, symbol-only, Tag-only, and protected-placeholder-only rows are
restored as automatic passthrough Targets.

## Interfaces

Primary CLI:

```text
phraseloom export SOURCE.xlsx
phraseloom export SOURCE.xlsx --group-similar
phraseloom restore SOURCE_STRINGS.xlsx
```

Primary GUI:

```text
Export Strings
Export Strings page -> Restore Translations button
```

Source, Target, Context, output path, and Tag configuration remain available as
secondary export options. Clustering thresholds are not part of the primary UX.
