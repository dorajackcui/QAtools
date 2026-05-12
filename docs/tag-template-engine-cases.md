# Tag And Template Engine Cases

Date: 2026-05-11

This document summarizes the current behavior of `phraseloom.tag_engine` and
`phraseloom.template_engine` after the protected-token redesign. It is meant to
describe the cases the code handles today, including edge cases and workflow
boundaries.

## Responsibility Split

| Layer | Owns | Does not own |
|---|---|---|
| `tag_engine` | Protect raw inline structure before template parsing. `tag_engine` protects only tag names allowed by the active tag rules. Unknown angle-bracket labels remain normal text. Raw `{...}` placeholders remain protected when `raw_braces.protect_all = true`. | Numeric variable extraction, template grouping, entity clustering. |
| `template_engine` | Extract reusable numeric/color variables from protected text and infer/apply target templates. | Raw tag parsing, raw `{...}` placeholder parsing, restoring original raw spans. |
| `workflow` / `excel_io` | Connect both layers when reading rows, building translation units, validating fill output, and restoring raw text. | Deep tag nesting QA beyond the current warnings. |

## Protected Token Contract

Protected tokens are translator-facing and use one monotonically increasing
sequence per source segment.

| Token shape | Meaning | Example raw span |
|---|---|---|
| `{N>` | Opening protected span | `<color=#123>`, `[color=#ff0]` |
| `<N}` | Closing protected span | `</color>`, `[/]`, `[/b]` |
| `{N}` | Single protected span | `<br/>`, `{0}`, `{player.name}` |

Current rules:

- `N` starts at `1`; `{0}` is raw source syntax, not a protected token.
- Each raw span gets its own number in source occurrence order.
- Opening and closing tags do not share the same number.
- Pairing is stored in `TagToken.partner_index`, not encoded in token numbers.
- `parse_protected_token("{N}")` returns the generic `single` kind; the original
  `TagToken.kind` distinguishes self-closing tags from raw brace placeholders.
- Legacy helper names such as `make_tag_placeholder` are aliases for protected
  token helpers.
- Tag extraction is governed by `phraseloom/tag_rules.toml` by default. The
  default angle allowlist is `color`, `size`, `img`, `br`, `i`, `u`, `outline`,
  and `c`; the default BBCode allowlist is `color`, `b`, `i`, `u`, and `size`.
- Angle tag rules also carry pairing semantics. By default `c` is an alias for
  `color`, `img` and `br` are single tags even without a trailing `/`, and `i`
  and `outline` are optional pairs: they are paired only when a later named
  close exists.
- Generated workbooks record tag-rule metadata. Translated and TM workbook
  loading validates matching metadata when both sides include it; legacy
  no-metadata workbooks still load.

## Tag Engine Extraction Cases

`extract_tags()` scans one source string and returns protected text, token
metadata, and warnings.

| Case | Raw input | Protected output | Notes |
|---|---|---|---|
| Allowed angle tag | `<color=#123>HP {a}</>` | `{1>HP {2}<3}` | `color` is in the default allowlist. |
| Unknown angle label | `<Activate> HP {a}` | `<Activate> HP {1}` | `activate` is not in the default allowlist. |
| Custom allowed angle tag pair | `<a href="shop">VIP10</a>` | `{1>VIP10<2}` | Protected only when `a` is allowed by the active custom tag rules. |
| Shorthand angle close | `<color=#123>Text</>` | `{1>Text<2}` | `</>` closes the nearest protected open tag. |
| Self-closing angle tag | `<img src="coin.png"/>` | `{1}` | Single token. |
| Configured single angle tag | `<br>Line<img src="coin.png">` | `{1}Line{2}` | `br` and `img` are single by config even when not written as `/>`. |
| Angle close alias | `<color=#123>Text</c>` | `{1>Text<2}` | `c` canonicalizes to `color` before stack matching. |
| Optional-pair angle tag without close | `<i><size={a}>Voir</size>` | `{1}{2>Voir<3}` | `i` becomes a single token when no later named `</i>` exists. |
| Optional-pair style wrapper without close | `<outline width={a}><color=#fff>Time</c>` | `{1}{2>Time<3}` | `outline` behaves the same way when no later named `</outline>` exists. |
| Optional-pair angle tag with close | `<i>Voir</i>` | `{1>Voir<2}` | A later named close keeps `i` as a normal pair. |
| BBCode pair | `[b]Bold[/b]` | `{1>Bold<2}` | Named BBCode close must match the stack top. |
| BBCode parameter with shorthand close | `[color=#ff0]Bonus[/]` | `{1>Bonus<2}` | Color value stays out of template parsing. |
| Raw numeric placeholder | `Hit deals {0} damage` | `Hit deals {1} damage` | Complete raw `{...}` spans are protected. |
| Raw named placeholder | `Hello {player}` | `Hello {1}` | All raw brace contents are considered non-translatable. |
| Raw complex placeholder | `Value {player.name:N2}` | `Value {1}` | Colon and dots inside raw braces are protected. |
| Mixed tag and raw placeholder | `[color=#1213]Hit {0}[/]` | `{1>Hit {2}<3}` | One index sequence covers tags and raw braces together. |
| Legacy placeholder text | `{t1_op}<color=#123>x</>` | `{1}{2>x<3}` | Old placeholder-looking text is treated as raw `{...}`. |
| Repeated same raw span | `<br/> A <br/>` | `{1} A {2}` | Repeated spans get distinct tokens. |
| Plain UI bracket text | `Press [OK] to continue` | `Press [OK] to continue` | Unmatched plain `[name]` text is left raw. |
| Incomplete raw brace | `Use {abc` | `Use {abc` | Left raw without warning. |
| Nested raw brace text | `Use {a{b}c}` | `Use {a{1}c}` | Only the complete inner `{b}` is protected. |

### Tag Pairing And Warning Cases

| Case | Raw input | Protected output | Warning behavior |
|---|---|---|---|
| Unpaired close | `</color>Text` | `</color>Text` | Close remains raw; warning includes `unpaired close tag`. |
| Unclosed open | `<color=#123>Text` | `{1>Text` | Open is serialized; warning includes `open tag has no close partner`. |
| Misnested named tags | `<color=#123><i>x</color>y</i>` | `{1>{2>x</color>y<3}` | The `</color>` close is not cross-matched over `<i>`; warnings are emitted. |
| Self-closing validation | `<br/>` then target `{1}` | `{1}` | Counts as valid because source has one single protected token. |

The pairing model is intentionally conservative. Named closes only match the
top of the stack after angle aliases are canonicalized. Shorthand closes such
as `[/]` and `</>` pop the nearest open tag. Optional-pair angle tags are
opened only when a later named close for the same canonical tag exists;
otherwise they are serialized as single protected spans.

## Tag Engine Target-Side Helpers

These helpers are used when existing targets, examples, TM rows, or translated
todo workbooks need to line up with source-side protected metadata.

| Function | Current behavior | Important edge |
|---|---|---|
| `serialize_known_tags(target, tags)` | Replaces exact raw spans from source metadata with their source protected tokens, one occurrence at a time. | It does not independently discover new target tags. If the target raw tag differs from source metadata, it stays raw and a `source_protected_span_not_found` warning is recorded. |
| `validate_tag_placeholders(target, tags)` | Compares source protected-token counts with protected tokens found in target text. | Validation is count-based. It reports missing and extra tokens, but does not enforce ordering or nesting. |
| `restore_tags(target, tags)` | Replaces known protected tokens with their original raw spans. | Unknown protected-looking tokens remain unchanged. |
| `is_tag_only_segment(source)` | Returns true when removing protected tokens and whitespace leaves nothing. | `{1}`, `{1><2}`, and `{1> <2}` are protected-only; `{1>Click<2}` and `{1} 100 coins` are not. |

Example fill path:

| Source raw | Source protected | Translator target unit | Row values | Serialized target | Restored target |
|---|---|---|---|---|---|
| `<color=#123>VIP10</>` | `{1>VIP10<2}` | `{1>Pack VIP{num1}<2}` | `num1=10` | `{1>Pack VIP10<2}` | `<color=#123>Pack VIP10</>` |
| `Hit deals {0} damage` | `Hit deals {1} damage` | `Inflige {1} degats` | none | `Inflige {1} degats` | `Inflige {0} degats` |

If the target unit is `{1>Pack VIP{num1}<2} {9}`, validation reports extra
`{9}` and restore leaves `{9}` unchanged.

## Template Engine Parsing Cases

`parse_template()` receives protected text, not raw workbook text. It scans only
outside the protected token literals themselves. Text between `{N>` and `<N}` is
ordinary text and can still become a template.

| Input to template parser | Template | Values | Notes |
|---|---|---|---|
| `VIP10 Paid Pack` | `VIP{num1} Paid Pack` | `num1=10` | Plain integers become `num`. |
| `Clear Story 10-20` | `Clear Story {stage1}` | `stage1=10-20` | Hyphen-only numeric ranges become `stage`. |
| `Time 10:30` | `Time {seq1}` | `seq1=10:30` | Colon-separated values become `seq`. |
| `Date 2026/05/11` | `Date {seq1}` | `seq1=2026/05/11` | Slash-separated values become `seq`. |
| `Version 1.5` | `Version {seq1}` | `seq1=1.5` | Current regex order classifies dotted decimals as `seq`. |
| `Gain 0.99 coins` | `Gain {seq1} coins` | `seq1=0.99` | Same dotted-decimal behavior. |
| `Color #ff00aa` | `Color {color1}` | `color1=#ff00aa` | Bare 6-digit hex colors become `color`. |
| `Hit deals {1} damage` | `Hit deals {1} damage` | none | Protected single token is ignored by variable extraction. |
| `{1>VIP10 Pack<2} {3}` | `{1>VIP{num1} Pack<2} {3}` | `num1=10` | Digits in `{1>` and `<2}` are ignored; `VIP10` is not. |
| `{1} 100 coins` | `{1} {num1} coins` | `num1=100` | Text outside protected tokens is still parsed. |

Variable key selection is current implementation behavior:

- `colorN` for `#[0-9A-Fa-f]{6}`.
- `stageN` for hyphen-only numeric ranges such as `10-20`.
- `seqN` for separator sequences using `.`, `/`, or `:`, including current
  dotted decimal matches such as `0.99`.
- `numN` for plain integer matches.

## Template Candidate And Grouping Cases

Template grouping happens in `workflow`, but it depends directly on
`template_engine.is_candidate_template()`.

| Case | Parsed template | Candidate? | Workflow result |
|---|---|---|---|
| `VIP10 Paid Pack` | `VIP{num1} Paid Pack` | Yes | Can group with `VIP20 Paid Pack` as one template unit. |
| `Hit deals {1} damage` | `Hit deals {1} damage` | No | No template values, so it remains a segment unit. |
| `{1>VIP10 Pack<2}` | `{1>VIP{num1} Pack<2}` | Yes | Protected tags do not block text inside them from clustering. |
| `10-20` | `{stage1}` | No | Literal content after removing variables is too small, so it remains a segment and can auto-fill as non-translatable. |
| `...` | `...` | No | No values and no letters; segment can auto-fill as non-translatable. |
| `{1}` | `{1}` | No | Segment can auto-fill as protected-only. |

A template group is emitted only when the group has at least
`min_group_size` unique protected source strings.

## Target Template Inference Cases

`infer_target_template(values, target_text)` replaces known source values in the
target, but only outside protected token literals.

| Source values | Target text | Inferred target template | Notes |
|---|---|---|---|
| `num1=10` | `VIP10pack` | `VIP{num1}pack` | Normal single-value inference. |
| `num1=1` | `{1>Niveau 1<2}` | `{1>Niveau {num1}<2}` | Does not replace the `1` inside `{1>` or `2` inside `<2}`. |
| none | `Login failed` | none | Falls back to segment behavior. |
| `num1=10, num2=20` | `Pack 20 / 10` | `Pack {num2} / {num1}` | Values are sorted longest-first before replacement. |

Current caveat: if two source variables have the same value, inference is
value-based and may collapse multiple target occurrences to the first matching
key. The unit warning path can then report missing source variables.

`apply_target_template(target_template, values)` is a direct placeholder
replacement. It does not validate protected tokens by itself; workflow validates
after applying row values.

## Workflow Integration Cases

### Reading Source Rows

`excel_io._read_source_rows()` performs this sequence per row:

```text
raw source
  -> extract_tags(raw source)
  -> parse_template(protected source)

raw existing target
  -> serialize_known_tags(raw existing target, source tag metadata)
```

Stored `RowItem` fields keep both forms:

- `raw_source` and `raw_existing_target` preserve workbook text.
- `source` and `existing_target` hold protected text.
- `tag_tokens`, `tag_warnings`, and `target_tag_warnings` carry metadata and
  warning state for later fill.

### Building Translation Units

| Unit path | Input condition | Target unit source |
|---|---|---|
| Provided template unit | Matching unit from TM, translated workbook, or example. | `tm_pairs`, `translation_units`, or `example: ...` |
| Suggested template unit | Existing targets infer a common target template. | `existing_target` |
| Provided segment unit | Matching segment from TM, translated workbook, or example. | Same as provided source. |
| Suggested segment unit | Existing targets agree on a raw segment translation. | `existing_target` |
| Protected-only segment | Source unit is only protected tokens and whitespace. | `tag_only` |
| Non-translatable segment | Source has no letters after trimming. | `non_translatable` |
| New unit | No target can be inferred or supplied. | blank target, sent to translator todo. |

### Filling Targets

For each row, workflow performs:

```text
target_unit
  -> apply_target_template(row values) if unit is a template
  -> validate_tag_placeholders(serialized target, row tag metadata)
  -> restore_tags(serialized target, row tag metadata)
  -> write output target
```

Fill is warning-based. A protected-token mismatch does not block output writing.
The warning is surfaced in workbook QA/status columns.

## Current Boundaries To Remember

- Raw `{...}` spans are protected before template parsing. The template layer
  should not try to interpret `{player}`, `{0}`, or legacy-looking `{t1_op}`.
- Protected-token validation is count-based, not order-based.
- Existing target serialization uses source-side raw spans. It does not support
  independently invented target-only tags unless they are already written as
  protected tokens in the translator todo. Equivalent-looking raw tags can still
  miss serialization when their raw text differs, for example normal space vs
  non-breaking space in attributes.
- Text inside opening and closing protected tags is still templateable.
- Text inside a single protected token is not visible to template parsing.
- Dot decimals currently use `seqN`, not `numN`, because the separator regex is
  matched before the plain numeric regex.
