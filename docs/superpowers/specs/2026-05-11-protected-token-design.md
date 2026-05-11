# PhraseLoom Protected Token Design

Date: 2026-05-11

## Goal

Replace the current tag-only placeholder contract with a broader protected token
layer. This layer serializes all non-translatable inline structure before
template parsing, writes the serialized text into translator todo workbooks, and
restores the original raw text during fill.

The protected layer covers:

- HTML-like tags
- BBCode-like tags
- every complete non-nested raw `{...}` segment in source text

This keeps tag syntax, placeholder syntax, colors inside tags, and raw
placeholder internals out of the template engine.

## Recommended Approach

Keep `phraseloom.tag_engine` as the protected-token serializer module. Its
behavioral boundary expands from "tags only" to "inline protected spans."

Example:

```text
raw source:
[color=#1213] 击打造成{0}伤害[/]

protected source:
{1> 击打造成{2}伤害<3}
```

Here:

- `{1>` maps to `[color=#1213]`
- `{2}` maps to `{0}`
- `<3}` maps to `[/]`

The translator and any LLM-assisted translation workflow sees the protected
tokens directly in `source_unit` and `target_unit`. The tokens are intentionally
compact because they are part of the translator-facing contract.

## Token Contract

Protected tokens use one global, monotonically increasing sequence per source
segment. Numbers are assigned by raw-text occurrence order.

```text
{N>   opening protected span
<N}   closing protected span
{N}   single protected span
```

Rules:

- `N` starts at `1`.
- Each token instance gets its own unique number.
- Opening and closing tags do not share an id.
- Pairing is stored in metadata, not inferred from matching numbers.
- Single tokens cover self-closing tags and raw `{...}` segments.
- The raw token text is restored from metadata during fill.

Examples:

| Raw input | Protected output | Notes |
|---|---|---|
| `<a href="shop">VIP10</a>` | `{1>VIP10<2}` | angle open and close |
| `<img src="coin.png"/>` | `{1}` | self-closing tag |
| `[color=#ff0]Bonus[/]` | `{1>Bonus<2}` | BBCode open and close |
| `造成{0}伤害` | `造成{1}伤害` | raw placeholder |
| `Hello {player}` | `Hello {1}` | named raw placeholder |
| `获得{item.name:N2}` | `获得{1}` | complex raw placeholder |
| `<br/> A <br/>` | `{1} A {2}` | repeated raw tags get distinct tokens |

## Extraction Cases

The protected layer keeps the current conservative tag behavior and adds raw
brace placeholders.

| Case | Raw input | Protected output | Behavior |
|---|---|---|---|
| Angle pair | `<a>x</a>` | `{1>x<2}` | serialize both tags |
| Shorthand angle close | `<a>x</>` | `{1>x<2}` | close restores to `</>` |
| Self-closing angle tag | `<br/>` | `{1}` | single token |
| BBCode pair | `[b]x[/b]` | `{1>x<2}` | serialize both tags |
| BBCode parameter tag | `[color=#ff0]x[/]` | `{1>x<2}` | color stays out of templates |
| UI text | `Press [OK] to continue` | `Press [OK] to continue` | unchanged |
| Unpaired close | `</a>Text` | `</a>Text` | unchanged with warning |
| Unclosed open | `<a>Text` | `{1>Text` | serialized with warning |
| Misnested tags | `<a><b>x</a>y</b>` | `{1>{2>x</a>y<3}` | do not cross-match |
| Raw braces | `Use {abc-1}` | `Use {1}` | every complete `{...}` is protected |
| Incomplete braces | `Use {abc` | `Use {abc` | unchanged without warning |

## Template Layer Boundary

`phraseloom.template_engine` treats protected tokens as literal, indivisible
text. It does not parse, rename, infer, or apply values inside them.

The template layer may still extract dynamic values outside protected tokens:

- integers and decimals
- stage/range-like numbers such as `1-2`
- date/time/sequence-like values such as `10:30` or `2026/05/11`
- bare color values such as `#ff00aa` when they appear outside protected tokens

It no longer handles raw `{...}` placeholders. Those are protected before the
template layer runs.

| Input to template layer | Template output | Values | Notes |
|---|---|---|---|
| `造成{1}伤害` | `造成{1}伤害` | none | `{1}` is protected |
| `{1>VIP10 Pack<2}` | `{1>VIP{num1} Pack<2}` | `num1=10` | text inside tags can template |
| `{1>造成{2}伤害<3}` | `{1>造成{2}伤害<3}` | none | protected placeholder remains literal |
| `{1} 100 coins` | `{1} {num1} coins` | `num1=100` | single token remains literal |
| `颜色#ff00aa` | `颜色{color1}` | `color1=#ff00aa` | bare color remains templateable |
| `{1}` | `{1}` | none | protected-only unit |

Implementation implication: template parsing and target-template inference must
be protected-aware. They should split text into protected and ordinary spans,
then run variable detection only on ordinary spans. This prevents values such as
`1` or `2` from corrupting `{1>`, `<2}`, or `{1}`.

## Workflow

Extraction and TM creation:

```text
raw source row
  -> protected-token extraction
  -> template parsing on protected text
  -> TM matching and todo workbook writing
```

Existing target serialization:

```text
raw source row
  -> protected-token extraction and metadata
raw existing target
  -> replace known raw protected spans with source metadata tokens
```

Fill:

```text
target_unit
  -> apply template values
  -> validate protected tokens
  -> restore raw protected spans
  -> write workbook output
```

## Validation And Warnings

Validation remains warning-based. A mismatch does not block fill output.

Warnings should cover:

- source protected token missing from target
- target contains extra protected token
- unpaired close tag
- unclosed open tag
- known source raw protected span not found in existing target

The first implementation can keep count-based token validation. A stricter
ordering or nesting QA pass can be added later without changing the token
contract.

## Tag-Only And Protected-Only Units

The existing tag-only behavior becomes protected-only behavior.

A unit should auto-fill when removing all protected tokens and whitespace leaves
an empty string:

```text
{1}
{1><2}
{1> <2}
```

These should not auto-fill:

```text
{1>Click<2}
{1} 100 coins
```

## Compatibility

The new protected-token contract replaces `{tN_op}`, `{tN_cl}`, and `{tN_sf}` in
newly generated workbooks. The implementation should centralize token parsing
behind helpers so old compatibility shims can be handled deliberately if needed.
Some internal Python helper names may retain `tag_` wording for compatibility,
but their behavior follows the protected-token contract.

Existing workbooks that already contain `{tN_op}` tokens are not part of the new
translator-facing format. This spec does not require legacy workbook support.
If it becomes necessary later, add a dedicated legacy reader path rather than
keeping old tokens in newly generated todo workbooks.

## Testing

Focused tests should cover:

- token helper generation and parsing
- current angle tag, BBCode, shorthand close, self-closing, and warning cases
- raw `{...}` protection for simple, named, and complex brace contents
- template parsing that preserves `{N>`, `<N}`, and `{N}`
- target-template inference that does not replace digits inside protected tokens
- protected-only auto-fill
- fill validation for missing and extra protected tokens
- restoration of mixed tags and raw brace placeholders
