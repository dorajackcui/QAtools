# Protected Token Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace PhraseLoom's `{tN_op}` tag placeholders with translator-facing protected tokens `{N>`, `<N}`, and `{N}` that also protect every raw `{...}` placeholder before template parsing.

**Architecture:** Keep the public three-step workflow unchanged. Expand `phraseloom.tag_engine` into a pure protected-token layer, make `phraseloom.template_engine` split around protected tokens before extracting template variables, and keep `workflow.py` responsible for applying templates, validating protected tokens, and restoring row-level raw spans during fill.

**Tech Stack:** Python 3.11+, standard-library `re`, `dataclasses`, `unittest`, `openpyxl`, existing PhraseLoom package modules.

---

## Scope

This plan implements the protected-token foundation described in `docs/superpowers/specs/2026-05-11-protected-token-design.md`.

It does not integrate entity clustering into the main workflow. Entity clustering should get a separate plan after this foundation lands, using the protected text and a semantic text view derived from it.

## File Structure

- Modify `phraseloom/tag_engine.py`: generate, parse, validate, and restore `{N>`, `<N}`, and `{N}` tokens; protect raw `{...}` spans; keep conservative tag pairing behavior.
- Modify `phraseloom/template_engine.py`: remove raw `{...}` placeholder extraction and run variable detection only outside protected tokens.
- Modify `phraseloom/excel_io.py`: keep source/target row serialization calls, update variable summaries and QA labels to the protected-token vocabulary.
- Modify `phraseloom/workflow.py`: keep orchestration shape, update imports and warning names to protected-token helpers.
- Modify `tests/test_tag_engine.py`: replace old `{tN_*}` expectations with the new token contract and add raw brace protection coverage.
- Modify `tests/test_template_workflow.py`: update template and workbook workflow expectations for protected tokens.
- Modify `tests/test_tag_workflow_testfiles.py`: update real-workbook tag workflow expectations and QA warning checks.
- Modify `agent.md`: update onboarding notes so future sessions learn the new protected-token contract.

---

### Task 1: Add New Protected Token Helper Contract

**Files:**
- Modify: `tests/test_tag_engine.py`
- Modify: `phraseloom/tag_engine.py`

- [ ] **Step 1: Replace the placeholder helper test with the protected-token contract**

In `tests/test_tag_engine.py`, replace `test_placeholder_helpers_accept_only_tag_namespace` with:

```python
    def test_protected_token_helpers_accept_new_contract(self):
        from phraseloom.tag_engine import (
            RAW_PLACEHOLDER,
            TAG_CLOSE,
            TAG_OPEN,
            TAG_SELF,
            is_protected_token,
            is_tag_placeholder,
            make_protected_token,
            make_tag_placeholder,
            parse_protected_token,
        )

        self.assertEqual(TAG_OPEN, "op")
        self.assertEqual(TAG_CLOSE, "cl")
        self.assertEqual(TAG_SELF, "sf")
        self.assertEqual(RAW_PLACEHOLDER, "ph")
        self.assertEqual(make_protected_token(1, TAG_OPEN), "{1>")
        self.assertEqual(make_protected_token(2, TAG_CLOSE), "<2}")
        self.assertEqual(make_protected_token(3, TAG_SELF), "{3}")
        self.assertEqual(make_protected_token(4, RAW_PLACEHOLDER), "{4}")
        self.assertEqual(make_tag_placeholder(1, TAG_OPEN), "{1>")
        self.assertTrue(is_protected_token("{1>"))
        self.assertTrue(is_protected_token("<2}"))
        self.assertTrue(is_protected_token("{3}"))
        self.assertTrue(is_tag_placeholder("{3}"))
        self.assertEqual(parse_protected_token("{1>"), (1, TAG_OPEN))
        self.assertEqual(parse_protected_token("<2}"), (2, TAG_CLOSE))
        self.assertEqual(parse_protected_token("{3}"), (3, "single"))
        self.assertIsNone(parse_protected_token("{num1}"))
        self.assertIsNone(parse_protected_token("{t1_op}"))
```

- [ ] **Step 2: Run the helper test and verify it fails**

Run:

```powershell
py -3 -m unittest tests.test_tag_engine.TagEngineTests.test_protected_token_helpers_accept_new_contract -v
```

Expected: FAIL because `RAW_PLACEHOLDER`, `make_protected_token`, `parse_protected_token`, and `is_protected_token` do not exist yet.

- [ ] **Step 3: Extend `TagToken` with pair metadata**

In `phraseloom/tag_engine.py`, change the `TagToken` dataclass to:

```python
@dataclass(frozen=True)
class TagToken:
    index: int
    kind: str
    placeholder: str
    raw: str
    partner_index: int | None = None
```

Close tokens will receive `partner_index` in Task 2 when extraction creates them.

- [ ] **Step 4: Add protected-token constants and helper functions**

In `phraseloom/tag_engine.py`, replace the old placeholder constants and regexes near the top with:

```python
TAG_OPEN = "op"
TAG_CLOSE = "cl"
TAG_SELF = "sf"
RAW_PLACEHOLDER = "ph"
PROTECTED_SINGLE = "single"

PROTECTED_TOKEN_RE = re.compile(r"\{([1-9]\d*)>|<([1-9]\d*)\}|\{([1-9]\d*)\}")
_PROTECTED_TOKEN_FULL_RE = re.compile(
    r"^\{([1-9]\d*)>$|^<([1-9]\d*)\}$|^\{([1-9]\d*)\}$"
)
TAG_PLACEHOLDER_RE = PROTECTED_TOKEN_RE
```

Replace `make_tag_placeholder`, `parse_tag_placeholder`, and `is_tag_placeholder` with:

```python
def make_protected_token(index: int, kind: str) -> str:
    if index < 1:
        raise ValueError("protected token index must be >= 1")
    if kind == TAG_OPEN:
        return f"{{{index}>"
    if kind == TAG_CLOSE:
        return f"<{index}}}"
    if kind in {TAG_SELF, RAW_PLACEHOLDER, PROTECTED_SINGLE}:
        return f"{{{index}}}"
    raise ValueError(f"unknown protected token kind: {kind}")


def parse_protected_token(token: str) -> tuple[int, str] | None:
    found = _PROTECTED_TOKEN_FULL_RE.match(token)
    if not found:
        return None
    if found.group(1):
        return int(found.group(1)), TAG_OPEN
    if found.group(2):
        return int(found.group(2)), TAG_CLOSE
    return int(found.group(3)), PROTECTED_SINGLE


def is_protected_token(token: str) -> bool:
    return parse_protected_token(token) is not None


def make_tag_placeholder(index: int, kind: str) -> str:
    return make_protected_token(index, kind)


def parse_tag_placeholder(placeholder: str) -> tuple[int, str] | None:
    return parse_protected_token(placeholder)


def is_tag_placeholder(placeholder: str) -> bool:
    return is_protected_token(placeholder)
```

- [ ] **Step 5: Update `__all__` exports**

In `phraseloom/tag_engine.py`, update the exported names to include the new helpers while keeping old aliases:

```python
__all__ = [
    "RAW_PLACEHOLDER",
    "PROTECTED_SINGLE",
    "TAG_OPEN",
    "TAG_CLOSE",
    "TAG_SELF",
    "PROTECTED_TOKEN_RE",
    "TAG_PLACEHOLDER_RE",
    "TagToken",
    "TagExtraction",
    "TagValidation",
    "make_protected_token",
    "parse_protected_token",
    "is_protected_token",
    "make_tag_placeholder",
    "parse_tag_placeholder",
    "is_tag_placeholder",
    "extract_tags",
    "is_tag_only_segment",
    "restore_tags",
    "validate_tag_placeholders",
    "serialize_known_tags",
]
```

- [ ] **Step 6: Run the helper test and verify it passes**

Run:

```powershell
py -3 -m unittest tests.test_tag_engine.TagEngineTests.test_protected_token_helpers_accept_new_contract -v
```

Expected: PASS.

- [ ] **Step 7: Commit the helper contract**

Run:

```powershell
git add phraseloom/tag_engine.py tests/test_tag_engine.py
git commit -m "Add protected token helpers"
```

Expected: commit succeeds.

---

### Task 2: Serialize Tags And Raw Brace Placeholders With New Tokens

**Files:**
- Modify: `tests/test_tag_engine.py`
- Modify: `phraseloom/tag_engine.py`

- [ ] **Step 1: Update extraction tests for angle tags**

In `tests/test_tag_engine.py`, replace `test_extracts_angle_tags_and_preserves_token_fields` with:

```python
    def test_extracts_angle_tags_and_preserves_token_fields(self):
        from phraseloom.tag_engine import TAG_CLOSE, TAG_OPEN, TAG_SELF, extract_tags

        result = extract_tags('<a href="shop">VIP10</a> <img src="coin.png"/>')

        self.assertEqual(result.text, "{1>VIP10<2} {3}")
        self.assertEqual(
            tuple(
                (tag.index, tag.kind, tag.placeholder, tag.raw, tag.partner_index)
                for tag in result.tags
            ),
            (
                (1, TAG_OPEN, "{1>", '<a href="shop">', None),
                (2, TAG_CLOSE, "<2}", "</a>", 1),
                (3, TAG_SELF, "{3}", '<img src="coin.png"/>', None),
            ),
        )
        self.assertEqual(result.warnings, ())
```

- [ ] **Step 2: Update BBCode extraction tests**

Replace `test_extracts_bbcode_tags` and `test_extracts_named_bbcode_close_tags` with:

```python
    def test_extracts_bbcode_tags(self):
        from phraseloom.tag_engine import extract_tags

        result = extract_tags("[color=#ff0]Bonus[/]")

        self.assertEqual(result.text, "{1>Bonus<2}")
        self.assertEqual(result.tags[0].raw, "[color=#ff0]")
        self.assertEqual(result.tags[1].raw, "[/]")
        self.assertEqual(result.tags[1].partner_index, 1)
        self.assertEqual(result.warnings, ())

    def test_extracts_named_bbcode_close_tags(self):
        from phraseloom.tag_engine import TAG_CLOSE, TAG_OPEN, extract_tags

        result = extract_tags("[b]Bold[/b]")

        self.assertEqual(result.text, "{1>Bold<2}")
        self.assertEqual(
            tuple(
                (tag.index, tag.kind, tag.placeholder, tag.raw, tag.partner_index)
                for tag in result.tags
            ),
            (
                (1, TAG_OPEN, "{1>", "[b]", None),
                (2, TAG_CLOSE, "<2}", "[/b]", 1),
            ),
        )
        self.assertEqual(result.warnings, ())
```

- [ ] **Step 3: Add raw brace extraction tests**

Add these tests to `TagEngineTests`:

```python
    def test_extracts_raw_brace_placeholders_as_single_tokens(self):
        from phraseloom.tag_engine import RAW_PLACEHOLDER, extract_tags

        result = extract_tags("Hit deals {0} damage to {player.name:N2}.")

        self.assertEqual(result.text, "Hit deals {1} damage to {2}.")
        self.assertEqual(
            tuple((tag.index, tag.kind, tag.placeholder, tag.raw) for tag in result.tags),
            (
                (1, RAW_PLACEHOLDER, "{1}", "{0}"),
                (2, RAW_PLACEHOLDER, "{2}", "{player.name:N2}"),
            ),
        )
        self.assertEqual(result.warnings, ())

    def test_extracts_tags_and_raw_braces_in_original_order(self):
        from phraseloom.tag_engine import RAW_PLACEHOLDER, TAG_CLOSE, TAG_OPEN, extract_tags

        result = extract_tags("[color=#1213] 击打造成{0}伤害[/]")

        self.assertEqual(result.text, "{1> 击打造成{2}伤害<3}")
        self.assertEqual(
            tuple((tag.index, tag.kind, tag.placeholder, tag.raw, tag.partner_index) for tag in result.tags),
            (
                (1, TAG_OPEN, "{1>", "[color=#1213]", None),
                (2, RAW_PLACEHOLDER, "{2}", "{0}", None),
                (3, TAG_CLOSE, "<3}", "[/]", 1),
            ),
        )

    def test_incomplete_raw_braces_stay_raw_without_warning(self):
        from phraseloom.tag_engine import extract_tags

        result = extract_tags("Use {abc")

        self.assertEqual(result.text, "Use {abc")
        self.assertEqual(result.tags, ())
        self.assertEqual(result.warnings, ())
```

- [ ] **Step 4: Run extraction tests and verify failures**

Run:

```powershell
py -3 -m unittest tests.test_tag_engine -v
```

Expected: FAIL with old `{tN_*}` serialized text and no raw brace token extraction.

- [ ] **Step 5: Add raw brace spans to the scanner**

In `phraseloom/tag_engine.py`, add this regex after `_TAG_SPAN_RE`:

```python
_RAW_BRACE_RE = re.compile(r"\{[^{}]+\}")
_PROTECTED_SPAN_RE = re.compile(
    _TAG_SPAN_RE.pattern + r"|" + _RAW_BRACE_RE.pattern
)
```

In `extract_tags`, replace:

```python
    reserved_indexes = {
        int(found.group(1)) for found in TAG_PLACEHOLDER_RE.finditer(source)
    }
    warnings = [
        f"reserved_tag_placeholder: {found.group(0)}"
        for found in TAG_PLACEHOLDER_RE.finditer(source)
    ]
    spans = list(_TAG_SPAN_RE.finditer(source))
```

with:

```python
    warnings: list[str] = []
    spans = list(_PROTECTED_SPAN_RE.finditer(source))
```

- [ ] **Step 6: Update extraction loop token assignment**

In `extract_tags`, replace the body that classifies `raw` and calls `_take_next_index` with this structure:

```python
        if _RAW_BRACE_RE.fullmatch(raw):
            index = next_index
            next_index += 1
            placeholder = make_protected_token(index, RAW_PLACEHOLDER)
            tags.append(TagToken(index, RAW_PLACEHOLDER, placeholder, raw))
            chunks.append(placeholder)
            pos = found.end()
            continue

        if _is_unmatched_plain_bbcode_open(raw, found.start(), matched_plain_bbcode_starts):
            chunks.append(raw)
            pos = found.end()
            continue

        kind, name = _classify_raw_tag(raw)

        if kind == TAG_OPEN:
            index = next_index
            next_index += 1
            stack.append((index, name))
            placeholder = make_protected_token(index, TAG_OPEN)
            token = TagToken(index, TAG_OPEN, placeholder, raw)
        elif kind == TAG_SELF:
            index = next_index
            next_index += 1
            placeholder = make_protected_token(index, TAG_SELF)
            token = TagToken(index, TAG_SELF, placeholder, raw)
        else:
            matched = _pop_matching_open(stack, name)
            if matched is None:
                chunks.append(raw)
                warnings.append(f"unpaired close tag: {raw}")
                pos = found.end()
                continue
            index = next_index
            next_index += 1
            placeholder = make_protected_token(index, TAG_CLOSE)
            token = TagToken(index, TAG_CLOSE, placeholder, raw, matched)

        tags.append(token)
        chunks.append(placeholder)
        pos = found.end()
```

Remove the old `_take_next_index` helper because protected token numbers are always sequential in extraction order.

- [ ] **Step 7: Update unclosed-open warning text generation**

In `extract_tags`, replace:

```python
f"open tag has no close partner: {make_tag_placeholder(index, TAG_OPEN)}"
```

with:

```python
f"open tag has no close partner: {make_protected_token(index, TAG_OPEN)}"
```

- [ ] **Step 8: Update misnested and shorthand tests**

In `tests/test_tag_engine.py`, update existing expected serialized text:

```python
self.assertEqual(result.text, "{1>Text")
self.assertEqual(result.text, "{1>Text<2}")
self.assertEqual(result.text, "{1>{2>x</a>y<3}")
self.assertEqual([tag.placeholder for tag in result.tags], ["{1>", "{2>", "<3}"])
```

Also update any `astuple(tag)` assertions to compare explicit five-field tuples because `TagToken` now includes `partner_index`.

- [ ] **Step 9: Run extraction tests and verify they pass**

Run:

```powershell
py -3 -m unittest tests.test_tag_engine -v
```

Expected: PASS for extraction tests after all old `{tN_*}` assertions are updated in this file.

- [ ] **Step 10: Commit extraction changes**

Run:

```powershell
git add phraseloom/tag_engine.py tests/test_tag_engine.py
git commit -m "Serialize protected spans with compact tokens"
```

Expected: commit succeeds.

---

### Task 3: Update Protected-Only, Validation, Restoration, And Known Target Serialization

**Files:**
- Modify: `tests/test_tag_engine.py`
- Modify: `phraseloom/tag_engine.py`

- [ ] **Step 1: Update protected-only tests**

Replace `test_identifies_tag_only_segments` with:

```python
    def test_identifies_protected_only_segments(self):
        from phraseloom.tag_engine import is_tag_only_segment

        self.assertTrue(is_tag_only_segment("{1}"))
        self.assertTrue(is_tag_only_segment("{1><2}"))
        self.assertTrue(is_tag_only_segment("{1> <2}"))
        self.assertFalse(is_tag_only_segment("{1>Click<2}"))
        self.assertFalse(is_tag_only_segment("{1} 100 coins"))
```

- [ ] **Step 2: Update restoration and validation tests**

Replace the old restoration and validation tests with:

```python
    def test_restore_tags_replaces_known_protected_tokens_only(self):
        from phraseloom.tag_engine import extract_tags, restore_tags

        extraction = extract_tags('<a href="shop">VIP{0}</a>')

        restored = restore_tags("{1>VIP{2}<3} {9}", extraction.tags)

        self.assertEqual(restored, '<a href="shop">VIP{0}</a> {9}')

    def test_validate_tag_placeholders_reports_extra_counts(self):
        from phraseloom.tag_engine import extract_tags, validate_tag_placeholders

        extraction = extract_tags('<a href="shop">VIP10</a>')

        validation = validate_tag_placeholders("{1>VIP10<2} {3}", extraction.tags)

        self.assertEqual(validation.warnings, ("protected_token_mismatch: extra {3}",))

    def test_validate_tag_placeholders_reports_missing_counts(self):
        from phraseloom.tag_engine import extract_tags, validate_tag_placeholders

        extraction = extract_tags("<a>x</a>")

        validation = validate_tag_placeholders("{1>x", extraction.tags)

        self.assertEqual(validation.warnings, ("protected_token_mismatch: missing <2}",))
```

- [ ] **Step 3: Update repeated known target serialization tests**

Replace `test_serialize_known_tags_preserves_repeated_raw_tags_in_order` with:

```python
    def test_serialize_known_tags_preserves_repeated_raw_spans_in_order(self):
        from phraseloom.tag_engine import TAG_SELF, TagToken, serialize_known_tags

        tags = (
            TagToken(1, TAG_SELF, "{1}", "<br/>"),
            TagToken(2, TAG_SELF, "{2}", "<br/>"),
        )

        result = serialize_known_tags("<br/> A <br/>", tags)

        self.assertEqual(result.text, "{1} A {2}")
        self.assertEqual(result.tags, tags)
        self.assertEqual(result.warnings, ())
```

- [ ] **Step 4: Run focused tests and verify failures**

Run:

```powershell
py -3 -m unittest tests.test_tag_engine -v
```

Expected: FAIL because `is_tag_only_segment`, `validate_tag_placeholders`, and warning messages still use old token counting.

- [ ] **Step 5: Update protected-only detection**

In `phraseloom/tag_engine.py`, replace `is_tag_only_segment` with:

```python
def is_tag_only_segment(source: str) -> bool:
    text = "" if source is None else str(source)
    stripped = text.strip()
    if not stripped:
        return False
    remainder = PROTECTED_TOKEN_RE.sub("", stripped)
    return remainder.strip() == ""
```

- [ ] **Step 6: Update validation to count exact protected token strings**

Replace `validate_tag_placeholders` with:

```python
def validate_tag_placeholders(text: str, tags: tuple[TagToken, ...]) -> TagValidation:
    source_counts = Counter(tag.placeholder for tag in tags)
    target_counts = Counter(
        found.group(0) for found in PROTECTED_TOKEN_RE.finditer("" if text is None else str(text))
    )
    warnings: list[str] = []

    for placeholder in sorted(source_counts, key=_protected_token_sort_key):
        missing = source_counts[placeholder] - target_counts[placeholder]
        warnings.extend(
            f"protected_token_mismatch: missing {placeholder}" for _ in range(missing)
        )

    for placeholder in sorted(target_counts, key=_protected_token_sort_key):
        extra = target_counts[placeholder] - source_counts[placeholder]
        warnings.extend(
            f"protected_token_mismatch: extra {placeholder}" for _ in range(extra)
        )

    return TagValidation(tuple(warnings))
```

Replace `_placeholder_sort_key` with:

```python
def _protected_token_sort_key(placeholder: str) -> tuple[int, int]:
    parsed = parse_protected_token(placeholder)
    if parsed is None:
        return 10**9, 99
    index, kind = parsed
    kind_order = {TAG_OPEN: 0, PROTECTED_SINGLE: 1, TAG_CLOSE: 2}
    return index, kind_order[kind]
```

- [ ] **Step 7: Update target serialization warning names**

In `serialize_known_tags`, replace:

```python
warnings.append(f"source_tag_not_found: {tag.raw}")
```

with:

```python
warnings.append(f"source_protected_span_not_found: {tag.raw}")
```

- [ ] **Step 8: Run tag engine tests and verify they pass**

Run:

```powershell
py -3 -m unittest tests.test_tag_engine -v
```

Expected: PASS.

- [ ] **Step 9: Commit validation and restoration changes**

Run:

```powershell
git add phraseloom/tag_engine.py tests/test_tag_engine.py
git commit -m "Validate and restore protected tokens"
```

Expected: commit succeeds.

---

### Task 4: Make Template Parsing Protected-Aware

**Files:**
- Modify: `tests/test_template_workflow.py`
- Modify: `phraseloom/template_engine.py`

- [ ] **Step 1: Replace the named placeholder template test**

In `tests/test_template_workflow.py`, replace `test_preserves_named_placeholders_and_uses_readable_numeric_names` with:

```python
    def test_template_parser_ignores_raw_brace_placeholders_after_protection(self):
        from phraseloom.tag_engine import extract_tags
        from phraseloom.template_engine import parse_template

        protected = extract_tags("Player reaches level {a}").text
        stage = parse_template("Clear Story 10-20")

        self.assertEqual(protected, "Player reaches level {1}")
        self.assertEqual(parse_template(protected).template, "Player reaches level {1}")
        self.assertEqual(parse_template(protected).values, {})
        self.assertEqual(stage.template, "Clear Story {stage1}")
        self.assertEqual(stage.values, {"stage1": "10-20"})
```

- [ ] **Step 2: Update the tag placeholder template test**

Replace `test_template_parser_preserves_tag_placeholders_without_values` with:

```python
    def test_template_parser_preserves_protected_tokens_without_values(self):
        from phraseloom.template_engine import parse_template

        match = parse_template("{1>VIP10 Pack<2} {3}")

        self.assertEqual(match.text, "{1>VIP10 Pack<2} {3}")
        self.assertEqual(match.template, "{1>VIP{num1} Pack<2} {3}")
        self.assertEqual(match.values, {"num1": "10"})
```

- [ ] **Step 3: Add inference protection test**

Add this test near the template engine tests in `TemplateDemoTests`:

```python
    def test_target_template_inference_does_not_replace_protected_token_digits(self):
        from phraseloom.template_engine import infer_target_template, parse_template

        match = parse_template("{1>Level 1<2}")

        self.assertEqual(match.template, "{1>Level {num1}<2}")
        self.assertEqual(match.values, {"num1": "1"})
        self.assertEqual(infer_target_template(match.values, "{1>Niveau 1<2}"), "{1>Niveau {num1}<2}")
```

- [ ] **Step 4: Run template tests and verify failures**

Run:

```powershell
py -3 -m unittest tests.test_template_workflow.TemplateDemoTests.test_template_parser_ignores_raw_brace_placeholders_after_protection tests.test_template_workflow.TemplateDemoTests.test_template_parser_preserves_protected_tokens_without_values tests.test_template_workflow.TemplateDemoTests.test_target_template_inference_does_not_replace_protected_token_digits -v
```

Expected: FAIL because `template_engine` still extracts raw `{...}` and `infer_target_template` replaces values inside protected tokens.

- [ ] **Step 5: Update variable regexes**

In `phraseloom/template_engine.py`, change imports and regexes at the top to:

```python
from .tag_engine import PROTECTED_TOKEN_RE

VAR_RE = re.compile(r"#[0-9A-Fa-f]{6}|\d+(?:[./:-]\d+)+|\d+(?:\.\d+)?")
PLACEHOLDER_RE = re.compile(r"\{[A-Za-z_][A-Za-z0-9_]*\}")
```

Remove `NAMED_PLACEHOLDER_RE`.

- [ ] **Step 6: Add protected span splitting helpers**

In `phraseloom/template_engine.py`, add these helpers before `parse_template`:

```python
def _iter_protected_aware_spans(source: str):
    pos = 0
    for found in PROTECTED_TOKEN_RE.finditer(source):
        if found.start() > pos:
            yield source[pos : found.start()], False
        yield found.group(0), True
        pos = found.end()
    if pos < len(source):
        yield source[pos:], False


def _replace_outside_protected_tokens(source: str, old: str, new: str) -> tuple[str, bool]:
    changed = False
    chunks: list[str] = []
    for chunk, protected in _iter_protected_aware_spans(source):
        if protected:
            chunks.append(chunk)
            continue
        replaced = chunk.replace(old, new)
        if replaced != chunk:
            changed = True
        chunks.append(replaced)
    return "".join(chunks), changed
```

- [ ] **Step 7: Rewrite `parse_template` around protected spans**

Replace the body of `parse_template` with:

```python
def parse_template(text: object) -> TemplateMatch:
    source = "" if text is None else str(text)
    chunks: list[str] = []
    values: dict[str, str] = {}
    counters: dict[str, int] = defaultdict(int)

    for span, protected in _iter_protected_aware_spans(source):
        if protected:
            chunks.append(span)
            continue
        pos = 0
        for found in VAR_RE.finditer(span):
            chunks.append(span[pos : found.start()])
            value = found.group(0)
            key = _variable_key(value, counters)
            chunks.append("{" + key + "}")
            values[key] = value
            pos = found.end()
        chunks.append(span[pos:])

    return TemplateMatch(source, "".join(chunks), values)
```

- [ ] **Step 8: Remove named placeholder key handling**

Replace `_variable_key` with:

```python
def _variable_key(value: str, counters: dict[str, int]) -> str:
    if value.startswith("#"):
        prefix = "color"
    elif re.fullmatch(r"\d+-\d+(?:-\d+)*", value):
        prefix = "stage"
    elif re.fullmatch(r"\d+(?:[./:]\d+)+", value):
        prefix = "seq"
    else:
        prefix = "num"
    counters[prefix] += 1
    return f"{prefix}{counters[prefix]}"
```

- [ ] **Step 9: Protect target-template inference**

In `infer_target_template`, replace:

```python
        if value in target_template:
            target_template = target_template.replace(value, token)
            tokens[token] = "{" + key + "}"
            matched = True
```

with:

```python
        target_template, replaced = _replace_outside_protected_tokens(
            target_template, value, token
        )
        if replaced:
            tokens[token] = "{" + key + "}"
            matched = True
```

- [ ] **Step 10: Update exports**

In `__all__`, remove `"NAMED_PLACEHOLDER_RE"` and keep `"PLACEHOLDER_RE"` and `"VAR_RE"`.

- [ ] **Step 11: Run template workflow tests**

Run:

```powershell
py -3 -m unittest tests.test_template_workflow -v
```

Expected: FAIL only in workflow tests that still expect old `{tN_*}` workbook text.

- [ ] **Step 12: Commit template changes**

Run:

```powershell
git add phraseloom/template_engine.py tests/test_template_workflow.py
git commit -m "Make template parsing protected-token aware"
```

Expected: commit succeeds after the three focused template parser tests pass. Broader workflow tests may still fail from old expected strings and are updated in Task 5.

---

### Task 5: Update Workflow Tests For Translator-Facing Protected Tokens

**Files:**
- Modify: `tests/test_template_workflow.py`
- Modify: `tests/test_tag_workflow_testfiles.py`
- Modify: `phraseloom/workflow.py`
- Modify: `phraseloom/excel_io.py`

- [ ] **Step 1: Update tag workflow expectations in `tests/test_template_workflow.py`**

In `test_tag_only_units_autofill_and_template_fill_restores_raw_tags`, change the assertions and written target unit to:

```python
            self.assertEqual(
                source_to_row["{1}"][target_idx - 1].value,
                "{1}",
            )
            self.assertEqual(
                source_to_row["{1>VIP{num1} Pack<2}"][variables_idx - 1].value,
                "{num1}=10,20",
            )
            source_to_row["{1>VIP{num1} Pack<2}"][target_idx - 1].value = (
                "{1>Pack VIP{num1}<2}"
            )
```

In `test_read_source_rows_serializes_source_and_existing_target_tags`, change expected row fields to:

```python
        self.assertEqual(rows[0].source, "{1>VIP10<2}")
        self.assertEqual(rows[0].existing_target, "{1>VIP10 Pack FR<2}")
        self.assertEqual(rows[0].match.template, "{1>VIP{num1}<2}")
```

- [ ] **Step 2: Add workflow coverage for raw brace placeholders**

Add this test to `TemplateDemoTests`:

```python
    def test_raw_brace_placeholders_are_translator_facing_protected_tokens(self):
        from phraseloom.workflow import fill_target_column_workbook, generate_workbook

        with tempfile.TemporaryDirectory() as tmp:
            input_path = Path(tmp) / "source.xlsx"
            pack_path = Path(tmp) / "pack.xlsx"
            filled_path = Path(tmp) / "filled.xlsx"

            wb = Workbook()
            ws = wb.active
            ws.append(["source", "target"])
            ws.append(["Hit deals {0} damage", ""])
            ws.append(["Hit deals {value} damage", ""])
            wb.save(input_path)

            generate_workbook(
                input_path,
                pack_path,
                source_col="source",
                target_col="target",
                min_group_size=2,
                use_existing_targets=False,
            )

            pack = load_workbook(pack_path)
            units = pack["translation_units"]
            headers = [cell.value for cell in units[1]]
            source_idx = headers.index("source_unit") + 1
            target_idx = headers.index("target_unit") + 1
            rows = {
                row[source_idx - 1].value: row
                for row in units.iter_rows(min_row=2)
            }

            self.assertIn("Hit deals {1} damage", rows)
            rows["Hit deals {1} damage"][target_idx - 1].value = "Inflige {1} degats"
            pack.save(pack_path)

            fill_target_column_workbook(
                input_path,
                filled_path,
                source_col="source",
                target_col="target",
                template_workbook=pack_path,
                min_group_size=2,
            )

            filled = load_workbook(filled_path, data_only=True)
            output_rows = list(filled.active.iter_rows(values_only=True))
            self.assertEqual(output_rows[1][1], "Inflige {0} degats")
            self.assertEqual(output_rows[2][1], "Inflige {value} degats")
```

- [ ] **Step 3: Update real-workbook tag workflow test expectations**

In `tests/test_tag_workflow_testfiles.py`, replace old source strings such as:

```python
"{t1_op}VIP{num1} Pack{t1_cl}"
```

with:

```python
"{1>VIP{num1} Pack<2}"
```

Replace old self-closing strings such as:

```python
"{t1_sf}"
```

with:

```python
"{1}"
```

Replace old mismatch warning checks that search for `"tag_mismatch"` with checks for `"protected_token_mismatch"`.

- [ ] **Step 4: Run workflow tests and verify failures**

Run:

```powershell
py -3 -m unittest tests.test_template_workflow tests.test_tag_workflow_testfiles -v
```

Expected: FAIL in places where production code still emits old QA labels, warning text, or variable summaries.

- [ ] **Step 5: Update workflow imports and QA warning names**

In `phraseloom/workflow.py`, keep the existing imported function names to avoid a broad rename, but update warning matching in `_workbook_stats` only if the code directly checks `"tag_mismatch:"`. The current warning merge can keep the function names because `validate_tag_placeholders` now returns `protected_token_mismatch`.

In `_unit_warning`, keep:

```python
source_placeholders = {
    placeholder
    for placeholder in PLACEHOLDER_RE.findall(source_unit)
    if not is_tag_placeholder(placeholder)
}
```

This continues to collect `{num1}` template variables and ignores protected tokens because `PLACEHOLDER_RE` only matches alphabetic placeholder names.

- [ ] **Step 6: Update QA labels in `excel_io.py`**

In `_write_output_workbook`, replace the QA rows:

```python
[
    "tag_mismatch_rows",
    sum(1 for result_row in result_rows if "tag_mismatch:" in result_row.warning),
]
```

with:

```python
[
    "protected_token_mismatch_rows",
    sum(
        1
        for result_row in result_rows
        if "protected_token_mismatch:" in result_row.warning
    ),
]
```

Replace:

```python
"tag_warning_rows"
"tag_only_units"
```

with:

```python
"protected_token_warning_rows"
"protected_only_units"
```

Do not change workbook sheet names or column names.

- [ ] **Step 7: Confirm variable summaries ignore protected tokens**

In `phraseloom/excel_io.py`, leave `_variables_summary` structurally unchanged and keep the existing compatibility import:

```python
from .tag_engine import extract_tags, is_tag_placeholder, serialize_known_tags
```

Keep the existing filter:

```python
if not is_tag_placeholder(placeholder)
```

`PLACEHOLDER_RE` only matches alphabetic template placeholders such as `{num1}`, so protected tokens such as `{1}`, `{1>`, and `<2}` are not included in the variables summary.

- [ ] **Step 8: Run workflow tests**

Run:

```powershell
py -3 -m unittest tests.test_template_workflow tests.test_tag_workflow_testfiles -v
```

Expected: PASS.

- [ ] **Step 9: Commit workflow updates**

Run:

```powershell
git add phraseloom/workflow.py phraseloom/excel_io.py tests/test_template_workflow.py tests/test_tag_workflow_testfiles.py
git commit -m "Update workflow for protected token workbooks"
```

Expected: commit succeeds.

---

### Task 6: Update Onboarding Documentation

**Files:**
- Modify: `agent.md`
- Modify: `docs/superpowers/specs/2026-05-11-protected-token-design.md`

- [ ] **Step 1: Update `agent.md` project purpose section**

In `agent.md`, replace the old tag placeholder paragraph:

```markdown
Tag extraction is now integrated into the main workflow as an internal
pre-template layer. It serializes recognized tags into protected placeholders
such as `{t1_op}`, `{t1_cl}`, and `{t2_sf}`, runs the normal template/TM flow on
that serialized text, then restores raw tags during fill.
```

with:

```markdown
Protected-token extraction is integrated into the main workflow as an internal
pre-template layer. It serializes recognized tags and every complete raw `{...}`
placeholder into translator-facing protected tokens such as `{1>`, `<2}`, and
`{3}`, runs the normal template/TM flow on that serialized text, then restores
raw spans during fill.
```

- [ ] **Step 2: Update `agent.md` engineering notes**

Replace the old note:

```markdown
- Tag placeholders are reserved as `{tN_op}`, `{tN_cl}`, and `{tN_sf}`. The tag
  extractor owns that namespace; template parsing must preserve these
  placeholders and must not include them in normal template variables.
```

with:

```markdown
- Protected tokens are reserved as `{N>`, `<N}`, and `{N}`. The protected-token
  extractor owns those tokens; template parsing must preserve them and must not
  include their numbers in normal template variables.
```

Replace:

```markdown
- Tag-only units should auto-fill with `target_unit_source = "tag_only"`.
```

with:

```markdown
- Protected-only units should auto-fill with `target_unit_source = "tag_only"`
  for workbook compatibility.
```

- [ ] **Step 3: Document compatibility helper names in the spec**

Add this sentence to the Compatibility section of `docs/superpowers/specs/2026-05-11-protected-token-design.md`:

```markdown
Some internal Python helper names may retain `tag_` wording for compatibility,
but their behavior follows the protected-token contract.
```

- [ ] **Step 4: Run documentation diff check**

Run:

```powershell
git diff --check
```

Expected: no output and exit code 0.

- [ ] **Step 5: Commit documentation updates**

Run:

```powershell
git add agent.md docs/superpowers/specs/2026-05-11-protected-token-design.md
git commit -m "Update docs for protected token contract"
```

Expected: commit succeeds.

---

### Task 7: Full Verification

**Files:**
- Verify: entire repository

- [ ] **Step 1: Run the full unit suite**

Run:

```powershell
py -3 -m unittest discover -v
```

Expected: all tests pass. The old onboarding note said 48 tests before this feature; the final count will be higher after new protected-token tests are added.

- [ ] **Step 2: Run the syntax check for changed modules**

Run:

```powershell
py -3 -m py_compile phraseloom/tag_engine.py phraseloom/template_engine.py phraseloom/workflow.py phraseloom/excel_io.py phraseloom/models.py
```

Expected: no output and exit code 0.

- [ ] **Step 3: Inspect final git status**

Run:

```powershell
git status --short
```

Expected: no modified tracked files. Ignored generated workbook output under `testfiles/` is acceptable only if it is ignored and not shown by `git status --short`.

- [ ] **Step 4: Record verification in the final response**

Final response should include:

```text
Implemented protected-token extraction using {N>, <N}, and {N}.
Updated template parsing so variables are extracted only outside protected tokens.
Verified with: py -3 -m unittest discover -v; py -3 -m py_compile ...
```

Do not claim entity clustering integration in this implementation; describe it as the next independent plan if the user asks for it.
