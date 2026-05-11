# Tag Extractor Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add PhraseLoom's internal tag extraction, tag placeholder preservation, tag-only auto-fill, and tag restoration without changing the user's three-step workflow.

**Architecture:** Add a pure `phraseloom.tag_engine` layer that serializes raw tags into `{tN_op}` / `{tN_cl}` / `{tN_sf}` placeholders before template parsing. Keep template parsing responsible only for non-tag variables, keep Excel I/O responsible for row serialization and workbook output, and let workflow orchestration apply template values before restoring row-level raw tags.

**Tech Stack:** Python 3.11+, standard-library `re`, `dataclasses`, `unittest`, `openpyxl`, existing PhraseLoom modules.

---

## File Structure

- Create `phraseloom/tag_engine.py`: pure tag scanning, placeholder helpers, tag-only checks, validation, and restoration.
- Modify `phraseloom/models.py`: add row tag metadata and a row fill result model so row-level tag warnings can reach workbook writers.
- Modify `phraseloom/template_engine.py`: preserve tag placeholders while parsing normal template variables.
- Modify `phraseloom/excel_io.py`: tag-serialize source and target text on read, write row-level warnings and QA counts, and accept row fill result objects.
- Modify `phraseloom/workflow.py`: auto-fill tag-only units, validate and restore tags after template application, and keep stats accurate.
- Create `tests/test_tag_engine.py`: focused pure tests for `tag_engine.py`.
- Create `tests/test_tag_workflow_testfiles.py`: workflow tests that create and use workbooks under `testfiles/`.

The `testfiles/` directory is intentionally ignored by git. Workflow tests in this plan create deterministic fixture workbooks there at runtime and clean up only generated `*_l10n/` output directories they own.

---

### Task 1: Add The Pure Tag Engine

**Files:**
- Create: `phraseloom/tag_engine.py`
- Test: `tests/test_tag_engine.py`

- [ ] **Step 1: Write failing placeholder helper tests**

Create `tests/test_tag_engine.py` with this initial content:

```python
import unittest


class TagEngineTests(unittest.TestCase):
    def test_placeholder_helpers_generate_and_parse_reserved_namespace(self):
        from phraseloom.tag_engine import (
            TAG_CLOSE,
            TAG_OPEN,
            TAG_SELF,
            is_tag_placeholder,
            make_tag_placeholder,
            parse_tag_placeholder,
        )

        self.assertEqual(make_tag_placeholder(1, TAG_OPEN), "{t1_op}")
        self.assertEqual(make_tag_placeholder(1, TAG_CLOSE), "{t1_cl}")
        self.assertEqual(make_tag_placeholder(2, TAG_SELF), "{t2_sf}")
        self.assertTrue(is_tag_placeholder("{t1_op}"))
        self.assertTrue(is_tag_placeholder("{t25_sf}"))
        self.assertFalse(is_tag_placeholder("{num1}"))
        self.assertFalse(is_tag_placeholder("{tag1_op}"))
        self.assertEqual(parse_tag_placeholder("{t1_op}"), (1, TAG_OPEN))
        self.assertEqual(parse_tag_placeholder("{t2_sf}"), (2, TAG_SELF))
        self.assertIsNone(parse_tag_placeholder("{num1}"))


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run the helper test to verify it fails**

Run:

```powershell
py -3 -m unittest tests.test_tag_engine.TagEngineTests.test_placeholder_helpers_generate_and_parse_reserved_namespace -v
```

Expected: FAIL with `ModuleNotFoundError: No module named 'phraseloom.tag_engine'`.

- [ ] **Step 3: Add minimal placeholder helpers**

Create `phraseloom/tag_engine.py` with:

```python
from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass

TAG_PLACEHOLDER_PREFIX = "t"
TAG_OPEN = "op"
TAG_CLOSE = "cl"
TAG_SELF = "sf"

TAG_PLACEHOLDER_RE = re.compile(r"\{t([1-9][0-9]*)_(op|cl|sf)\}")


@dataclass(frozen=True)
class TagToken:
    index: int
    kind: str
    placeholder: str
    raw: str


@dataclass(frozen=True)
class TagExtraction:
    text: str
    tags: tuple[TagToken, ...]
    warnings: tuple[str, ...] = ()


@dataclass(frozen=True)
class TagValidation:
    warnings: tuple[str, ...]


def make_tag_placeholder(index: int, kind: str) -> str:
    return "{" + TAG_PLACEHOLDER_PREFIX + str(index) + "_" + kind + "}"


def parse_tag_placeholder(value: str) -> tuple[int, str] | None:
    match = TAG_PLACEHOLDER_RE.fullmatch(value)
    if not match:
        return None
    return int(match.group(1)), match.group(2)


def is_tag_placeholder(value: str) -> bool:
    return parse_tag_placeholder(value) is not None


__all__ = [
    "TAG_CLOSE",
    "TAG_OPEN",
    "TAG_PLACEHOLDER_PREFIX",
    "TAG_PLACEHOLDER_RE",
    "TAG_SELF",
    "TagExtraction",
    "TagToken",
    "TagValidation",
    "is_tag_placeholder",
    "make_tag_placeholder",
    "parse_tag_placeholder",
]
```

- [ ] **Step 4: Run the helper test to verify it passes**

Run:

```powershell
py -3 -m unittest tests.test_tag_engine.TagEngineTests.test_placeholder_helpers_generate_and_parse_reserved_namespace -v
```

Expected: PASS.

- [ ] **Step 5: Add failing extraction, restore, and validation tests**

Append these methods inside `TagEngineTests`:

```python
    def test_extracts_angle_and_bbcode_tags_conservatively(self):
        from phraseloom.tag_engine import extract_tags

        html = extract_tags('<a href="shop">VIP10</a> <img src="coin.png"/>')
        self.assertEqual(html.text, "{t1_op}VIP10{t1_cl} {t2_sf}")
        self.assertEqual(
            [(tag.index, tag.kind, tag.placeholder, tag.raw) for tag in html.tags],
            [
                (1, "op", "{t1_op}", '<a href="shop">'),
                (1, "cl", "{t1_cl}", "</a>"),
                (2, "sf", "{t2_sf}", '<img src="coin.png"/>'),
            ],
        )
        self.assertEqual(html.warnings, ())

        bbcode = extract_tags("[color=#ff0]Bonus[/]")
        self.assertEqual(bbcode.text, "{t1_op}Bonus{t1_cl}")
        self.assertEqual(
            [(tag.index, tag.kind, tag.placeholder, tag.raw) for tag in bbcode.tags],
            [
                (1, "op", "{t1_op}", "[color=#ff0]"),
                (1, "cl", "{t1_cl}", "[/]"),
            ],
        )

    def test_preserves_suspicious_text_and_reports_namespace_conflicts(self):
        from phraseloom.tag_engine import extract_tags

        result = extract_tags("Use 3 < 5 and keep {t1_op}")

        self.assertEqual(result.text, "Use 3 < 5 and keep {t1_op}")
        self.assertEqual(result.tags, ())
        self.assertIn("reserved tag placeholder namespace", "; ".join(result.warnings))

    def test_detects_tag_only_segments(self):
        from phraseloom.tag_engine import is_tag_only_segment

        self.assertTrue(is_tag_only_segment("{t1_sf}"))
        self.assertTrue(is_tag_only_segment("{t1_op}{t1_cl}"))
        self.assertTrue(is_tag_only_segment("{t1_op} {t1_cl}"))
        self.assertFalse(is_tag_only_segment("{t1_op}Click{t1_cl}"))
        self.assertFalse(is_tag_only_segment("{t1_sf} 100 coins"))

    def test_restores_known_tags_and_leaves_unknown_placeholders_visible(self):
        from phraseloom.tag_engine import extract_tags, restore_tags, validate_tag_placeholders

        extraction = extract_tags('<a href="shop">VIP10</a>')
        restored = restore_tags("{t1_op}VIP11{t1_cl}{t2_sf}", extraction.tags)
        validation = validate_tag_placeholders("{t1_op}VIP11{t1_cl}{t2_sf}", extraction.tags)

        self.assertEqual(restored, '<a href="shop">VIP11</a>{t2_sf}')
        self.assertEqual(validation.warnings, ("tag_mismatch: extra {t2_sf}",))
```

- [ ] **Step 6: Run the expanded tag engine tests to verify they fail**

Run:

```powershell
py -3 -m unittest tests.test_tag_engine -v
```

Expected: FAIL with missing functions such as `extract_tags`, `is_tag_only_segment`, `restore_tags`, and `validate_tag_placeholders`.

- [ ] **Step 7: Implement extraction, tag-only checks, validation, and restore**

Replace `phraseloom/tag_engine.py` with:

```python
from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass

TAG_PLACEHOLDER_PREFIX = "t"
TAG_OPEN = "op"
TAG_CLOSE = "cl"
TAG_SELF = "sf"

TAG_PLACEHOLDER_RE = re.compile(r"\{t([1-9][0-9]*)_(op|cl|sf)\}")
ANGLE_SELF_RE = re.compile(r"<[A-Za-z][A-Za-z0-9:_-]*(?:\s+[^<>]*)?/>")
ANGLE_OPEN_RE = re.compile(r"<[A-Za-z][A-Za-z0-9:_-]*(?:\s+[^<>]*)?>")
ANGLE_CLOSE_RE = re.compile(r"</[A-Za-z][A-Za-z0-9:_-]*>")
ANGLE_SHORT_CLOSE_RE = re.compile(r"</>")
BBCODE_OPEN_RE = re.compile(r"\[[A-Za-z][A-Za-z0-9:_-]*(?:=[^\[\]]+)?\]")
BBCODE_CLOSE_RE = re.compile(r"\[/\]")
RESERVED_NAMESPACE_RE = re.compile(r"\{t[0-9]+_(?:op|cl|sf)\}")


@dataclass(frozen=True)
class TagToken:
    index: int
    kind: str
    placeholder: str
    raw: str


@dataclass(frozen=True)
class TagExtraction:
    text: str
    tags: tuple[TagToken, ...]
    warnings: tuple[str, ...] = ()


@dataclass(frozen=True)
class TagValidation:
    warnings: tuple[str, ...]


def make_tag_placeholder(index: int, kind: str) -> str:
    return "{" + TAG_PLACEHOLDER_PREFIX + str(index) + "_" + kind + "}"


def parse_tag_placeholder(value: str) -> tuple[int, str] | None:
    match = TAG_PLACEHOLDER_RE.fullmatch(value)
    if not match:
        return None
    return int(match.group(1)), match.group(2)


def is_tag_placeholder(value: str) -> bool:
    return parse_tag_placeholder(value) is not None


def extract_tags(text: object) -> TagExtraction:
    source = "" if text is None else str(text)
    warnings: list[str] = []
    if RESERVED_NAMESPACE_RE.search(source):
        warnings.append("reserved tag placeholder namespace appears in raw text")

    chunks: list[str] = []
    tags: list[TagToken] = []
    open_stack: list[int] = []
    next_index = 1
    pos = 0

    while pos < len(source):
        match, kind = _match_tag(source, pos)
        if not match:
            chunks.append(source[pos])
            pos += 1
            continue

        raw = match.group(0)
        if kind == TAG_SELF:
            index = next_index
            next_index += 1
        elif kind == TAG_OPEN:
            index = next_index
            next_index += 1
            open_stack.append(index)
        else:
            if not open_stack:
                chunks.append(raw)
                warnings.append(f"unpaired close tag left raw: {raw}")
                pos = match.end()
                continue
            index = open_stack.pop()

        placeholder = make_tag_placeholder(index, kind)
        chunks.append(placeholder)
        tags.append(TagToken(index=index, kind=kind, placeholder=placeholder, raw=raw))
        pos = match.end()

    for index in open_stack:
        warnings.append(f"open tag has no close partner: {make_tag_placeholder(index, TAG_OPEN)}")

    return TagExtraction("".join(chunks), tuple(tags), tuple(warnings))


def _match_tag(source: str, pos: int) -> tuple[re.Match[str], str] | tuple[None, None]:
    for pattern, kind in (
        (ANGLE_SELF_RE, TAG_SELF),
        (ANGLE_CLOSE_RE, TAG_CLOSE),
        (ANGLE_SHORT_CLOSE_RE, TAG_CLOSE),
        (ANGLE_OPEN_RE, TAG_OPEN),
        (BBCODE_CLOSE_RE, TAG_CLOSE),
        (BBCODE_OPEN_RE, TAG_OPEN),
    ):
        match = pattern.match(source, pos)
        if match:
            return match, kind
    return None, None


def is_tag_only_segment(source: str) -> bool:
    compact = re.sub(r"\s+", "", source)
    if not compact:
        return False
    remainder = TAG_PLACEHOLDER_RE.sub("", compact)
    return remainder == ""


def restore_tags(text: str, tags: tuple[TagToken, ...]) -> str:
    result = text
    raw_by_placeholder = {tag.placeholder: tag.raw for tag in tags}
    for placeholder, raw in raw_by_placeholder.items():
        result = result.replace(placeholder, raw)
    return result


def validate_tag_placeholders(text: str, tags: tuple[TagToken, ...]) -> TagValidation:
    expected = Counter(tag.placeholder for tag in tags)
    actual = Counter(match.group(0) for match in TAG_PLACEHOLDER_RE.finditer(text))
    missing: list[str] = []
    extra: list[str] = []
    for placeholder, count in expected.items():
        missing.extend([placeholder] * max(count - actual[placeholder], 0))
    for placeholder, count in actual.items():
        extra.extend([placeholder] * max(count - expected[placeholder], 0))
    warnings: list[str] = []
    if missing:
        warnings.append("tag_mismatch: missing " + ", ".join(missing))
    if extra:
        warnings.append("tag_mismatch: extra " + ", ".join(extra))
    return TagValidation(tuple(warnings))


def serialize_known_tags(text: object, tags: tuple[TagToken, ...]) -> TagExtraction:
    result = "" if text is None else str(text)
    if not result:
        return TagExtraction("", (), ())
    warnings: list[str] = []
    used: list[TagToken] = []
    for tag in sorted(tags, key=lambda item: len(item.raw), reverse=True):
        if tag.raw in result:
            result = result.replace(tag.raw, tag.placeholder)
            used.append(tag)
        else:
            warnings.append(f"source tag not found in target: {tag.placeholder}")
    return TagExtraction(result, tuple(used), tuple(warnings))


__all__ = [
    "TAG_CLOSE",
    "TAG_OPEN",
    "TAG_PLACEHOLDER_PREFIX",
    "TAG_PLACEHOLDER_RE",
    "TAG_SELF",
    "TagExtraction",
    "TagToken",
    "TagValidation",
    "extract_tags",
    "is_tag_only_segment",
    "is_tag_placeholder",
    "make_tag_placeholder",
    "parse_tag_placeholder",
    "restore_tags",
    "serialize_known_tags",
    "validate_tag_placeholders",
]
```

- [ ] **Step 8: Run tag engine tests to verify they pass**

Run:

```powershell
py -3 -m unittest tests.test_tag_engine -v
```

Expected: all tests in `tests.test_tag_engine` pass.

- [ ] **Step 9: Commit the pure tag engine**

Run:

```powershell
git add phraseloom/tag_engine.py tests/test_tag_engine.py
git commit -m "Add pure tag extraction engine"
```

---

### Task 2: Keep Tag Placeholders Out Of Template Variables

**Files:**
- Modify: `phraseloom/template_engine.py`
- Test: `tests/test_template_workflow.py`

- [ ] **Step 1: Add failing template parser regression**

Append this test method to `TemplateDemoTests` in `tests/test_template_workflow.py`:

```python
    def test_template_parser_preserves_tag_placeholders_without_values(self):
        from phraseloom.template_engine import parse_template

        match = parse_template("{t1_op}VIP10 Pack{t1_cl}")

        self.assertEqual(match.text, "{t1_op}VIP10 Pack{t1_cl}")
        self.assertEqual(match.template, "{t1_op}VIP{num1} Pack{t1_cl}")
        self.assertEqual(match.values, {"num1": "10"})
```

- [ ] **Step 2: Run the regression to verify it fails**

Run:

```powershell
py -3 -m unittest tests.test_template_workflow.TemplateDemoTests.test_template_parser_preserves_tag_placeholders_without_values -v
```

Expected: FAIL because `match.values` contains `t1_op` and `t1_cl`.

- [ ] **Step 3: Exclude tag placeholders from `VAR_RE` handling**

Modify `phraseloom/template_engine.py`:

```python
from .tag_engine import is_tag_placeholder
```

Then replace the loop body in `parse_template()` with this version:

```python
    for found in VAR_RE.finditer(source):
        value = found.group(0)
        if is_tag_placeholder(value):
            continue
        chunks.append(source[pos : found.start()])
        key = _variable_key(value, counters)
        chunks.append("{" + key + "}")
        values[key] = value
        pos = found.end()
```

The `continue` intentionally leaves `pos` unchanged so the final literal slice
keeps the tag placeholder in place.

- [ ] **Step 4: Run template parser tests**

Run:

```powershell
py -3 -m unittest tests.test_template_workflow.TemplateDemoTests.test_template_parser_preserves_tag_placeholders_without_values -v
```

Expected: PASS.

- [ ] **Step 5: Run existing template workflow regression tests**

Run:

```powershell
py -3 -m unittest tests.test_template_workflow -v
```

Expected: all `tests.test_template_workflow` tests pass.

- [ ] **Step 6: Commit template placeholder isolation**

Run:

```powershell
git add phraseloom/template_engine.py tests/test_template_workflow.py
git commit -m "Preserve tag placeholders during template parsing"
```

---

### Task 3: Carry Tag Metadata Through Rows And Fill Results

**Files:**
- Modify: `phraseloom/models.py`
- Modify: `phraseloom/excel_io.py`
- Modify: `phraseloom/workflow.py`
- Test: `tests/test_template_workflow.py`

- [ ] **Step 1: Add failing model compatibility test**

Append this test method to `TemplateDemoTests` in `tests/test_template_workflow.py`:

```python
    def test_row_item_carries_optional_tag_metadata(self):
        from phraseloom.models import RowFillResult, RowItem
        from phraseloom.template_engine import parse_template

        row = RowItem(2, "VIP{num1}", "", parse_template("VIP10"), ("VIP10",))
        result = RowFillResult(row=row, unit=None, auto_target=None, warning="tag warning")

        self.assertEqual(row.raw_source, "")
        self.assertEqual(row.tag_tokens, ())
        self.assertEqual(row.tag_warnings, ())
        self.assertEqual(result.warning, "tag warning")
```

- [ ] **Step 2: Run the model compatibility test to verify it fails**

Run:

```powershell
py -3 -m unittest tests.test_template_workflow.TemplateDemoTests.test_row_item_carries_optional_tag_metadata -v
```

Expected: FAIL because `RowFillResult` and the new `RowItem` fields do not exist.

- [ ] **Step 3: Extend models**

Modify `phraseloom/models.py`:

```python
from .tag_engine import TagToken
```

Replace the `RowItem` class with:

```python
@dataclass(frozen=True)
class RowItem:
    row_number: int
    source: str
    existing_target: str
    match: TemplateMatch
    original_values: tuple[object, ...]
    raw_source: str = ""
    raw_existing_target: str = ""
    tag_tokens: tuple[TagToken, ...] = ()
    tag_warnings: tuple[str, ...] = ()
    target_tag_warnings: tuple[str, ...] = ()
```

Add this class after `TranslationUnit`:

```python
@dataclass(frozen=True)
class RowFillResult:
    row: RowItem
    unit: TranslationUnit | None
    auto_target: str | None
    warning: str = ""
```

Add `"RowFillResult"` to `__all__`.

- [ ] **Step 4: Update writer signatures to accept `RowFillResult`**

In `phraseloom/excel_io.py`, import the new model:

```python
from .models import RowFillResult, RowItem, TranslationUnit
```

Change these signatures:

```python
def _write_output_workbook(
    output_path: Path,
    input_path: Path,
    units: list[TranslationUnit],
    result_rows: list[RowFillResult],
) -> None:
```

```python
def _write_target_column_workbook(
    output_path: Path,
    input_path: Path,
    target_col: str | int,
    result_rows: list[RowFillResult],
) -> None:
```

Then replace tuple loop unpacking in `_write_output_workbook()` with object access. Use these exact loop forms:

```python
    filled_rows = sum(1 for result in result_rows if result.auto_target)
```

```python
    for result_row in result_rows:
        row = result_row.row
        unit = result_row.unit
        auto_target = result_row.auto_target
        row_warning = result_row.warning
```

For warning selection inside `source_map` and `filled_workbook`, use:

```python
        if unit is None:
            warning = _merge_warnings("no translation unit", row_warning)
            fill_status = "unit_not_found"
        elif not unit.target_unit:
            warning = _merge_warnings("fill target_unit in to_translate, then rerun fill", row_warning)
            fill_status = "missing_target_unit"
        else:
            warning = _merge_warnings(unit.warning, row_warning)
            fill_status = "filled"
```

Add this helper near `_variables_summary()`:

```python
def _merge_warnings(*warnings: str) -> str:
    return "; ".join(warning for warning in warnings if warning)
```

In `_write_target_column_workbook()`, replace:

```python
        for row, _, auto_target in result_rows:
            if auto_target:
                ws.cell(row=row.row_number, column=target_index).value = auto_target
```

with:

```python
        for result_row in result_rows:
            if result_row.auto_target:
                ws.cell(row=result_row.row.row_number, column=target_index).value = result_row.auto_target
```

- [ ] **Step 5: Update workflow result rows to use `RowFillResult`**

In `phraseloom/workflow.py`, import:

```python
from .models import RowFillResult, RowItem, TranslationUnit
```

Change the `_build_fill_context()` return annotation to:

```python
) -> tuple[
    list[RowItem],
    list[TranslationUnit],
    list[RowFillResult],
    int,
]:
```

In `_build_fill_context()`, replace:

```python
        result_rows.append((row, unit, auto_target))
```

with:

```python
        result_rows.append(RowFillResult(row=row, unit=unit, auto_target=auto_target))
```

Update `_write_output_workbook()` and `_write_target_column_workbook()` call sites only through the changed `result_rows` variable; the call shape stays the same.

- [ ] **Step 6: Update stats and QA tuple comprehensions**

In `phraseloom/excel_io.py`, replace tuple comprehensions:

```python
sum(1 for _, unit, auto in result_rows if unit and not unit.target_unit)
sum(1 for _, _, auto in result_rows if auto)
```

with:

```python
sum(1 for result in result_rows if result.unit and not result.unit.target_unit)
sum(1 for result in result_rows if result.auto_target)
```

- [ ] **Step 7: Run model and existing workflow tests**

Run:

```powershell
py -3 -m unittest tests.test_template_workflow -v
```

Expected: all `tests.test_template_workflow` tests pass.

- [ ] **Step 8: Commit row metadata plumbing**

Run:

```powershell
git add phraseloom/models.py phraseloom/excel_io.py phraseloom/workflow.py tests/test_template_workflow.py
git commit -m "Carry row-level fill warnings"
```

---

### Task 4: Tag-Serialize Source And Existing Target On Read

**Files:**
- Modify: `phraseloom/excel_io.py`
- Test: `tests/test_template_workflow.py`

- [ ] **Step 1: Add failing source row serialization test**

Append this test method to `TemplateDemoTests` in `tests/test_template_workflow.py`:

```python
    def test_read_source_rows_serializes_source_and_existing_target_tags(self):
        from phraseloom.excel_io import _read_source_rows

        with tempfile.TemporaryDirectory() as tmp:
            input_path = Path(tmp) / "tagged.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.append(["source", "target"])
            ws.append(['<a href="shop">VIP10</a>', '<a href="shop">VIP10 Pack FR</a>'])
            wb.save(input_path)

            rows = _read_source_rows(input_path, "source", "target")

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0].raw_source, '<a href="shop">VIP10</a>')
        self.assertEqual(rows[0].source, "{t1_op}VIP10{t1_cl}")
        self.assertEqual(rows[0].existing_target, "{t1_op}VIP10 Pack FR{t1_cl}")
        self.assertEqual(rows[0].match.template, "{t1_op}VIP{num1}{t1_cl}")
        self.assertEqual(rows[0].match.values, {"num1": "10"})
        self.assertEqual(rows[0].tag_warnings, ())
```

- [ ] **Step 2: Run the serialization test to verify it fails**

Run:

```powershell
py -3 -m unittest tests.test_template_workflow.TemplateDemoTests.test_read_source_rows_serializes_source_and_existing_target_tags -v
```

Expected: FAIL because `_read_source_rows()` still uses raw source and raw target.

- [ ] **Step 3: Wire tag extraction into `_read_source_rows()`**

In `phraseloom/excel_io.py`, add:

```python
from .tag_engine import extract_tags, serialize_known_tags
```

Replace the source/target construction block in `_read_source_rows()` with:

```python
            raw_source = str(source_value).strip()
            source_extraction = extract_tags(raw_source)
            target_value = _cell_value(row, target_index) if target_index else ""
            raw_existing_target = "" if target_value is None else str(target_value).strip()
            target_extraction = serialize_known_tags(
                raw_existing_target, source_extraction.tags
            )
            rows.append(
                RowItem(
                    row_number,
                    source_extraction.text,
                    target_extraction.text,
                    parse_template(source_extraction.text),
                    tuple(row),
                    raw_source=raw_source,
                    raw_existing_target=raw_existing_target,
                    tag_tokens=source_extraction.tags,
                    tag_warnings=source_extraction.warnings,
                    target_tag_warnings=target_extraction.warnings,
                )
            )
```

- [ ] **Step 4: Run row serialization tests**

Run:

```powershell
py -3 -m unittest tests.test_template_workflow.TemplateDemoTests.test_read_source_rows_serializes_source_and_existing_target_tags -v
```

Expected: PASS.

- [ ] **Step 5: Run tag engine and workflow tests together**

Run:

```powershell
py -3 -m unittest tests.test_tag_engine tests.test_template_workflow -v
```

Expected: both modules pass.

- [ ] **Step 6: Commit row tag serialization**

Run:

```powershell
git add phraseloom/excel_io.py tests/test_template_workflow.py
git commit -m "Serialize row tags before template parsing"
```

---

### Task 5: Auto-Fill Tag-Only Units And Restore Tags In Workflow

**Files:**
- Modify: `phraseloom/workflow.py`
- Modify: `phraseloom/excel_io.py`
- Test: `tests/test_template_workflow.py`

- [ ] **Step 1: Add failing tag-only unit and restore test**

Append this test method to `TemplateDemoTests` in `tests/test_template_workflow.py`:

```python
    def test_tag_only_units_autofill_and_template_fill_restores_raw_tags(self):
        from phraseloom.workflow import fill_target_column_workbook, generate_workbook

        with tempfile.TemporaryDirectory() as tmp:
            input_path = Path(tmp) / "source.xlsx"
            pack_path = Path(tmp) / "pack.xlsx"
            filled_path = Path(tmp) / "filled.xlsx"

            wb = Workbook()
            ws = wb.active
            ws.append(["source", "target"])
            ws.append(['<img src="coin.png"/>', ""])
            ws.append(['<a href="shop">VIP10 Pack</a>', ""])
            ws.append(['<a href="shop">VIP20 Pack</a>', ""])
            wb.save(input_path)

            stats = generate_workbook(
                input_path,
                pack_path,
                source_col="source",
                target_col="target",
                min_group_size=2,
                use_existing_targets=False,
            )

            self.assertEqual(stats["prefilled_translation_unit_count"], 1)
            self.assertEqual(stats["untranslated_translation_unit_count"], 1)

            pack = load_workbook(pack_path)
            units = pack["translation_units"]
            headers = [cell.value for cell in units[1]]
            source_idx = headers.index("source_unit") + 1
            target_idx = headers.index("target_unit") + 1
            variables_idx = headers.index("variables") + 1
            source_to_row = {
                row[source_idx - 1].value: row
                for row in units.iter_rows(min_row=2)
            }
            self.assertEqual(
                source_to_row["{t1_sf}"][target_idx - 1].value,
                "{t1_sf}",
            )
            self.assertEqual(
                source_to_row["{t1_op}VIP{num1} Pack{t1_cl}"][variables_idx - 1].value,
                "{num1}=10,20",
            )
            source_to_row["{t1_op}VIP{num1} Pack{t1_cl}"][target_idx - 1].value = (
                "{t1_op}Pack VIP{num1}{t1_cl}"
            )
            pack.save(pack_path)

            fill_stats = fill_target_column_workbook(
                input_path,
                filled_path,
                source_col="source",
                target_col="target",
                template_workbook=pack_path,
                min_group_size=2,
            )

            self.assertEqual(fill_stats["autofilled_count"], 3)
            filled = load_workbook(filled_path, data_only=True)
            rows = list(filled.active.iter_rows(values_only=True))
            self.assertEqual(rows[1][1], '<img src="coin.png"/>')
            self.assertEqual(rows[2][1], '<a href="shop">Pack VIP10</a>')
            self.assertEqual(rows[3][1], '<a href="shop">Pack VIP20</a>')
```

- [ ] **Step 2: Run the workflow test to verify it fails**

Run:

```powershell
py -3 -m unittest tests.test_template_workflow.TemplateDemoTests.test_tag_only_units_autofill_and_template_fill_restores_raw_tags -v
```

Expected: FAIL because tag-only segments are not auto-filled and fill writes serialized placeholders.

- [ ] **Step 3: Import tag helpers into workflow**

In `phraseloom/workflow.py`, add:

```python
from .tag_engine import (
    is_tag_placeholder,
    is_tag_only_segment,
    restore_tags,
    validate_tag_placeholders,
)
```

- [ ] **Step 4: Auto-fill tag-only segments before generic non-translatable segments**

In `_build_translation_units()`, replace:

```python
        elif _is_non_translatable_segment(source_unit):
            target_unit = source_unit
            target_unit_source = "non_translatable"
```

with:

```python
        elif is_tag_only_segment(source_unit):
            target_unit = source_unit
            target_unit_source = "tag_only"
        elif _is_non_translatable_segment(source_unit):
            target_unit = source_unit
            target_unit_source = "non_translatable"
```

- [ ] **Step 5: Validate and restore tags after template application**

In `_build_fill_context()`, replace auto target construction:

```python
        auto_target = (
            apply_target_template(target_template, row.match.values)
            if target_template and unit and unit.unit_type == "template"
            else target_template
            if target_template
            else None
        )
        if auto_target:
            autofilled_count += 1
        result_rows.append(RowFillResult(row=row, unit=unit, auto_target=auto_target))
```

with:

```python
        serialized_target = (
            apply_target_template(target_template, row.match.values)
            if target_template and unit and unit.unit_type == "template"
            else target_template
            if target_template
            else None
        )
        row_warning = ""
        auto_target = None
        if serialized_target:
            validation = validate_tag_placeholders(serialized_target, row.tag_tokens)
            row_warning = _merge_warning_parts(
                *row.tag_warnings,
                *row.target_tag_warnings,
                *validation.warnings,
            )
            auto_target = restore_tags(serialized_target, row.tag_tokens)
        elif row.tag_warnings or row.target_tag_warnings:
            row_warning = _merge_warning_parts(*row.tag_warnings, *row.target_tag_warnings)
        if auto_target:
            autofilled_count += 1
        result_rows.append(
            RowFillResult(row=row, unit=unit, auto_target=auto_target, warning=row_warning)
        )
```

Add this helper near `_format_rate()`:

```python
def _merge_warning_parts(*warnings: str) -> str:
    return "; ".join(warning for warning in warnings if warning)
```

- [ ] **Step 6: Add QA counts for tag rows**

In `phraseloom/excel_io.py`, add QA rows in `_write_output_workbook()` after `warning_units`:

```python
    qa.append(["tag_mismatch_rows", sum(1 for result in result_rows if "tag_mismatch" in result.warning)])
    qa.append(["tag_warning_rows", sum(1 for result in result_rows if result.warning)])
    qa.append(["tag_only_units", sum(1 for unit in units if unit.target_unit_source == "tag_only")])
```

No schema constant is required for these QA keys because existing QA rows use string keys.

- [ ] **Step 7: Keep tag placeholders out of variable summaries and template variable warnings**

In `phraseloom/excel_io.py`, import:

```python
from .tag_engine import extract_tags, is_tag_placeholder, serialize_known_tags
```

Then replace the placeholder collection in `_variables_summary()` with:

```python
    placeholders = [
        placeholder
        for placeholder in PLACEHOLDER_RE.findall(unit.source_unit)
        if not is_tag_placeholder(placeholder)
    ]
```

In `phraseloom/workflow.py`, replace the placeholder set construction in `_unit_warning()` with:

```python
    source_placeholders = {
        placeholder
        for placeholder in PLACEHOLDER_RE.findall(source_unit)
        if not is_tag_placeholder(placeholder)
    }
    target_placeholders = {
        placeholder
        for placeholder in PLACEHOLDER_RE.findall(target_unit)
        if not is_tag_placeholder(placeholder)
    }
```

- [ ] **Step 8: Run the tag-only restore test**

Run:

```powershell
py -3 -m unittest tests.test_template_workflow.TemplateDemoTests.test_tag_only_units_autofill_and_template_fill_restores_raw_tags -v
```

Expected: PASS.

- [ ] **Step 9: Run all current tests**

Run:

```powershell
py -3 -m unittest discover -v
```

Expected: all discovered tests pass.

- [ ] **Step 10: Commit workflow restore and tag-only auto-fill**

Run:

```powershell
git add phraseloom/workflow.py phraseloom/excel_io.py tests/test_template_workflow.py
git commit -m "Restore tags after template fill"
```

---

### Task 6: Add Testfiles-Based End-To-End Workflow Tests

**Files:**
- Create: `tests/test_tag_workflow_testfiles.py`

- [ ] **Step 1: Create failing end-to-end tests that use `testfiles/`**

Create `tests/test_tag_workflow_testfiles.py`:

```python
import shutil
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook


TESTFILES = Path("testfiles")
TAG_TM = TESTFILES / "tag_tm.xlsx"
TAG_SOURCE = TESTFILES / "tag_source.xlsx"


def ensure_tag_testfiles() -> None:
    TESTFILES.mkdir(exist_ok=True)

    tm = Workbook()
    ws = tm.active
    ws.append(["source", "target"])
    ws.append(['<a href="shop">VIP10 Pack</a>', '<a href="shop">Pack VIP10 FR</a>'])
    ws.append(['<a href="shop">VIP20 Pack</a>', '<a href="shop">Pack VIP20 FR</a>'])
    ws.append(["Login failed", "Login failed FR"])
    tm.save(TAG_TM)

    source = Workbook()
    ws = source.active
    ws.append(["source", "target"])
    ws.append(['<a href="shop">VIP30 Pack</a>', ""])
    ws.append(['<a href="shop">VIP40 Pack</a>', ""])
    ws.append(['<img src="coin.png"/>', ""])
    ws.append(["Brand new line", ""])
    source.save(TAG_SOURCE)


def clean_generated_outputs() -> None:
    for folder in (TESTFILES / "tag_tm_l10n", TESTFILES / "tag_source_l10n"):
        if folder.exists():
            shutil.rmtree(folder)


class TagWorkflowTestfilesTests(unittest.TestCase):
    def setUp(self):
        ensure_tag_testfiles()
        clean_generated_outputs()

    def tearDown(self):
        clean_generated_outputs()

    def test_tm_extract_and_fill_use_testfiles_with_tags(self):
        from phraseloom.workflow import (
            fill_target_column_workbook,
            generate_tm_pairs,
            generate_workbook,
        )

        tm_pairs = TESTFILES / "tag_tm_l10n" / "tag_tm_reusable_units.xlsx"
        pack = TESTFILES / "tag_source_l10n" / "tag_source_tm_prefill_pack.xlsx"
        filled = TESTFILES / "tag_source_l10n" / "tag_source_filled_result.xlsx"

        tm_stats = generate_tm_pairs(
            TAG_TM,
            tm_pairs,
            source_col="source",
            target_col="target",
            min_group_size=2,
        )
        self.assertEqual(tm_stats["template_pair_count"], 1)

        stats = generate_workbook(
            TAG_SOURCE,
            pack,
            source_col="source",
            target_col="target",
            tm_workbook=tm_pairs,
            min_group_size=2,
            use_existing_targets=False,
        )
        self.assertEqual(stats["prefilled_translation_unit_count"], 2)
        self.assertEqual(stats["untranslated_translation_unit_count"], 1)

        standalone_todo = TESTFILES / "tag_source_l10n" / "tag_source_translator_todo.xlsx"
        todo_book = load_workbook(standalone_todo)
        ws = todo_book["to_translate"]
        headers = [cell.value for cell in ws[1]]
        source_idx = headers.index("source_unit")
        target_idx = headers.index("target_unit") + 1
        todo_rows = list(ws.iter_rows(min_row=2))
        self.assertEqual([row[source_idx].value for row in todo_rows], ["Brand new line"])
        todo_rows[0][target_idx - 1].value = "Brand new line FR"
        todo_book.save(standalone_todo)

        fill_stats = fill_target_column_workbook(
            TAG_SOURCE,
            filled,
            source_col="source",
            target_col="target",
            template_workbook=standalone_todo,
            min_group_size=2,
        )
        self.assertEqual(fill_stats["autofilled_count"], 4)

        filled_book = load_workbook(filled, data_only=True)
        rows = list(filled_book.active.iter_rows(values_only=True))
        self.assertEqual(rows[1][1], '<a href="shop">Pack VIP30 FR</a>')
        self.assertEqual(rows[2][1], '<a href="shop">Pack VIP40 FR</a>')
        self.assertEqual(rows[3][1], '<img src="coin.png"/>')
        self.assertEqual(rows[4][1], "Brand new line FR")

    def test_fill_writes_target_when_tag_mismatch_warning_exists(self):
        from phraseloom.workflow import generate_workbook

        pack = TESTFILES / "tag_source_l10n" / "tag_source_tm_prefill_pack.xlsx"
        report = TESTFILES / "tag_source_l10n" / "tag_source_report.xlsx"

        generate_workbook(
            TAG_SOURCE,
            pack,
            source_col="source",
            target_col="target",
            min_group_size=2,
            use_existing_targets=False,
        )

        book = load_workbook(pack)
        ws = book["translation_units"]
        headers = [cell.value for cell in ws[1]]
        source_idx = headers.index("source_unit") + 1
        target_idx = headers.index("target_unit") + 1
        for row in ws.iter_rows(min_row=2):
            if row[source_idx - 1].value == "{t1_op}VIP{num1} Pack{t1_cl}":
                row[target_idx - 1].value = "Pack VIP{num1} FR"
            elif row[source_idx - 1].value == "Brand new line":
                row[target_idx - 1].value = "Brand new line FR"
        book.save(pack)

        generate_workbook(
            TAG_SOURCE,
            report,
            source_col="source",
            target_col="target",
            template_workbook=pack,
            min_group_size=2,
            use_existing_targets=False,
        )

        report_book = load_workbook(report, data_only=True)
        source_map = report_book["source_map"]
        headers = [cell.value for cell in source_map[1]]
        auto_idx = headers.index("auto_target")
        warning_idx = headers.index("warning")
        rows = list(source_map.iter_rows(min_row=2, values_only=True))

        self.assertEqual(rows[0][auto_idx], "Pack VIP30 FR")
        self.assertEqual(rows[1][auto_idx], "Pack VIP40 FR")
        self.assertTrue(any("tag_mismatch" in str(row[warning_idx]) for row in rows))


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run the testfiles workflow tests**

Run:

```powershell
py -3 -m unittest tests.test_tag_workflow_testfiles -v
```

Expected: all `tests.test_tag_workflow_testfiles` tests pass and generate only ignored files under `testfiles/`.

- [ ] **Step 3: Run all tests**

Run:

```powershell
py -3 -m unittest discover -v
```

Expected: all discovered tests pass.

- [ ] **Step 4: Commit testfiles-based workflow coverage**

Run:

```powershell
git add tests/test_tag_workflow_testfiles.py
git commit -m "Add testfiles tag workflow coverage"
```

---

### Task 7: Final Verification With Real Local Testfiles

**Files:**
- Inspect: `testfiles/TM.xlsx`
- Inspect: `testfiles/for_test.xlsx`
- No source changes expected.

- [ ] **Step 1: Run unit and workflow regression suite**

Run:

```powershell
py -3 -m unittest discover -v
```

Expected: all tests pass.

- [ ] **Step 2: Run TM extraction against the local `testfiles/TM.xlsx` sample**

Run:

```powershell
py -3 -m phraseloom.cli tm-extract testfiles\TM.xlsx --source-col source --target-col target
```

Expected: command exits 0 and writes:

```text
testfiles\TM_l10n\TM_reusable_units.xlsx
```

- [ ] **Step 3: Run extraction against the local `testfiles/for_test.xlsx` sample**

Run:

```powershell
py -3 -m phraseloom.cli extract testfiles\for_test.xlsx --source-col source --target-col target --tm testfiles\TM_l10n\TM_reusable_units.xlsx --no-existing-targets
```

Expected: command exits 0 and writes:

```text
testfiles\for_test_l10n\for_test_tm_prefill_pack.xlsx
testfiles\for_test_l10n\for_test_translator_todo.xlsx
```

- [ ] **Step 4: Run fill against the local `testfiles/for_test.xlsx` sample**

Run:

```powershell
py -3 -m phraseloom.cli fill testfiles\for_test.xlsx --templates testfiles\for_test_l10n\for_test_translator_todo.xlsx --source-col source --target-col target --mode target-column
```

Expected: command exits 0 and writes:

```text
testfiles\for_test_l10n\for_test_filled_result.xlsx
```

- [ ] **Step 5: Check git status**

Run:

```powershell
git -c core.excludesfile= status --short
```

Expected: no tracked source changes. Ignored Excel outputs under `testfiles/` do not appear.

- [ ] **Step 6: Commit final adjustments if Task 7 revealed tracked fixes**

If Task 7 required source fixes, run:

```powershell
git add phraseloom tests
git commit -m "Stabilize tag extractor workflow"
```

If Task 7 required no source fixes, do not create an empty commit.

---

## Self-Review Checklist

- Spec coverage: Tasks 1-5 cover placeholder contract, extraction, template isolation, tag-only auto-fill, validation, restoration, and warnings. Task 6 covers `testfiles/` workflow fixtures. Task 7 covers the existing local `testfiles/TM.xlsx` and `testfiles/for_test.xlsx` samples.
- Placeholder scan: This plan contains concrete file paths, code snippets, commands, and expected results. It avoids open-ended implementation placeholders.
- Type consistency: `TagToken`, `TagExtraction`, `TagValidation`, `RowItem`, `RowFillResult`, `restore_tags()`, `validate_tag_placeholders()`, and `is_tag_only_segment()` are introduced before tasks use them.
- UX constraint: The plan does not add a command, required CLI flag, or translator-managed sheet.
