# Configurable Tag Rules Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a TOML-backed tag allowlist so PhraseLoom protects real structural tags, leaves translatable angle-bracket labels such as `<Activate>` raw, and keeps TM source/target serialization aligned for safe prefill and fill.

**Architecture:** Add a focused `phraseloom.tag_rules` module that loads and hashes tag rules once per workflow. Pass the loaded `TagRules` into row reading and tag extraction; keep target serialization source-metadata-driven. Write tag-rule metadata into generated workbooks and reject explicit config mismatches.

**Tech Stack:** Python 3.11, `tomllib`, `importlib.resources`, `hashlib`, `json`, `openpyxl`, `unittest`, existing PhraseLoom CLI/workflow modules.

---

## File Structure

- Create `phraseloom/tag_rules.py`: dataclass, default packaged config loading, custom config loading, normalized hashing, membership helpers.
- Create `phraseloom/tag_rules.toml`: default allowlist config used by CLI and workflow.
- Modify `pyproject.toml`: include `*.toml` package data.
- Modify `phraseloom/tag_engine.py`: accept `TagRules`, skip disallowed angle/BBCode spans, keep raw `{...}` behavior.
- Modify `phraseloom/excel_io.py`: pass tag rules into `_read_source_rows`, write/read tag-rule metadata, validate config mismatch while loading TM/todo workbooks.
- Modify `phraseloom/workflow.py`: load rules once, pass them through TM extract/extract/fill/examples, merge row tag warnings into unit warnings.
- Modify `phraseloom/cli.py`: add `--tag-config` to `tm-extract`, `extract`, `fill`, and legacy command paths.
- Modify `phraseloom/workbook_schema.py`: add metadata keys for tag-rule tracking.
- Modify tests in `tests/test_tag_engine.py` and `tests/test_template_workflow.py`.
- Add `tests/test_tag_rules.py`.

## Task 1: Add Tag Rules Loader And Default Config

**Files:**
- Create: `phraseloom/tag_rules.py`
- Create: `phraseloom/tag_rules.toml`
- Modify: `pyproject.toml`
- Test: `tests/test_tag_rules.py`

- [ ] **Step 1: Write failing tag-rules tests**

Create `tests/test_tag_rules.py` with:

```python
import tempfile
import unittest
from pathlib import Path


class TagRulesTests(unittest.TestCase):
    def test_default_rules_allow_known_formatting_tags(self):
        from phraseloom.tag_rules import default_tag_rules

        rules = default_tag_rules()

        self.assertEqual(rules.version, 1)
        self.assertTrue(rules.allows_angle("color"))
        self.assertTrue(rules.allows_angle("COLOR"))
        self.assertTrue(rules.allows_angle("img"))
        self.assertTrue(rules.allows_angle("c"))
        self.assertFalse(rules.allows_angle("activate"))
        self.assertTrue(rules.allows_bbcode("color"))
        self.assertTrue(rules.protect_raw_braces)

    def test_custom_rules_load_from_toml(self):
        from phraseloom.tag_rules import load_tag_rules

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "tag_rules.toml"
            path.write_text(
                "\n".join(
                    [
                        "version = 1",
                        "",
                        "[angle_tags]",
                        'mode = "allowlist"',
                        'allowed = ["foo"]',
                        "",
                        "[bbcode_tags]",
                        'mode = "allowlist"',
                        'allowed = ["bar"]',
                        "",
                        "[raw_braces]",
                        "protect_all = false",
                    ]
                ),
                encoding="utf-8",
            )

            rules = load_tag_rules(path)

        self.assertTrue(rules.allows_angle("foo"))
        self.assertFalse(rules.allows_angle("color"))
        self.assertTrue(rules.allows_bbcode("bar"))
        self.assertFalse(rules.protect_raw_braces)

    def test_normalized_hash_ignores_order_and_case(self):
        from phraseloom.tag_rules import TagRules, normalized_tag_rules_hash

        left = TagRules(
            version=1,
            angle_allowed=frozenset({"color", "img"}),
            bbcode_allowed=frozenset({"b", "color"}),
            protect_raw_braces=True,
            source="left",
        )
        right = TagRules(
            version=1,
            angle_allowed=frozenset({"IMG", "COLOR"}),
            bbcode_allowed=frozenset({"COLOR", "B"}),
            protect_raw_braces=True,
            source="right",
        )

        self.assertEqual(
            normalized_tag_rules_hash(left),
            normalized_tag_rules_hash(right),
        )

    def test_invalid_mode_reports_config_error(self):
        from phraseloom.errors import ConfigError
        from phraseloom.tag_rules import load_tag_rules

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "bad.toml"
            path.write_text(
                "\n".join(
                    [
                        "version = 1",
                        "",
                        "[angle_tags]",
                        'mode = "denylist"',
                        'allowed = ["color"]',
                        "",
                        "[bbcode_tags]",
                        'mode = "allowlist"',
                        'allowed = ["color"]',
                        "",
                        "[raw_braces]",
                        "protect_all = true",
                    ]
                ),
                encoding="utf-8",
            )

            with self.assertRaises(ConfigError) as raised:
                load_tag_rules(path)

        self.assertIn("angle_tags.mode must be 'allowlist'", str(raised.exception))


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run the new tests and verify they fail**

Run:

```powershell
py -3 -m unittest tests.test_tag_rules -v
```

Expected: FAIL because `phraseloom.tag_rules` does not exist.

- [ ] **Step 3: Add the default TOML config**

Create `phraseloom/tag_rules.toml`:

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

- [ ] **Step 4: Include TOML package data**

Add this section to `pyproject.toml`:

```toml
[tool.setuptools.package-data]
phraseloom = ["*.toml"]
```

- [ ] **Step 5: Implement `phraseloom/tag_rules.py`**

Create `phraseloom/tag_rules.py`:

```python
from __future__ import annotations

import hashlib
import json
import tomllib
from dataclasses import dataclass
from functools import lru_cache
from importlib import resources
from pathlib import Path
from typing import Any

from .errors import ConfigError


@dataclass(frozen=True)
class TagRules:
    version: int
    angle_allowed: frozenset[str]
    bbcode_allowed: frozenset[str]
    protect_raw_braces: bool
    source: str = "default"

    def allows_angle(self, name: str | None) -> bool:
        return bool(name) and str(name).lower() in self.angle_allowed

    def allows_bbcode(self, name: str | None) -> bool:
        return bool(name) and str(name).lower() in self.bbcode_allowed


def default_tag_rules() -> TagRules:
    return _default_tag_rules()


@lru_cache(maxsize=1)
def _default_tag_rules() -> TagRules:
    text = resources.files("phraseloom").joinpath("tag_rules.toml").read_text(
        encoding="utf-8"
    )
    return _parse_tag_rules(tomllib.loads(text), source="default")


def load_tag_rules(path: str | Path | None = None) -> TagRules:
    if path is None:
        return default_tag_rules()
    config_path = Path(path)
    try:
        data = tomllib.loads(config_path.read_text(encoding="utf-8"))
    except OSError as error:
        raise ConfigError(f"Cannot read tag config {config_path}: {error}") from error
    except tomllib.TOMLDecodeError as error:
        raise ConfigError(f"Invalid tag config {config_path}: {error}") from error
    return _parse_tag_rules(data, source=str(config_path))


def normalized_tag_rules_hash(rules: TagRules) -> str:
    payload = {
        "version": rules.version,
        "angle_tags": sorted(name.lower() for name in rules.angle_allowed),
        "bbcode_tags": sorted(name.lower() for name in rules.bbcode_allowed),
        "raw_braces": {"protect_all": bool(rules.protect_raw_braces)},
    }
    encoded = json.dumps(payload, sort_keys=True, separators=(",", ":")).encode(
        "utf-8"
    )
    return hashlib.sha256(encoded).hexdigest()


def _parse_tag_rules(data: dict[str, Any], *, source: str) -> TagRules:
    version = data.get("version")
    if version != 1:
        raise ConfigError(f"tag config version must be 1, got {version!r}")

    angle = _read_allowlist_section(data, "angle_tags")
    bbcode = _read_allowlist_section(data, "bbcode_tags")
    raw_braces = data.get("raw_braces")
    if not isinstance(raw_braces, dict):
        raise ConfigError("tag config needs [raw_braces]")
    protect_raw_braces = raw_braces.get("protect_all")
    if not isinstance(protect_raw_braces, bool):
        raise ConfigError("raw_braces.protect_all must be true or false")

    return TagRules(
        version=version,
        angle_allowed=frozenset(angle),
        bbcode_allowed=frozenset(bbcode),
        protect_raw_braces=protect_raw_braces,
        source=source,
    )


def _read_allowlist_section(data: dict[str, Any], section_name: str) -> set[str]:
    section = data.get(section_name)
    if not isinstance(section, dict):
        raise ConfigError(f"tag config needs [{section_name}]")
    mode = section.get("mode")
    if mode != "allowlist":
        raise ConfigError(f"{section_name}.mode must be 'allowlist'")
    allowed = section.get("allowed")
    if not isinstance(allowed, list) or not all(
        isinstance(item, str) and item.strip() for item in allowed
    ):
        raise ConfigError(f"{section_name}.allowed must be a list of tag names")
    return {item.strip().lower() for item in allowed}


__all__ = [
    "TagRules",
    "default_tag_rules",
    "load_tag_rules",
    "normalized_tag_rules_hash",
]
```

- [ ] **Step 6: Run tag-rules tests and commit**

Run:

```powershell
py -3 -m unittest tests.test_tag_rules -v
```

Expected: all 4 tests pass.

Commit:

```powershell
git add pyproject.toml phraseloom/tag_rules.py phraseloom/tag_rules.toml tests/test_tag_rules.py
git commit -m "Add configurable tag rules"
```

## Task 2: Make Tag Extraction Use The Allowlist

**Files:**
- Modify: `phraseloom/tag_engine.py`
- Modify: `tests/test_tag_engine.py`

- [ ] **Step 1: Add failing extraction tests**

Append these tests to `TagEngineTests` in `tests/test_tag_engine.py`:

```python
    def test_default_rules_keep_unlisted_angle_label_raw(self):
        from phraseloom.tag_engine import RAW_PLACEHOLDER, extract_tags

        result = extract_tags("<Activate> HP increased by {a}%")

        self.assertEqual(result.text, "<Activate> HP increased by {1}%")
        self.assertEqual(
            tuple((tag.kind, tag.placeholder, tag.raw) for tag in result.tags),
            ((RAW_PLACEHOLDER, "{1}", "{a}"),),
        )
        self.assertEqual(result.warnings, ())

    def test_default_rules_keep_unlisted_angle_label_with_spaces_raw(self):
        from phraseloom.tag_engine import RAW_PLACEHOLDER, extract_tags

        result = extract_tags("<Weather Change> before turn {a}")

        self.assertEqual(result.text, "<Weather Change> before turn {1}")
        self.assertEqual(
            tuple((tag.kind, tag.placeholder, tag.raw) for tag in result.tags),
            ((RAW_PLACEHOLDER, "{1}", "{a}"),),
        )
        self.assertEqual(result.warnings, ())

    def test_default_rules_still_extract_allowed_color_pair(self):
        from phraseloom.tag_engine import RAW_PLACEHOLDER, TAG_CLOSE, TAG_OPEN, extract_tags

        result = extract_tags("<color=#123>HP {a}</>")

        self.assertEqual(result.text, "{1>HP {2}<3}")
        self.assertEqual(
            tuple((tag.kind, tag.placeholder, tag.raw, tag.partner_index) for tag in result.tags),
            (
                (TAG_OPEN, "{1>", "<color=#123>", None),
                (RAW_PLACEHOLDER, "{2}", "{a}", None),
                (TAG_CLOSE, "<3}", "</>", 1),
            ),
        )
        self.assertEqual(result.warnings, ())

    def test_default_rules_leave_unlisted_named_close_raw(self):
        from phraseloom.tag_engine import extract_tags

        result = extract_tags("<foo>x</foo>")

        self.assertEqual(result.text, "<foo>x</foo>")
        self.assertEqual(result.tags, ())
        self.assertEqual(result.warnings, ())
```

- [ ] **Step 2: Run the focused tests and verify they fail**

Run:

```powershell
py -3 -m unittest tests.test_tag_engine.TagEngineTests.test_default_rules_keep_unlisted_angle_label_raw tests.test_tag_engine.TagEngineTests.test_default_rules_still_extract_allowed_color_pair tests.test_tag_engine.TagEngineTests.test_default_rules_leave_unlisted_named_close_raw -v
```

Expected: FAIL because all angle-like tags are still extracted.

- [ ] **Step 3: Update imports and `extract_tags` signature**

Modify the imports near the top of `phraseloom/tag_engine.py`:

```python
from .tag_rules import TagRules, default_tag_rules
```

Change the function signature and initialize rules:

```python
def extract_tags(text: str, rules: TagRules | None = None) -> TagExtraction:
    source = "" if text is None else str(text)
    active_rules = default_tag_rules() if rules is None else rules
```

- [ ] **Step 4: Gate raw-brace extraction on config**

Replace the raw-brace branch inside `extract_tags` with:

```python
        if _RAW_BRACE_RE.fullmatch(raw):
            if not active_rules.protect_raw_braces:
                chunks.append(raw)
                pos = found.end()
                continue
            index = next_index
            next_index += 1
            placeholder = make_protected_token(index, RAW_PLACEHOLDER)
            tags.append(TagToken(index, RAW_PLACEHOLDER, placeholder, raw))
            chunks.append(placeholder)
            pos = found.end()
            continue
```

- [ ] **Step 5: Add tag allowlist helpers**

Add these helpers below `_classify_raw_tag`:

```python
def _is_raw_tag_allowed(raw: str, kind: str, name: str | None, rules: TagRules) -> bool:
    if raw.startswith("<"):
        if raw == "</>":
            return True
        return rules.allows_angle(name)
    if raw == "[/]":
        return True
    return rules.allows_bbcode(name)
```

- [ ] **Step 6: Skip disallowed tag-like spans before stack mutation**

After:

```python
        kind, name = _classify_raw_tag(raw)
```

insert:

```python
        if not _is_raw_tag_allowed(raw, kind, name, active_rules):
            chunks.append(raw)
            pos = found.end()
            continue
```

Keep shorthand close handling in the close branch. If `</>` or `[/]` appears
with no protected open on the stack, the existing unpaired close warning remains
valid because there is no allowed structural opener to close.

- [ ] **Step 7: Make plain BBCode matching rules-aware**

Change:

```python
matched_plain_bbcode_starts = _matched_plain_bbcode_open_starts(spans)
```

to:

```python
matched_plain_bbcode_starts = _matched_plain_bbcode_open_starts(spans, active_rules)
```

Change `_matched_plain_bbcode_open_starts` to accept rules:

```python
def _matched_plain_bbcode_open_starts(
    spans: list[re.Match[str]], rules: TagRules
) -> set[int]:
```

Inside its loop, change the close branch to:

```python
        if _is_named_bbcode_close(raw):
            name = raw[2:-1].strip().lower()
            if rules.allows_bbcode(name):
                later_closes_by_name[name] = later_closes_by_name.get(name, 0) + 1
```

Change the open branch to:

```python
        elif _is_plain_bbcode_open(raw):
            name = _bbcode_tag_name(raw)
            if rules.allows_bbcode(name) and later_closes_by_name.get(name, 0) > 0:
                matched_starts.add(found.start())
                later_closes_by_name[name] -= 1
```

- [ ] **Step 8: Update old tests that used `<a>` as an always-allowed tag**

Because default rules no longer allow `a`, replace tag-like test examples that
need extraction with allowed tags:

- `<a href="shop">VIP10</a>` -> `<color=#fff>VIP10</color>`
- `<a>x</a>` -> `<color=#fff>x</color>`
- `<a>Text</>` -> `<color>Text</>`

Keep one explicit test for `<a>` as disallowed if useful:

```python
    def test_default_rules_leave_anchor_like_text_raw(self):
        from phraseloom.tag_engine import extract_tags

        result = extract_tags('<a href="shop">VIP10</a>')

        self.assertEqual(result.text, '<a href="shop">VIP10</a>')
        self.assertEqual(result.tags, ())
        self.assertEqual(result.warnings, ())
```

- [ ] **Step 9: Run focused tag-engine tests and commit**

Run:

```powershell
py -3 -m unittest tests.test_tag_engine -v
```

Expected: all tag-engine tests pass.

Commit:

```powershell
git add phraseloom/tag_engine.py tests/test_tag_engine.py
git commit -m "Respect tag allowlist during extraction"
```

## Task 3: Propagate Tag Rules Through Workflow And CLI

**Files:**
- Modify: `phraseloom/excel_io.py`
- Modify: `phraseloom/workflow.py`
- Modify: `phraseloom/cli.py`
- Modify: `tests/test_template_workflow.py`

- [ ] **Step 1: Add failing CLI and workflow tests**

Add this test to `TemplateDemoTests` in `tests/test_template_workflow.py`:

```python
    def test_custom_tag_config_controls_workflow_extraction(self):
        from phraseloom.workflow import generate_workbook

        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            input_path = tmp_path / "source.xlsx"
            output_path = tmp_path / "pack.xlsx"
            config_path = tmp_path / "tag_rules.toml"

            wb = Workbook()
            ws = wb.active
            ws.append(["source", "target"])
            ws.append(["<foo>Power {a}</foo>", ""])
            ws.append(["<foo>Power {b}</foo>", ""])
            wb.save(input_path)

            config_path.write_text(
                "\n".join(
                    [
                        "version = 1",
                        "",
                        "[angle_tags]",
                        'mode = "allowlist"',
                        'allowed = ["foo"]',
                        "",
                        "[bbcode_tags]",
                        'mode = "allowlist"',
                        'allowed = ["color"]',
                        "",
                        "[raw_braces]",
                        "protect_all = true",
                    ]
                ),
                encoding="utf-8",
            )

            generate_workbook(
                input_path,
                output_path,
                source_col="source",
                target_col="target",
                min_group_size=2,
                use_existing_targets=False,
                tag_config=config_path,
            )

            out = load_workbook(output_path, data_only=True)
            units = out["translation_units"]
            headers = [cell.value for cell in units[1]]
            source_idx = headers.index("source_unit")
            source_units = [
                row[source_idx]
                for row in units.iter_rows(min_row=2, values_only=True)
            ]

        self.assertIn("{1>Power {2}<3}", source_units)
```

Add this test:

```python
    def test_cli_accepts_tag_config_option(self):
        from phraseloom.cli import main

        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            input_path = tmp_path / "source.xlsx"
            output_path = tmp_path / "pack.xlsx"
            config_path = tmp_path / "tag_rules.toml"

            wb = Workbook()
            ws = wb.active
            ws.append(["source", "target"])
            ws.append(["<foo>Power {a}</foo>", ""])
            ws.append(["<foo>Power {b}</foo>", ""])
            wb.save(input_path)

            config_path.write_text(
                "\n".join(
                    [
                        "version = 1",
                        "",
                        "[angle_tags]",
                        'mode = "allowlist"',
                        'allowed = ["foo"]',
                        "",
                        "[bbcode_tags]",
                        'mode = "allowlist"',
                        'allowed = ["color"]',
                        "",
                        "[raw_braces]",
                        "protect_all = true",
                    ]
                ),
                encoding="utf-8",
            )

            exit_code = main(
                [
                    "extract",
                    str(input_path),
                    "-o",
                    str(output_path),
                    "--source-col",
                    "source",
                    "--target-col",
                    "target",
                    "--tag-config",
                    str(config_path),
                    "--no-existing-targets",
                ]
            )

        self.assertEqual(exit_code, 0)
        self.assertTrue(output_path.exists())
```

- [ ] **Step 2: Run tests and verify they fail**

Run:

```powershell
py -3 -m unittest tests.test_template_workflow.TemplateDemoTests.test_custom_tag_config_controls_workflow_extraction tests.test_template_workflow.TemplateDemoTests.test_cli_accepts_tag_config_option -v
```

Expected: FAIL because workflow and CLI do not accept `tag_config`.

- [ ] **Step 3: Update `_read_source_rows` to accept rules**

In `phraseloom/excel_io.py`, import:

```python
from .tag_rules import TagRules, default_tag_rules
```

Change `_read_source_rows` signature:

```python
def _read_source_rows(
    input_path: Path,
    source_col: str | int,
    target_col: str | int | None,
    *,
    tag_rules: TagRules | None = None,
) -> list[RowItem]:
```

At the top of the function body:

```python
    active_tag_rules = default_tag_rules() if tag_rules is None else tag_rules
```

Change source extraction:

```python
            source_extraction = extract_tags(raw_source, rules=active_tag_rules)
```

- [ ] **Step 4: Update workflow public APIs**

In `phraseloom/workflow.py`, import:

```python
from .tag_rules import TagRules, load_tag_rules
```

Add this keyword-only argument to `generate_workbook`, after
`use_existing_targets`:

```python
    tag_config: str | Path | None = None,
```

Add this keyword-only argument to `fill_target_column_workbook`, after
`min_group_size`:

```python
    tag_config: str | Path | None = None,
```

Add this keyword-only argument to `generate_tm_pairs`, after `min_group_size`:

```python
    tag_config: str | Path | None = None,
```

In each public function, load once:

```python
    tag_rules = load_tag_rules(tag_config)
```

Pass `tag_rules=tag_rules` into `_build_fill_context`, `generate_tm_pairs`, and
`_read_source_rows`.

Change `_build_fill_context` signature:

```python
    tag_rules: TagRules,
```

Change its row read:

```python
    rows = _read_source_rows(input_path, source_col, target_col, tag_rules=tag_rules)
```

Change `_build_provided_units` signature:

```python
    tag_rules: TagRules,
```

Change example extraction:

```python
        source_extraction = extract_tags(source, rules=tag_rules)
```

- [ ] **Step 5: Pass tag rules into provided-unit loading**

Update `_build_fill_context` to call:

```python
    provided_units, provided_sources = _build_provided_units(
        examples,
        template_workbook,
        tm_workbook,
        tag_rules=tag_rules,
    )
```

Keep `_load_translated_units` unchanged in this task. Metadata validation is
added in Task 5.

- [ ] **Step 6: Add CLI `--tag-config` arguments**

In each CLI parser (`_main_tm_extract`, `_main_extract`, `_main_fill`,
`_main_legacy`), add:

```python
    parser.add_argument(
        "--tag-config",
        type=Path,
        help="TOML file defining which tag-like spans are protected",
    )
```

Pass `tag_config=args.tag_config` into workflow calls.

- [ ] **Step 7: Run focused tests and commit**

Run:

```powershell
py -3 -m unittest tests.test_template_workflow.TemplateDemoTests.test_custom_tag_config_controls_workflow_extraction tests.test_template_workflow.TemplateDemoTests.test_cli_accepts_tag_config_option -v
```

Expected: both tests pass.

Run:

```powershell
py -3 -m unittest tests.test_template_workflow.TemplateDemoTests.test_top_level_help_promotes_package_cli -v
```

Expected: PASS.

Commit:

```powershell
git add phraseloom/excel_io.py phraseloom/workflow.py phraseloom/cli.py tests/test_template_workflow.py
git commit -m "Thread tag rules through workflow"
```

## Task 4: Prove TM Target Units Are Source-Metadata Aligned

**Files:**
- Modify: `tests/test_template_workflow.py`
- Modify: `phraseloom/workflow.py`

- [ ] **Step 1: Add failing TM normalization tests**

Add this test to `TemplateDemoTests`:

```python
    def test_tm_prefill_restores_current_row_color_attributes(self):
        from phraseloom.workflow import (
            fill_target_column_workbook,
            generate_tm_pairs,
            generate_workbook,
        )

        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            tm_input = tmp_path / "tm.xlsx"
            tm_pairs = tmp_path / "tm_pairs.xlsx"
            source_input = tmp_path / "source.xlsx"
            pack_path = tmp_path / "pack.xlsx"
            filled_path = tmp_path / "filled.xlsx"

            wb = Workbook()
            ws = wb.active
            ws.append(["source", "target"])
            ws.append(["<color=#123>Source</>", "<color=#123>Target</>"])
            wb.save(tm_input)

            generate_tm_pairs(
                tm_input,
                tm_pairs,
                source_col="source",
                target_col="target",
            )

            wb = Workbook()
            ws = wb.active
            ws.append(["source", "target"])
            ws.append(["<color=#3333>Source</>", ""])
            wb.save(source_input)

            generate_workbook(
                source_input,
                pack_path,
                source_col="source",
                target_col="target",
                tm_workbook=tm_pairs,
                use_existing_targets=False,
            )
            fill_target_column_workbook(
                source_input,
                filled_path,
                source_col="source",
                target_col="target",
                template_workbook=pack_path,
            )

            filled = load_workbook(filled_path, data_only=True)
            rows = list(filled.active.iter_rows(values_only=True))

        self.assertEqual(rows[1][1], "<color=#3333>Target</>")
```

Add this test:

```python
    def test_tm_prefill_keeps_unlisted_angle_labels_translatable(self):
        from phraseloom.workflow import generate_tm_pairs

        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            tm_input = tmp_path / "tm.xlsx"
            tm_pairs = tmp_path / "tm_pairs.xlsx"

            wb = Workbook()
            ws = wb.active
            ws.append(["source", "target"])
            ws.append(
                [
                    "<Activate> HP increased by {a}%",
                    "<Active> PV augmentes de {a}%",
                ]
            )
            wb.save(tm_input)

            generate_tm_pairs(
                tm_input,
                tm_pairs,
                source_col="source",
                target_col="target",
            )

            out = load_workbook(tm_pairs, data_only=True)
            ws = out["tm_pairs"]
            headers = [cell.value for cell in ws[1]]
            source_idx = headers.index("source_unit")
            target_idx = headers.index("target_unit")
            row = next(ws.iter_rows(min_row=2, values_only=True))

        self.assertEqual(row[source_idx], "<Activate> HP increased by {1}%")
        self.assertEqual(row[target_idx], "<Active> PV augmentes de {1}%")
```

- [ ] **Step 2: Run tests and verify current failure pattern**

Run:

```powershell
py -3 -m unittest tests.test_template_workflow.TemplateDemoTests.test_tm_prefill_restores_current_row_color_attributes tests.test_template_workflow.TemplateDemoTests.test_tm_prefill_keeps_unlisted_angle_labels_translatable -v
```

Expected: at least the `<Activate>` test fails before Task 2 is implemented. If
Task 2 is already complete, both tests should pass. Continue to Step 3 to
improve warning propagation either way.

- [ ] **Step 3: Surface row serialization warnings in unit warnings**

In `phraseloom/workflow.py`, add:

```python
def _row_serialization_warnings(items: Iterable[RowItem]) -> list[str]:
    warnings: list[str] = []
    for item in items:
        warnings.extend(item.tag_warnings)
        warnings.extend(item.target_tag_warnings)
    return list(dict.fromkeys(warnings))
```

At the start of `_unit_warning`, after `warnings: list[str] = []`, add:

```python
    warnings.extend(_row_serialization_warnings(items))
```

Because `items` is iterated later in `_unit_warning`, make the first line:

```python
    item_list = list(items)
```

Then replace later loops over `items` in `_unit_warning` with `item_list`.

- [ ] **Step 4: Add warning propagation test**

Add this test to `TemplateDemoTests`:

```python
    def test_tm_pair_warning_includes_target_serialization_warning(self):
        from phraseloom.workflow import generate_tm_pairs

        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            tm_input = tmp_path / "tm.xlsx"
            tm_pairs = tmp_path / "tm_pairs.xlsx"

            wb = Workbook()
            ws = wb.active
            ws.append(["source", "target"])
            ws.append(["<br/>", "no matching span"])
            wb.save(tm_input)

            generate_tm_pairs(
                tm_input,
                tm_pairs,
                source_col="source",
                target_col="target",
            )

            out = load_workbook(tm_pairs, data_only=True)
            ws = out["tm_pairs"]
            headers = [cell.value for cell in ws[1]]
            warning_idx = headers.index("warning")
            row = next(ws.iter_rows(min_row=2, values_only=True))

        self.assertIn("source_protected_span_not_found: <br/>", row[warning_idx])
```

- [ ] **Step 5: Run focused tests and commit**

Run:

```powershell
py -3 -m unittest tests.test_template_workflow.TemplateDemoTests.test_tm_prefill_restores_current_row_color_attributes tests.test_template_workflow.TemplateDemoTests.test_tm_prefill_keeps_unlisted_angle_labels_translatable tests.test_template_workflow.TemplateDemoTests.test_tm_pair_warning_includes_target_serialization_warning -v
```

Expected: all pass.

Commit:

```powershell
git add phraseloom/workflow.py tests/test_template_workflow.py
git commit -m "Verify TM protected target alignment"
```

## Task 5: Write And Validate Tag-Rule Workbook Metadata

**Files:**
- Modify: `phraseloom/workbook_schema.py`
- Modify: `phraseloom/excel_io.py`
- Modify: `phraseloom/workflow.py`
- Modify: `tests/test_template_workflow.py`

- [ ] **Step 1: Add failing metadata tests**

Add this test:

```python
    def test_generated_workbooks_record_tag_rules_metadata(self):
        from phraseloom.tag_rules import default_tag_rules, normalized_tag_rules_hash
        from phraseloom.workflow import generate_tm_pairs, generate_workbook

        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source_input = tmp_path / "source.xlsx"
            pack_output = tmp_path / "pack.xlsx"
            tm_input = tmp_path / "tm.xlsx"
            tm_output = tmp_path / "tm_pairs.xlsx"

            wb = Workbook()
            ws = wb.active
            ws.append(["source", "target"])
            ws.append(["<color=#123>Source</>", ""])
            wb.save(source_input)

            generate_workbook(
                source_input,
                pack_output,
                source_col="source",
                target_col="target",
                use_existing_targets=False,
            )

            wb = Workbook()
            ws = wb.active
            ws.append(["source", "target"])
            ws.append(["<color=#123>Source</>", "<color=#123>Target</>"])
            wb.save(tm_input)

            generate_tm_pairs(
                tm_input,
                tm_output,
                source_col="source",
                target_col="target",
            )

            expected_hash = normalized_tag_rules_hash(default_tag_rules())
            pack = load_workbook(pack_output, data_only=True)
            tm = load_workbook(tm_output, data_only=True)

            pack_metadata = {
                row[0]: row[1]
                for row in pack["_metadata"].iter_rows(min_row=2, values_only=True)
            }
            tm_metadata = {
                row[0]: row[1]
                for row in tm["_metadata"].iter_rows(min_row=2, values_only=True)
            }

        self.assertEqual(pack_metadata["tag_rules_version"], 1)
        self.assertEqual(pack_metadata["tag_rules_hash"], expected_hash)
        self.assertEqual(pack_metadata["tag_rules_source"], "default")
        self.assertEqual(tm_metadata["tag_rules_hash"], expected_hash)
```

Add this test:

```python
    def test_tag_config_mismatch_reports_user_facing_error(self):
        from phraseloom.errors import ConfigError
        from phraseloom.workflow import generate_tm_pairs, generate_workbook

        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            first_config = tmp_path / "first.toml"
            second_config = tmp_path / "second.toml"
            tm_input = tmp_path / "tm.xlsx"
            tm_pairs = tmp_path / "tm_pairs.xlsx"
            source_input = tmp_path / "source.xlsx"
            pack_output = tmp_path / "pack.xlsx"

            first_config.write_text(
                "\n".join(
                    [
                        "version = 1",
                        "",
                        "[angle_tags]",
                        'mode = "allowlist"',
                        'allowed = ["color"]',
                        "",
                        "[bbcode_tags]",
                        'mode = "allowlist"',
                        'allowed = ["color"]',
                        "",
                        "[raw_braces]",
                        "protect_all = true",
                    ]
                ),
                encoding="utf-8",
            )
            second_config.write_text(
                "\n".join(
                    [
                        "version = 1",
                        "",
                        "[angle_tags]",
                        'mode = "allowlist"',
                        'allowed = ["foo"]',
                        "",
                        "[bbcode_tags]",
                        'mode = "allowlist"',
                        'allowed = ["color"]',
                        "",
                        "[raw_braces]",
                        "protect_all = true",
                    ]
                ),
                encoding="utf-8",
            )

            wb = Workbook()
            ws = wb.active
            ws.append(["source", "target"])
            ws.append(["<color=#123>Source</>", "<color=#123>Target</>"])
            wb.save(tm_input)

            generate_tm_pairs(
                tm_input,
                tm_pairs,
                source_col="source",
                target_col="target",
                tag_config=first_config,
            )

            wb = Workbook()
            ws = wb.active
            ws.append(["source", "target"])
            ws.append(["<color=#333>Source</>", ""])
            wb.save(source_input)

            with self.assertRaises(ConfigError) as raised:
                generate_workbook(
                    source_input,
                    pack_output,
                    source_col="source",
                    target_col="target",
                    tm_workbook=tm_pairs,
                    use_existing_targets=False,
                    tag_config=second_config,
                )

        self.assertIn("tag config mismatch", str(raised.exception))
```

- [ ] **Step 2: Run tests and verify failure**

Run:

```powershell
py -3 -m unittest tests.test_template_workflow.TemplateDemoTests.test_generated_workbooks_record_tag_rules_metadata tests.test_template_workflow.TemplateDemoTests.test_tag_config_mismatch_reports_user_facing_error -v
```

Expected: FAIL because metadata keys are not written or validated.

- [ ] **Step 3: Add schema constants**

In `phraseloom/workbook_schema.py`, add:

```python
TAG_RULES_VERSION_KEY = "tag_rules_version"
TAG_RULES_HASH_KEY = "tag_rules_hash"
TAG_RULES_SOURCE_KEY = "tag_rules_source"
```

- [ ] **Step 4: Add metadata helpers in `excel_io.py`**

Import:

```python
from .errors import ColumnNotFoundError, ConfigError, TranslationUnitLoadError
from .tag_rules import TagRules, default_tag_rules, normalized_tag_rules_hash
```

Replace `_add_metadata_sheet` with:

```python
def _add_metadata_sheet(wb, tag_rules: TagRules | None = None) -> None:
    metadata = wb.create_sheet(schema.METADATA_SHEET)
    metadata.append(schema.METADATA_COLUMNS)
    _append_metadata_rows(metadata, tag_rules)
    metadata.sheet_state = "hidden"
```

Add:

```python
def _append_metadata_rows(ws, tag_rules: TagRules | None = None) -> None:
    ws.append([schema.SCHEMA_VERSION_KEY, schema.SCHEMA_VERSION])
    active_rules = default_tag_rules() if tag_rules is None else tag_rules
    ws.append([schema.TAG_RULES_VERSION_KEY, active_rules.version])
    ws.append([schema.TAG_RULES_HASH_KEY, normalized_tag_rules_hash(active_rules)])
    ws.append([schema.TAG_RULES_SOURCE_KEY, active_rules.source])


def _workbook_metadata(wb) -> dict[str, object]:
    if schema.METADATA_SHEET not in wb.sheetnames:
        return {}
    ws = wb[schema.METADATA_SHEET]
    return {
        row[0]: row[1]
        for row in ws.iter_rows(min_row=2, max_col=2, values_only=True)
        if row and row[0] is not None
    }


def _validate_tag_rules_metadata(wb, path: Path, tag_rules: TagRules | None) -> None:
    if tag_rules is None:
        return
    metadata = _workbook_metadata(wb)
    actual_hash = metadata.get(schema.TAG_RULES_HASH_KEY)
    if actual_hash is None:
        return
    expected_hash = normalized_tag_rules_hash(tag_rules)
    if str(actual_hash) != expected_hash:
        raise ConfigError(
            f"Workbook {path} tag config mismatch. "
            f"Expected {expected_hash}, found {actual_hash}."
        )
```

- [ ] **Step 5: Pass tag rules into workbook writers**

Update signatures:

Update these writer signatures:

```python
def _write_output_workbook(
    output_path: Path,
    input_path: Path,
    units: list[TranslationUnit],
    result_rows: list[RowFillResult],
    *,
    tag_rules: TagRules | None = None,
) -> None:
```

```python
def _write_to_translate_workbook(
    output_path: Path,
    input_path: Path,
    units: list[TranslationUnit],
    *,
    tag_rules: TagRules | None = None,
) -> None:
```

```python
def _write_tm_workbook(
    output_path: Path,
    input_path: Path,
    units: list[TranslationUnit],
    rows: list[RowItem],
    *,
    tag_rules: TagRules | None = None,
) -> None:
```

In `_write_output_workbook`, call `_add_metadata_sheet(wb, tag_rules)` before
styling sheets.

In `_write_to_translate_workbook`, change:

```python
    _add_metadata_sheet(wb)
```

to:

```python
    _add_metadata_sheet(wb, tag_rules)
```

In `_write_tm_workbook`, call `_add_metadata_sheet(wb, tag_rules)` before styling
sheets.

- [ ] **Step 6: Validate metadata when loading translated units**

Change `_load_translated_units` signature:

```python
def _load_translated_units(
    path: Path, *, tag_rules: TagRules | None = None
) -> dict[tuple[str, str], str]:
```

After loading the workbook, before selecting sheets:

```python
        _validate_tag_rules_metadata(wb, path, tag_rules)
```

Update callers in `workflow.py`:

```python
        for key, target_unit in _load_translated_units(
            Path(tm_workbook), tag_rules=tag_rules
        ).items():
```

and:

```python
        for key, target_unit in _load_translated_units(
            Path(template_workbook), tag_rules=tag_rules
        ).items():
```

- [ ] **Step 7: Pass tag rules into writers from workflow**

Update calls:

```python
    _write_output_workbook(output_path, input_path, units, result_rows, tag_rules=tag_rules)
    _write_to_translate_workbook(to_translate_path, input_path, units, tag_rules=tag_rules)
    _write_tm_workbook(output_path, input_path, units, rows, tag_rules=tag_rules)
```

If line length exceeds local style, wrap arguments over multiple lines.

- [ ] **Step 8: Run metadata tests and commit**

Run:

```powershell
py -3 -m unittest tests.test_template_workflow.TemplateDemoTests.test_generated_workbooks_record_tag_rules_metadata tests.test_template_workflow.TemplateDemoTests.test_tag_config_mismatch_reports_user_facing_error -v
```

Expected: both pass.

Commit:

```powershell
git add phraseloom/workbook_schema.py phraseloom/excel_io.py phraseloom/workflow.py tests/test_template_workflow.py
git commit -m "Record tag rule metadata in workbooks"
```

## Task 6: Update Documentation For Current Cases

**Files:**
- Modify: `docs/tag-template-engine-cases.md`
- Modify: `docs/superpowers/specs/2026-05-11-configurable-tag-rules-design.md`
- Modify: `agent.md`

- [ ] **Step 1: Update case documentation**

In `docs/tag-template-engine-cases.md`, update the responsibility split and
extraction table so it says:

```markdown
`tag_engine` protects only tag names allowed by the active tag rules. Unknown
angle-bracket labels remain normal text. Raw `{...}` placeholders remain
protected when `raw_braces.protect_all = true`.
```

Add examples:

```markdown
| Allowed angle tag | `<color=#123>HP {a}</>` | `{1>HP {2}<3}` | `color` is in the default allowlist. |
| Unknown angle label | `<Activate> HP {a}` | `<Activate> HP {1}` | `activate` is not in the default allowlist. |
```

- [ ] **Step 2: Update onboarding notes**

In `agent.md`, add a short note near the protected-token contract:

```markdown
Tag extraction is governed by `phraseloom/tag_rules.toml`. The default
allowlist protects formatting tags such as `color`, `size`, `img`, `br`, `i`,
`u`, `outline`, and `c`; unknown angle-bracket labels such as `<Activate>` stay
translatable text.
```

- [ ] **Step 3: Run docs-related smoke tests and commit**

Run:

```powershell
py -3 -m unittest tests.test_template_workflow.CompatibilityShimTests -v
```

Expected: PASS.

Commit:

```powershell
git add docs/tag-template-engine-cases.md docs/superpowers/specs/2026-05-11-configurable-tag-rules-design.md agent.md
git commit -m "Document configurable tag rule behavior"
```

## Task 7: Full Automated Verification

**Files:**
- No code changes expected.

- [ ] **Step 1: Run full unittest suite**

Run:

```powershell
py -3 -m unittest discover -v
```

Expected: all tests pass.

- [ ] **Step 2: Compile changed Python modules**

Run:

```powershell
py -3 -m py_compile phraseloom/tag_rules.py phraseloom/tag_engine.py phraseloom/excel_io.py phraseloom/workflow.py phraseloom/cli.py phraseloom/workbook_schema.py
```

Expected: exit code 0.

- [ ] **Step 3: Commit any test-only corrections**

If tests required a correction, commit it:

```powershell
git add phraseloom tests docs agent.md pyproject.toml
git commit -m "Stabilize configurable tag rule tests"
```

If there are no corrections, skip this commit.

## Task 8: Business Verification With `testfiles\TM.xlsx`

**Files:**
- Use input: `D:\cat\PhraseLoom\testfiles\TM.xlsx`
- Generate outputs under: `D:\cat\PhraseLoom\testfiles\configurable_tag_rules_verification\`

- [ ] **Step 1: Create verification copies**

Run this script from repo root:

```powershell
@'
from pathlib import Path
from shutil import copy2
from openpyxl import load_workbook

root = Path(r"D:\cat\PhraseLoom")
src = root / "testfiles" / "TM.xlsx"
out = root / "testfiles" / "configurable_tag_rules_verification"
out.mkdir(parents=True, exist_ok=True)

tm_copy = out / "TM.xlsx"
source_copy = out / "TM_as_source_without_target.xlsx"
copy2(src, tm_copy)
copy2(src, source_copy)

wb = load_workbook(source_copy)
ws = wb.worksheets[0]
headers = [cell.value for cell in ws[1]]
target_col = headers.index("target") + 1
for row in range(2, ws.max_row + 1):
    ws.cell(row=row, column=target_col).value = None
wb.save(source_copy)
wb.close()

print(tm_copy)
print(source_copy)
'@ | py -3 -X utf8 -
```

Expected: script prints both copied workbook paths.

- [ ] **Step 2: Extract TM pairs from copied TM**

Run:

```powershell
py -3 -m phraseloom.cli tm-extract testfiles\configurable_tag_rules_verification\TM.xlsx -o testfiles\configurable_tag_rules_verification\TM_reusable_units.xlsx --source-col source --target-col target
```

Expected output includes:

```text
Wrote: testfiles\configurable_tag_rules_verification\TM_reusable_units.xlsx
TM source segments: 20532
TM pairs:
```

The exact TM pair count may change from the previous all-tag extraction because
unknown angle labels stay translatable.

- [ ] **Step 3: Extract target file using the copied TM as TM prefill**

Run:

```powershell
py -3 -m phraseloom.cli extract testfiles\configurable_tag_rules_verification\TM_as_source_without_target.xlsx -o testfiles\configurable_tag_rules_verification\TM_as_source_l10n\TM_as_source_tm_prefill_pack.xlsx --source-col source --target-col - --tm testfiles\configurable_tag_rules_verification\TM_reusable_units.xlsx --no-existing-targets
```

Expected:

```text
Total source rows: 20532
Already filled source rows: 20532
Units to translate: 0
```

If `Units to translate` is nonzero, inspect the missed source units before
continuing.

- [ ] **Step 4: Fill target column**

Run:

```powershell
py -3 -m phraseloom.cli fill testfiles\configurable_tag_rules_verification\TM_as_source_without_target.xlsx --templates testfiles\configurable_tag_rules_verification\TM_as_source_l10n\TM_as_source_translator_todo.xlsx -o testfiles\configurable_tag_rules_verification\TM_as_source_l10n\TM_as_source_filled_result.xlsx --source-col source --target-col target --mode target-column
```

Expected:

```text
Already filled source rows: 20532
Units to translate: 0
```

- [ ] **Step 5: Compare filled target against original TM target**

Run:

```powershell
@'
from pathlib import Path
from openpyxl import load_workbook
import re

base = Path(r"D:\cat\PhraseLoom\testfiles\configurable_tag_rules_verification")
original_path = base / "TM.xlsx"
filled_path = base / "TM_as_source_l10n" / "TM_as_source_filled_result.xlsx"

def read_rows(path):
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.worksheets[0]
        headers = [cell.value for cell in ws[1]]
        source_idx = headers.index("source")
        target_idx = headers.index("target")
        rows = []
        for row_number, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            source = "" if row[source_idx] is None else str(row[source_idx]).strip()
            target = "" if row[target_idx] is None else str(row[target_idx]).strip()
            if source:
                rows.append((row_number, source, target))
        return rows
    finally:
        wb.close()

original = read_rows(original_path)
filled = read_rows(filled_path)
if len(original) != len(filled):
    raise SystemExit(f"row count mismatch: original={len(original)} filled={len(filled)}")

mismatches = []
for (orig_row, orig_source, orig_target), (filled_row, filled_source, filled_target) in zip(original, filled):
    if orig_source != filled_source or orig_target != filled_target:
        mismatches.append((orig_row, orig_source, orig_target, filled_target))

protected_open_close_re = re.compile(r"\{[1-9]\d*>|<[1-9]\d*\}")
residual = [
    (row_number, source, target)
    for row_number, source, target in filled
    if protected_open_close_re.search(target)
]

print(f"original_rows={len(original)}")
print(f"filled_rows={len(filled)}")
print(f"mismatches={len(mismatches)}")
print(f"residual_open_close_tokens={len(residual)}")
if mismatches:
    print("first_mismatches=")
    for item in mismatches[:20]:
        print(item)
if residual:
    print("first_residual=")
    for item in residual[:20]:
        print(item)

if mismatches or residual:
    raise SystemExit(1)
'@ | py -3 -X utf8 -
```

Expected:

```text
original_rows=20532
filled_rows=20532
mismatches=0
residual_open_close_tokens=0
```

- [ ] **Step 6: Inspect the formerly broken `<Activate>` rows**

Run:

```powershell
@'
from pathlib import Path
from openpyxl import load_workbook

path = Path(r"D:\cat\PhraseLoom\testfiles\configurable_tag_rules_verification\TM_reusable_units.xlsx")
wb = load_workbook(path, read_only=True, data_only=True)
try:
    ws = wb["tm_pairs"]
    headers = [cell.value for cell in ws[1]]
    source_idx = headers.index("source_unit")
    target_idx = headers.index("target_unit")
    hits = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        source = "" if row[source_idx] is None else str(row[source_idx])
        target = "" if row[target_idx] is None else str(row[target_idx])
        if "<Activate>" in source or "<Activ" in target:
            hits.append((source, target))
    print(f"activate_hits={len(hits)}")
    for source, target in hits[:10]:
        print(source)
        print(target)
        print("---")
finally:
    wb.close()
'@ | py -3 -X utf8 -
```

Expected: matching rows show `<Activate>` and the translated angle label as raw
text, with raw placeholders converted to protected `{N}` tokens. Source units
for these rows must not start with `{1>`.

- [ ] **Step 7: Commit verification notes if a doc is updated**

If verification results are added to docs or `agent.md`, commit them:

```powershell
git add docs agent.md
git commit -m "Record configurable tag rule verification"
```

If no docs are updated, do not commit generated Excel files.

## Task 9: Final Review

**Files:**
- No code changes expected.

- [ ] **Step 1: Check git status**

Run:

```powershell
git status --short --branch
```

Expected: no unexpected source changes. Generated verification workbooks under
`testfiles/configurable_tag_rules_verification/` should remain untracked unless
the repository intentionally tracks test artifacts.

- [ ] **Step 2: Summarize results**

Prepare a final implementation summary that includes:

```text
- Tag rules config path and default allowlist.
- `<Activate>` no longer extracted as a protected open token.
- `<color>` attribute normalization still works and restores current-row raw tags.
- TM.xlsx self-prefill verification counts.
- Full unittest result.
```
