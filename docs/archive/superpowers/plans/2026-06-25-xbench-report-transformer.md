# Xbench Report Transformer Implementation Plan (Archived)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a CLI-first Toolshub utility that converts Xbench QA reports into a five-column, key-grouped Excel workbook for editing and filtering.

**Architecture:** Add a new focused package under `tools/xbench_report_transformer`. Keep parsing, grouping, and workbook output in one small CLI module because the feature is narrow, while exposing pure functions for unit tests. The tool reads Xbench rows, formats QA issues, groups by metadata-derived identity, and writes a new workbook without modifying the source report.

**Tech Stack:** Python 3.11, `openpyxl`, `argparse`, `unittest`, existing `tools.excel_output.build_prefixed_output_path`.

---

## File Structure

- Create `tools/xbench_report_transformer/__init__.py`
  - Package marker for imports and future GUI/workflow integration.
- Create `tools/xbench_report_transformer/transform_xbench_report.py`
  - Owns CLI parsing, metadata parsing, QA title formatting, row extraction, grouping, output workbook writing, and `main()`.
- Create `tools/xbench_report_transformer/README.md`
  - Short user-facing guide matching the existing per-tool README style.
- Create `tests/test_xbench_report_transformer.py`
  - Unit and process tests for metadata fallback, issue formatting, grouping, header detection, output generation, and non-mutation of the input workbook.
- Modify `README.md`
  - Add the new tool to the current tools list and navigation.
- Modify `docs/cli-usage.md`
  - Add CLI usage for the new transformer and mention the default output naming pattern.

## Implementation Notes

Use these names consistently:

- `OUTPUT_HEADERS = ("文件名", "key", "source", "target", "QA问题")`
- `OUTPUT_SHEET_NAME = "Xbench QA整理"`
- `HEADER_SOURCE = "source"`
- `HEADER_TARGET = "target"`
- `HEADER_COMMENTS = "comments"`
- `HEADER_METADATA = "metadata"`
- `FILE_NAME_EXTENSIONS = (".xlsx", ".xls", ".xlsm", ".csv", ".txt")`
- `build_default_output_path(input_path)` returns `build_prefixed_output_path(input_path, "xbench_transform_")`

Use dataclasses:

```python
@dataclass(frozen=True)
class ParsedMetadata:
    key: str
    file_name: str


@dataclass(frozen=True)
class XbenchIssue:
    issue_type: str
    source_term: str
    target_term: str


@dataclass(frozen=True)
class XbenchDetailRow:
    file_name: str
    key: str
    source: str
    target: str
    qa_issue: str
    group_key: str


@dataclass(frozen=True)
class TransformSummary:
    worksheet_title: str
    output_path: Path
    detail_count: int
    grouped_count: int
```

---

### Task 1: Pure Parsing Helpers

**Files:**
- Create: `tools/xbench_report_transformer/__init__.py`
- Create: `tools/xbench_report_transformer/transform_xbench_report.py`
- Test: `tests/test_xbench_report_transformer.py`

- [ ] **Step 1: Write the failing tests**

Create `tests/test_xbench_report_transformer.py` with these imports and test cases:

```python
from __future__ import annotations

import unittest

from tools.xbench_report_transformer.transform_xbench_report import (
    format_issue_text,
    parse_metadata,
    parse_qa_title,
)


class MetadataParsingTests(unittest.TestCase):
    def test_parse_two_metadata_lines_as_key_and_file_name(self) -> None:
        metadata = parse_metadata("LDLG_Text_ZH_q203101_Line_3\n磐城【配音】.xlsx")

        self.assertEqual(metadata.key, "LDLG_Text_ZH_q203101_Line_3")
        self.assertEqual(metadata.file_name, "磐城【配音】.xlsx")

    def test_parse_single_file_name_line_without_key(self) -> None:
        metadata = parse_metadata("磐城【配音】.xlsx")

        self.assertEqual(metadata.key, "")
        self.assertEqual(metadata.file_name, "磐城【配音】.xlsx")

    def test_parse_single_non_file_line_as_key(self) -> None:
        metadata = parse_metadata("LDLG_Text_ZH_q203101_Line_3")

        self.assertEqual(metadata.key, "LDLG_Text_ZH_q203101_Line_3")
        self.assertEqual(metadata.file_name, "")

    def test_parse_empty_metadata(self) -> None:
        metadata = parse_metadata(None)

        self.assertEqual(metadata.key, "")
        self.assertEqual(metadata.file_name, "")

    def test_parse_metadata_ignores_extra_lines_for_now(self) -> None:
        metadata = parse_metadata("Key_1\nfile.xlsx\nextra")

        self.assertEqual(metadata.key, "Key_1")
        self.assertEqual(metadata.file_name, "file.xlsx")


class QaTitleParsingTests(unittest.TestCase):
    def test_format_key_term_mismatch_title_as_term_pair_issue(self) -> None:
        issue = parse_qa_title("Key Term Mismatch (提示 / Avis)")

        self.assertEqual(issue.issue_type, "Key Term Mismatch")
        self.assertEqual(issue.source_term, "提示")
        self.assertEqual(issue.target_term, "Avis")
        self.assertEqual(format_issue_text(issue), "提示 -> Avis：Key Term Mismatch")

    def test_format_issue_preserves_quoted_terms(self) -> None:
        issue = parse_qa_title('Key Term Mismatch (“斑鸠” / "Colombe")')

        self.assertEqual(format_issue_text(issue), '“斑鸠” -> "Colombe"：Key Term Mismatch')

    def test_unparseable_title_is_used_as_issue_type(self) -> None:
        issue = parse_qa_title("Target same as Source")

        self.assertEqual(issue.issue_type, "Target same as Source")
        self.assertEqual(issue.source_term, "")
        self.assertEqual(issue.target_term, "")
        self.assertEqual(format_issue_text(issue), "Target same as Source")


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run the tests to verify RED**

Run:

```bash
python -m unittest tests.test_xbench_report_transformer
```

Expected: FAIL with `ModuleNotFoundError: No module named 'tools.xbench_report_transformer'`.

- [ ] **Step 3: Add the minimal implementation**

Create `tools/xbench_report_transformer/__init__.py`:

```python
"""Xbench report transformer package."""
```

Create `tools/xbench_report_transformer/transform_xbench_report.py` with this initial content:

```python
#!/usr/bin/env python3
"""Transform Xbench QA reports into a flat Toolshub-style workbook."""

from __future__ import annotations

import re
from dataclasses import dataclass


FILE_NAME_EXTENSIONS = (".xlsx", ".xls", ".xlsm", ".csv", ".txt")
QA_TITLE_PATTERN = re.compile(r"^(?P<issue_type>.*?)\s*\((?P<terms>.*)\)\s*$")


@dataclass(frozen=True)
class ParsedMetadata:
    key: str
    file_name: str


@dataclass(frozen=True)
class XbenchIssue:
    issue_type: str
    source_term: str
    target_term: str


def value_to_text(value: object) -> str:
    return "" if value is None else str(value).strip()


def looks_like_file_name(value: str) -> bool:
    normalized = value.strip().casefold()
    return normalized.endswith(FILE_NAME_EXTENSIONS)


def parse_metadata(value: object) -> ParsedMetadata:
    lines = [line.strip() for line in value_to_text(value).splitlines() if line.strip()]
    if len(lines) >= 2:
        return ParsedMetadata(key=lines[0], file_name=lines[1])
    if len(lines) == 1:
        if looks_like_file_name(lines[0]):
            return ParsedMetadata(key="", file_name=lines[0])
        return ParsedMetadata(key=lines[0], file_name="")
    return ParsedMetadata(key="", file_name="")


def parse_qa_title(title: object) -> XbenchIssue:
    title_text = value_to_text(title)
    match = QA_TITLE_PATTERN.match(title_text)
    if not match:
        return XbenchIssue(issue_type=title_text, source_term="", target_term="")

    issue_type = match.group("issue_type").strip()
    terms = match.group("terms").strip()
    source_term, separator, target_term = terms.partition(" / ")
    if not separator:
        return XbenchIssue(issue_type=title_text, source_term="", target_term="")
    return XbenchIssue(
        issue_type=issue_type,
        source_term=source_term.strip(),
        target_term=target_term.strip(),
    )


def format_issue_text(issue: XbenchIssue) -> str:
    if issue.source_term and issue.target_term:
        return f"{issue.source_term} -> {issue.target_term}：{issue.issue_type}"
    if issue.source_term:
        return f"{issue.source_term}：{issue.issue_type}"
    if issue.target_term:
        return f"-> {issue.target_term}：{issue.issue_type}"
    return issue.issue_type
```

- [ ] **Step 4: Run the tests to verify GREEN**

Run:

```bash
python -m unittest tests.test_xbench_report_transformer
```

Expected: PASS with all metadata and QA title parsing tests passing.

- [ ] **Step 5: Commit Task 1**

Run:

```bash
git add tools/xbench_report_transformer/__init__.py tools/xbench_report_transformer/transform_xbench_report.py tests/test_xbench_report_transformer.py
git commit -m "feat: add xbench parsing helpers"
```

---

### Task 2: Detail Row Extraction and Grouping

**Files:**
- Modify: `tools/xbench_report_transformer/transform_xbench_report.py`
- Modify: `tests/test_xbench_report_transformer.py`

- [ ] **Step 1: Add failing tests for row extraction and grouping**

Append these imports:

```python
from openpyxl import Workbook

from tools.xbench_report_transformer.transform_xbench_report import (
    collect_detail_rows,
    find_header_columns,
    group_detail_rows,
)
```

Add this test class:

```python
class RowExtractionAndGroupingTests(unittest.TestCase):
    def build_xbench_workbook(self) -> Workbook:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Xbench QA"
        worksheet["A1"] = "Exported QA Report"
        worksheet["C4"] = "Source"
        worksheet["D4"] = "Target"
        worksheet["E4"] = "Comments"
        worksheet["F4"] = "Metadata"
        worksheet["A5"] = "Key Term Mismatch (提示 / Avis)"
        worksheet["C6"] = "好刻意的提示！"
        worksheet["D6"] = "Sans blague !"
        worksheet["F6"] = "Key_1\nUI弹窗文字.xlsx"
        worksheet["A7"] = 'Key Term Mismatch (“斑鸠” / "Colombe")'
        worksheet["C8"] = "“斑鸠”&“诗人”"
        worksheet["D8"] = '"Colombe" & "Poète"'
        worksheet["F8"] = "Key_2\n磐城【配音】.xlsx"
        worksheet["A9"] = 'Key Term Mismatch (“诗人” / "Poète")'
        worksheet["C10"] = "“斑鸠”&“诗人”"
        worksheet["D10"] = '"Colombe" & "Poète"'
        worksheet["F10"] = "Key_2\n磐城【配音】.xlsx"
        return workbook

    def test_find_header_columns_detects_table_header_after_intro_rows(self) -> None:
        workbook = self.build_xbench_workbook()
        worksheet = workbook["Xbench QA"]

        header_row, columns = find_header_columns(worksheet)

        self.assertEqual(header_row, 4)
        self.assertEqual(columns["source"], 3)
        self.assertEqual(columns["target"], 4)
        self.assertEqual(columns["comments"], 5)
        self.assertEqual(columns["metadata"], 6)

    def test_collect_detail_rows_attaches_current_qa_issue(self) -> None:
        workbook = self.build_xbench_workbook()
        worksheet = workbook["Xbench QA"]

        detail_rows = collect_detail_rows(worksheet)

        self.assertEqual(len(detail_rows), 3)
        self.assertEqual(detail_rows[0].file_name, "UI弹窗文字.xlsx")
        self.assertEqual(detail_rows[0].key, "Key_1")
        self.assertEqual(detail_rows[0].source, "好刻意的提示！")
        self.assertEqual(detail_rows[0].target, "Sans blague !")
        self.assertEqual(detail_rows[0].qa_issue, "提示 -> Avis：Key Term Mismatch")
        self.assertEqual(detail_rows[0].group_key, "key:Key_1")

    def test_group_detail_rows_merges_duplicate_key_issues_with_chinese_semicolon(self) -> None:
        workbook = self.build_xbench_workbook()
        worksheet = workbook["Xbench QA"]

        grouped_rows = group_detail_rows(collect_detail_rows(worksheet))

        self.assertEqual(len(grouped_rows), 2)
        self.assertEqual(grouped_rows[1]["文件名"], "磐城【配音】.xlsx")
        self.assertEqual(grouped_rows[1]["key"], "Key_2")
        self.assertEqual(grouped_rows[1]["source"], "“斑鸠”&“诗人”")
        self.assertEqual(
            grouped_rows[1]["QA问题"],
            '“斑鸠” -> "Colombe"：Key Term Mismatch；“诗人” -> "Poète"：Key Term Mismatch',
        )

    def test_group_key_uses_file_name_and_source_when_metadata_has_only_file_name(self) -> None:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet["C1"] = "Source"
        worksheet["D1"] = "Target"
        worksheet["E1"] = "Comments"
        worksheet["F1"] = "Metadata"
        worksheet["A2"] = "Key Term Mismatch (提示 / Avis)"
        worksheet["C3"] = "提示"
        worksheet["D3"] = "Avis"
        worksheet["E3"] = "terms.xlsx"
        worksheet["F3"] = "UI弹窗文字.xlsx"
        worksheet["C4"] = "提示"
        worksheet["D4"] = "Avis"
        worksheet["E4"] = "terms.xlsx"
        worksheet["F4"] = "另一个文件.xlsx"

        detail_rows = collect_detail_rows(worksheet)
        grouped_rows = group_detail_rows(detail_rows)

        self.assertEqual(len(grouped_rows), 2)
        self.assertEqual(detail_rows[0].group_key, "file_source:UI弹窗文字.xlsx\x1f提示")
        self.assertEqual(detail_rows[1].group_key, "file_source:另一个文件.xlsx\x1f提示")

    def test_group_key_uses_source_when_metadata_is_empty(self) -> None:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet["A1"] = "Intro"
        worksheet["C2"] = "Source"
        worksheet["D2"] = "Target"
        worksheet["E2"] = "Comments"
        worksheet["F2"] = "Metadata"
        worksheet["A3"] = "Key Term Mismatch (提示 / Avis)"
        worksheet["C4"] = "提示"
        worksheet["D4"] = "Avis"

        detail_rows = collect_detail_rows(worksheet)

        self.assertEqual(len(detail_rows), 1)
        self.assertEqual(detail_rows[0].group_key, "source:提示")
```

- [ ] **Step 2: Run the tests to verify RED**

Run:

```bash
python -m unittest tests.test_xbench_report_transformer
```

Expected: FAIL with import errors for `collect_detail_rows`, `find_header_columns`, and `group_detail_rows`.

- [ ] **Step 3: Add row extraction and grouping implementation**

Extend `tools/xbench_report_transformer/transform_xbench_report.py` with imports, constants, dataclass, and functions:

```python
from collections import OrderedDict
from collections.abc import Iterable


OUTPUT_HEADERS = ("文件名", "key", "source", "target", "QA问题")
HEADER_SOURCE = "source"
HEADER_TARGET = "target"
HEADER_COMMENTS = "comments"
HEADER_METADATA = "metadata"
GROUP_KEY_SEPARATOR = "\x1f"


@dataclass(frozen=True)
class XbenchDetailRow:
    file_name: str
    key: str
    source: str
    target: str
    qa_issue: str
    group_key: str


def normalize_header(value: object) -> str:
    return value_to_text(value).casefold()


def find_header_columns(worksheet) -> tuple[int, dict[str, int]]:
    required_headers = {HEADER_SOURCE, HEADER_TARGET, HEADER_COMMENTS, HEADER_METADATA}
    for row_index in range(1, worksheet.max_row + 1):
        columns: dict[str, int] = {}
        for column_index in range(1, worksheet.max_column + 1):
            header = normalize_header(worksheet.cell(row_index, column_index).value)
            if header in required_headers:
                columns[header] = column_index
        if required_headers.issubset(columns):
            return row_index, columns
    expected = ", ".join(sorted(required_headers))
    raise ValueError(f"未找到 Xbench 明细表头，预期包含: {expected}")


def build_group_key(metadata: ParsedMetadata, source: str) -> str:
    if metadata.key:
        return f"key:{metadata.key}"
    if metadata.file_name:
        return f"file_source:{metadata.file_name}{GROUP_KEY_SEPARATOR}{source}"
    return f"source:{source}"


def is_qa_group_title(value: object, source: str, target: str, metadata: str) -> bool:
    title = value_to_text(value)
    return bool(title and not source and not target and not metadata)


def collect_detail_rows(worksheet) -> list[XbenchDetailRow]:
    header_row, columns = find_header_columns(worksheet)
    current_issue = XbenchIssue(issue_type="", source_term="", target_term="")
    detail_rows: list[XbenchDetailRow] = []

    for row_index in range(header_row + 1, worksheet.max_row + 1):
        first_cell = worksheet.cell(row_index, 1).value
        source = value_to_text(worksheet.cell(row_index, columns[HEADER_SOURCE]).value)
        target = value_to_text(worksheet.cell(row_index, columns[HEADER_TARGET]).value)
        metadata_text = value_to_text(worksheet.cell(row_index, columns[HEADER_METADATA]).value)

        if is_qa_group_title(first_cell, source, target, metadata_text):
            current_issue = parse_qa_title(first_cell)
            continue

        if not source and not target and not metadata_text:
            continue

        metadata = parse_metadata(metadata_text)
        detail_rows.append(
            XbenchDetailRow(
                file_name=metadata.file_name,
                key=metadata.key,
                source=source,
                target=target,
                qa_issue=format_issue_text(current_issue) if current_issue.issue_type else "",
                group_key=build_group_key(metadata, source),
            )
        )

    return detail_rows


def first_non_empty(existing: str, candidate: str) -> str:
    return existing if existing else candidate


def group_detail_rows(detail_rows: Iterable[XbenchDetailRow]) -> list[dict[str, str]]:
    grouped: OrderedDict[str, dict[str, object]] = OrderedDict()
    for detail_row in detail_rows:
        if detail_row.group_key not in grouped:
            grouped[detail_row.group_key] = {
                "文件名": detail_row.file_name,
                "key": detail_row.key,
                "source": detail_row.source,
                "target": detail_row.target,
                "issues": [],
            }
        group = grouped[detail_row.group_key]
        group["文件名"] = first_non_empty(str(group["文件名"]), detail_row.file_name)
        group["key"] = first_non_empty(str(group["key"]), detail_row.key)
        group["source"] = first_non_empty(str(group["source"]), detail_row.source)
        group["target"] = first_non_empty(str(group["target"]), detail_row.target)
        issues = group["issues"]
        if detail_row.qa_issue and detail_row.qa_issue not in issues:
            issues.append(detail_row.qa_issue)

    output_rows: list[dict[str, str]] = []
    for group in grouped.values():
        issues = group["issues"]
        output_rows.append(
            {
                "文件名": str(group["文件名"]),
                "key": str(group["key"]),
                "source": str(group["source"]),
                "target": str(group["target"]),
                "QA问题": "；".join(str(issue) for issue in issues),
            }
        )
    return output_rows
```

- [ ] **Step 4: Run the tests**

Run:

```bash
python -m unittest tests.test_xbench_report_transformer
```

Expected: PASS.

- [ ] **Step 5: Commit Task 2**

Run:

```bash
git add tools/xbench_report_transformer/transform_xbench_report.py tests/test_xbench_report_transformer.py
git commit -m "feat: group xbench report rows"
```

---

### Task 3: Workbook Output and CLI

**Files:**
- Modify: `tools/xbench_report_transformer/transform_xbench_report.py`
- Modify: `tests/test_xbench_report_transformer.py`

- [ ] **Step 1: Add failing process-level tests**

Append these imports:

```python
import tempfile
from pathlib import Path

from openpyxl import load_workbook

from tools.xbench_report_transformer.transform_xbench_report import process_excel
```

Add this test class:

```python
class ProcessExcelTests(unittest.TestCase):
    def create_report(self, path: Path) -> None:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Xbench QA"
        worksheet["A1"] = "Exported QA Report"
        worksheet["C4"] = "Source"
        worksheet["D4"] = "Target"
        worksheet["E4"] = "Comments"
        worksheet["F4"] = "Metadata"
        worksheet["A5"] = "Key Term Mismatch (提示 / Avis)"
        worksheet["C6"] = "好刻意的提示！"
        worksheet["D6"] = "Sans blague !"
        worksheet["E6"] = "terms.xlsx"
        worksheet["F6"] = "Key_1\nUI弹窗文字.xlsx"
        worksheet["A7"] = 'Key Term Mismatch (“诗人” / "Poète")'
        worksheet["C8"] = "“斑鸠”&“诗人”"
        worksheet["D8"] = '"Colombe" & "Poète"'
        worksheet["E8"] = "terms.xlsx"
        worksheet["F8"] = "Key_2\n磐城【配音】.xlsx"
        workbook.save(path)

    def test_process_excel_writes_flat_output_workbook_and_preserves_input(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "Xbench_QA_Report.xlsx"
            output_path = Path(tmp_dir) / "flat.xlsx"
            self.create_report(input_path)

            summary = process_excel(input_path, output_file=output_path)

            self.assertEqual(summary.worksheet_title, "Xbench QA")
            self.assertEqual(summary.output_path, output_path.resolve())
            self.assertEqual(summary.detail_count, 2)
            self.assertEqual(summary.grouped_count, 2)

            original = load_workbook(input_path)
            self.assertEqual(original.sheetnames, ["Xbench QA"])

            result = load_workbook(output_path)
            self.assertEqual(result.sheetnames, ["Xbench QA整理"])
            sheet = result["Xbench QA整理"]
            self.assertEqual(
                [sheet.cell(1, column).value for column in range(1, 6)],
                ["文件名", "key", "source", "target", "QA问题"],
            )
            self.assertEqual(sheet["A2"].value, "UI弹窗文字.xlsx")
            self.assertEqual(sheet["B2"].value, "Key_1")
            self.assertEqual(sheet["C2"].value, "好刻意的提示！")
            self.assertEqual(sheet["D2"].value, "Sans blague !")
            self.assertEqual(sheet["E2"].value, "提示 -> Avis：Key Term Mismatch")

    def test_process_excel_uses_default_prefixed_output_path(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "Xbench_QA_Report.xlsx"
            self.create_report(input_path)

            summary = process_excel(input_path)

            self.assertEqual(
                summary.output_path,
                (Path(tmp_dir) / "xbench_transform_Xbench_QA_Report.xlsx").resolve(),
            )
            self.assertTrue(summary.output_path.exists())

    def test_process_excel_rejects_missing_sheet(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "Xbench_QA_Report.xlsx"
            self.create_report(input_path)

            with self.assertRaisesRegex(ValueError, "工作表不存在"):
                process_excel(input_path, sheet="Missing")
```

- [ ] **Step 2: Run the tests to verify RED**

Run:

```bash
python -m unittest tests.test_xbench_report_transformer
```

Expected: FAIL with `ImportError` for `process_excel`.

- [ ] **Step 3: Add workbook output and CLI implementation**

Extend imports:

```python
import argparse
import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook

if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from tools.excel_output import build_prefixed_output_path
```

Add constants and summary dataclass:

```python
OUTPUT_SHEET_NAME = "Xbench QA整理"


@dataclass(frozen=True)
class TransformSummary:
    worksheet_title: str
    output_path: Path
    detail_count: int
    grouped_count: int
```

Add output and CLI functions:

```python
def build_default_output_path(input_path: Path) -> Path:
    return build_prefixed_output_path(input_path, "xbench_transform_")


def write_output_workbook(output_path: Path, rows: list[dict[str, str]]) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = OUTPUT_SHEET_NAME
    for column_index, header in enumerate(OUTPUT_HEADERS, start=1):
        worksheet.cell(1, column_index, header)
    for row_index, row in enumerate(rows, start=2):
        for column_index, header in enumerate(OUTPUT_HEADERS, start=1):
            worksheet.cell(row_index, column_index, row[header])
    workbook.save(output_path)


def process_excel(
    input_file: str | Path,
    sheet: str | None = None,
    output_file: str | Path | None = None,
) -> TransformSummary:
    input_path = Path(input_file).expanduser().resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"输入文件不存在: {input_path}")

    output_path = (
        Path(output_file).expanduser().resolve()
        if output_file
        else build_default_output_path(input_path)
    )

    workbook = load_workbook(input_path)
    try:
        if sheet and sheet not in workbook.sheetnames:
            raise ValueError(f"工作表不存在: {sheet}")
        worksheet = workbook[sheet] if sheet else workbook.active
        detail_rows = collect_detail_rows(worksheet)
        grouped_rows = group_detail_rows(detail_rows)
    finally:
        workbook.close()

    write_output_workbook(output_path, grouped_rows)
    return TransformSummary(
        worksheet_title=worksheet.title,
        output_path=output_path,
        detail_count=len(detail_rows),
        grouped_count=len(grouped_rows),
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="将 Xbench QA Report 转换为按 key 聚类的五列表格。")
    parser.add_argument("input_file", nargs="?", help="Xbench 导出的 QA Report xlsx 文件")
    parser.add_argument("-s", "--sheet", help="工作表名称，不填则使用当前活动工作表")
    parser.add_argument("-o", "--output", help="输出 Excel 文件路径，默认生成 xbench_transform_<原文件名>")
    return parser.parse_args()


def prompt_if_missing(args: argparse.Namespace) -> argparse.Namespace:
    interactive_mode = sys.stdin.isatty()
    if not args.input_file and not interactive_mode:
        raise ValueError("缺少输入文件路径，请传入 input_file 参数。")
    if not args.input_file:
        args.input_file = input("请输入 Xbench QA Report 文件路径: ").strip()
    if not args.sheet and interactive_mode and len(sys.argv) == 1:
        args.sheet = input("请输入工作表名称（直接回车使用当前活动工作表）: ").strip() or None
    return args


def main() -> None:
    args = prompt_if_missing(parse_args())
    summary = process_excel(
        input_file=args.input_file,
        sheet=args.sheet,
        output_file=args.output,
    )
    print("处理完成。")
    print(f"工作表: {summary.worksheet_title}")
    print(f"读取明细数: {summary.detail_count}")
    print(f"输出行数: {summary.grouped_count}")
    print(f"输出文件: {summary.output_path}")


if __name__ == "__main__":
    main()
```

- [ ] **Step 4: Run focused tests**

Run:

```bash
python -m unittest tests.test_xbench_report_transformer
```

Expected: PASS.

- [ ] **Step 5: Run the tool on the user's sample report**

Run:

```bash
python tools/xbench_report_transformer/transform_xbench_report.py "C:\Users\yizhi003\Desktop\xbench\Xbench_QA_Report.xlsx"
```

Expected output includes:

```text
处理完成。
工作表: Xbench QA
读取明细数: 249
输出行数: 223
输出文件: C:\Users\yizhi003\Desktop\xbench\xbench_transform_Xbench_QA_Report.xlsx
```

- [ ] **Step 6: Commit Task 3**

Run:

```bash
git add tools/xbench_report_transformer/transform_xbench_report.py tests/test_xbench_report_transformer.py
git commit -m "feat: write transformed xbench workbooks"
```

---

### Task 4: Documentation and Full Verification

**Files:**
- Create: `tools/xbench_report_transformer/README.md`
- Modify: `README.md`
- Modify: `docs/cli-usage.md`

- [ ] **Step 1: Add README for the new tool**

Create `tools/xbench_report_transformer/README.md`:

```markdown
# Xbench QA Report 转换

## 用途

把 ApSIC Xbench 导出的 QA Report 转换成更适合修改和筛选的五列表格：

```text
文件名, key, source, target, QA问题
```

工具不会修改原始 Xbench 报告，默认生成新的结果 Excel。

## 聚类规则

- `Metadata` 两行或更多：第一行作为 `key`，第二行作为 `文件名`。
- `Metadata` 只有一行且像文件名：`key` 留空，`文件名` 使用这一行，按 `文件名 + source` 聚类。
- `Metadata` 只有一行且不像文件名：这一行作为 `key`，按 `key` 聚类。
- `Metadata` 为空：按 `source` 聚类。

同一组内多个 QA 问题会写入同一个 `QA问题` 单元格，并用中文分号 `；` 连接。

## CLI

```bash
python3 tools/xbench_report_transformer/transform_xbench_report.py Xbench_QA_Report.xlsx
```

指定工作表和输出文件：

```bash
python3 tools/xbench_report_transformer/transform_xbench_report.py Xbench_QA_Report.xlsx \
  -s "Xbench QA" \
  -o xbench_flat.xlsx
```

默认输出文件名为 `xbench_transform_<原文件名>`。
```

- [ ] **Step 2: Update root README**

In `README.md`, add a new current-tool section near the other Excel tools:

```markdown
### Xbench QA Report 转换

- 目录：`tools/xbench_report_transformer`
- 用途：把 Xbench 导出的 QA Report 转换成 `文件名, key, source, target, QA问题` 五列表格
- 聚类：优先按 Metadata 第一行的 key 聚类；没有 key 时按文件名+source 或 source 降级聚类
- 输出方式：生成新的结果 Excel，默认文件名 `xbench_transform_<原文件名>`
- CLI：

```bash
python3 tools/xbench_report_transformer/transform_xbench_report.py Xbench_QA_Report.xlsx \
  -s "Xbench QA" \
  -o xbench_flat.xlsx
```

详情见 `tools/xbench_report_transformer/README.md`。
```

Also add `tools/xbench_report_transformer/README.md` to the documentation navigation list.

- [ ] **Step 3: Update CLI usage docs**

In `docs/cli-usage.md`, add a tool entry and example:

```markdown
## Xbench QA Report 转换

必填：`input_file`

常用可选：`-s/--sheet`、`-o/--output`

```bash
python3 tools/xbench_report_transformer/transform_xbench_report.py ./Xbench_QA_Report.xlsx \
  -s "Xbench QA" \
  -o ./xbench_flat.xlsx
```

输出列固定为：

```text
文件名, key, source, target, QA问题
```

`QA问题` 会使用 `源术语 -> 目标术语：问题类型` 格式。同一 key 下多个问题用中文分号 `；` 合并。
```

Add the default output naming bullet:

```markdown
- `xbench_transform_<原文件名>`
```

- [ ] **Step 4: Run focused tests**

Run:

```bash
python -m unittest tests.test_xbench_report_transformer
```

Expected: PASS.

- [ ] **Step 5: Run full test suite**

Run:

```bash
python -m unittest discover -s tests
```

Expected: PASS.

- [ ] **Step 6: Compile Python files**

Run:

```bash
python -m compileall -q tools tests
```

Expected: exit code 0 with no output.

- [ ] **Step 7: Check whitespace**

Run:

```bash
git diff --check
```

Expected: no output and exit code 0.

- [ ] **Step 8: Commit Task 4**

Run:

```bash
git add tools/xbench_report_transformer/README.md README.md docs/cli-usage.md
git commit -m "docs: document xbench report transformer"
```

---

## Self-Review

Spec coverage:

- Five-column output is covered in Task 3 process tests and implementation.
- Metadata fallback matrix is covered in Task 1 and Task 2 tests.
- Key grouping and fallback grouping are covered in Task 2.
- Chinese semicolon issue merge is covered in Task 2.
- Default output path and non-mutating behavior are covered in Task 3.
- CLI and README are covered in Tasks 3 and 4.

Placeholder scan:

- No TBD/TODO placeholders are present.
- Each task names concrete files, commands, expected results, and code snippets.
- Each implementation step has code content instead of vague instructions.

Type consistency:

- `ParsedMetadata`, `XbenchIssue`, `XbenchDetailRow`, and `TransformSummary` names are used consistently.
- `process_excel(...)` returns `TransformSummary` in both implementation and tests.
- Output row dictionaries consistently use `文件名`, `key`, `source`, `target`, and `QA问题`.
