from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Iterable

FILE_NAME_EXTENSIONS = (".xlsx", ".xls", ".xlsm", ".csv", ".txt")
OUTPUT_HEADERS = ("文件名", "key", "source", "target", "QA问题")
HEADER_SOURCE = "source"
HEADER_TARGET = "target"
HEADER_COMMENTS = "comments"
HEADER_METADATA = "metadata"
GROUP_KEY_SEPARATOR = "\x1f"
REQUIRED_HEADERS = (HEADER_COMMENTS, HEADER_METADATA, HEADER_SOURCE, HEADER_TARGET)
QA_TITLE_PATTERN = re.compile(r"^\s*(?P<issue_type>.*?)\s*\((?P<terms>.*)\)\s*$")


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


def value_to_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def normalize_header(value: object) -> str:
    return value_to_text(value).casefold()


def looks_like_file_name(value: str) -> bool:
    return value.casefold().endswith(FILE_NAME_EXTENSIONS)


def parse_metadata(value: object) -> ParsedMetadata:
    lines = [line.strip() for line in value_to_text(value).splitlines() if line.strip()]
    if len(lines) >= 2:
        return ParsedMetadata(key=lines[0], file_name=lines[1])
    if not lines:
        return ParsedMetadata(key="", file_name="")
    line = lines[0]
    if looks_like_file_name(line):
        return ParsedMetadata(key="", file_name=line)
    return ParsedMetadata(key=line, file_name="")


def parse_qa_title(title: object) -> XbenchIssue:
    text = value_to_text(title)
    match = QA_TITLE_PATTERN.match(text)
    if match is None:
        return XbenchIssue(issue_type=text, source_term="", target_term="")
    source_term, separator, target_term = match.group("terms").partition(" / ")
    if not separator:
        return XbenchIssue(issue_type=text, source_term="", target_term="")
    return XbenchIssue(
        issue_type=match.group("issue_type").strip(),
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


def find_header_columns(worksheet) -> tuple[int, dict[str, int]]:
    for row in worksheet.iter_rows():
        columns: dict[str, int] = {}
        for cell in row:
            header = normalize_header(cell.value)
            if header in REQUIRED_HEADERS:
                columns[header] = cell.column
        if all(header in columns for header in REQUIRED_HEADERS):
            return row[0].row, columns
    raise ValueError("未找到 Xbench 明细表头，预期包含: comments, metadata, source, target")


def build_group_key(metadata: ParsedMetadata, source: str) -> str:
    if metadata.key:
        return f"key:{metadata.key}"
    if metadata.file_name:
        return f"file_source:{metadata.file_name}{GROUP_KEY_SEPARATOR}{source}"
    return f"source:{source}"


def is_qa_group_title(value: object, source: str, target: str, metadata: str) -> bool:
    return bool(value_to_text(value)) and not source and not target and not metadata


def choose_qa_group_title(first_cell: object, comments: str, source: str, target: str, metadata: str) -> str:
    first_cell_text = value_to_text(first_cell)
    if first_cell_text and not source and not target and not metadata:
        return first_cell_text
    if comments and not first_cell_text and not source and not target and not metadata:
        return comments
    return ""


def collect_detail_rows(worksheet) -> list[XbenchDetailRow]:
    header_row, columns = find_header_columns(worksheet)
    detail_rows: list[XbenchDetailRow] = []
    current_issue = XbenchIssue(issue_type="", source_term="", target_term="")

    for row_index in range(header_row + 1, worksheet.max_row + 1):
        row_values = [value_to_text(cell.value) for cell in worksheet[row_index]]
        if not any(row_values):
            continue

        first_cell = worksheet.cell(row=row_index, column=1).value
        source = value_to_text(worksheet.cell(row=row_index, column=columns[HEADER_SOURCE]).value)
        target = value_to_text(worksheet.cell(row=row_index, column=columns[HEADER_TARGET]).value)
        comments = value_to_text(worksheet.cell(row=row_index, column=columns[HEADER_COMMENTS]).value)
        metadata_text = value_to_text(worksheet.cell(row=row_index, column=columns[HEADER_METADATA]).value)

        group_title = choose_qa_group_title(first_cell, comments, source, target, metadata_text)
        if group_title:
            current_issue = parse_qa_title(group_title)
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
                qa_issue=format_issue_text(current_issue),
                group_key=build_group_key(metadata, source),
            )
        )

    return detail_rows


def first_non_empty(existing: str, candidate: str) -> str:
    return existing if existing else candidate


def group_detail_rows(detail_rows: Iterable[XbenchDetailRow]) -> list[dict[str, str]]:
    grouped: dict[str, dict[str, str]] = {}
    grouped_issues: dict[str, list[str]] = {}

    for detail_row in detail_rows:
        if detail_row.group_key not in grouped:
            grouped[detail_row.group_key] = dict.fromkeys(OUTPUT_HEADERS, "")
            grouped_issues[detail_row.group_key] = []

        output_row = grouped[detail_row.group_key]
        output_row["文件名"] = first_non_empty(output_row["文件名"], detail_row.file_name)
        output_row["key"] = first_non_empty(output_row["key"], detail_row.key)
        output_row["source"] = first_non_empty(output_row["source"], detail_row.source)
        output_row["target"] = first_non_empty(output_row["target"], detail_row.target)

        issues = grouped_issues[detail_row.group_key]
        if detail_row.qa_issue and detail_row.qa_issue not in issues:
            issues.append(detail_row.qa_issue)
            output_row["QA问题"] = "；".join(issues)

    return list(grouped.values())
