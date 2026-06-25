from __future__ import annotations

import re
from dataclasses import dataclass

FILE_NAME_EXTENSIONS = (".xlsx", ".xls", ".xlsm", ".csv", ".txt")
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


def value_to_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


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
