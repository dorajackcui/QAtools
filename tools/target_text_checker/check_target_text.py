#!/usr/bin/env python3
"""Check configurable text conventions in Excel target cells."""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from openpyxl.utils import column_index_from_string

from tools.excel_output import (
    PROBLEM_BASE_HEADERS,
    build_prefixed_output_path,
    find_last_value_row,
    load_workbook_for_editing,
    validate_distinct_source_target_columns,
    write_output_table,
)


PROBLEM_SHEET_NAME = "Target文本规范问题"

ABNORMAL_PUNCTUATION_RULE = "abnormal-punctuation"
# Python callers that imported the original constant continue to select the
# upgraded rule. The old CLI value is normalized separately below.
ABNORMAL_ELLIPSIS_RULE = ABNORMAL_PUNCTUATION_RULE
LEGACY_ABNORMAL_ELLIPSIS_RULE = "abnormal-ellipsis"
CONSECUTIVE_SPACES_RULE = "consecutive-spaces"
LEADING_TRAILING_SPACES_RULE = "leading-trailing-spaces"
MIXED_WIDTH_RULE = "mixed-width"
SUPPORTED_RULES = (
    ABNORMAL_PUNCTUATION_RULE,
    CONSECUTIVE_SPACES_RULE,
    LEADING_TRAILING_SPACES_RULE,
    MIXED_WIDTH_RULE,
)
RULE_ALIASES = {
    LEGACY_ABNORMAL_ELLIPSIS_RULE: ABNORMAL_PUNCTUATION_RULE,
}
SUPPORTED_RULE_INPUTS = SUPPORTED_RULES + tuple(RULE_ALIASES)
RULE_LABELS = {
    ABNORMAL_PUNCTUATION_RULE: "异常标点符号",
    CONSECUTIVE_SPACES_RULE: "连续空格",
    LEADING_TRAILING_SPACES_RULE: "首尾空格",
    MIXED_WIDTH_RULE: "全半角混用",
}

_REPEATED_PUNCTUATION_PATTERN = re.compile(
    r"[.．。]{2,}|[,，、]{2,}|[:：]{2,}|[;；]{2,}"
)
_CONSECUTIVE_SPACE_RUN_PATTERN = re.compile(r" {2,}")

# Only compare equivalent character families. This avoids flagging ordinary text
# merely because, for example, it contains ASCII words and Chinese punctuation.
_WIDTH_FAMILIES = (
    ("逗号", frozenset(","), frozenset("，")),
    ("句号", frozenset("."), frozenset("．。")),
    ("冒号", frozenset(":"), frozenset("：")),
    ("分号", frozenset(";"), frozenset("；")),
    ("问号", frozenset("?"), frozenset("？")),
    ("感叹号", frozenset("!"), frozenset("！")),
    ("圆括号", frozenset("()"), frozenset("（）")),
    ("方括号", frozenset("[]"), frozenset("［］")),
    ("花括号", frozenset("{}"), frozenset("｛｝")),
    ("尖括号", frozenset("<>"), frozenset("＜＞")),
    (
        "字母",
        frozenset("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"),
        frozenset(
            "ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ"
            "ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ"
        ),
    ),
    ("数字", frozenset("0123456789"), frozenset("０１２３４５６７８９")),
)


@dataclass(frozen=True)
class TextIssue:
    rule: str
    issue_type: str
    description: str
    matched_content: str


@dataclass(frozen=True)
class CheckSummary:
    output_path: Path
    worksheet_title: str
    source_column: str
    target_column: str
    start_row: int
    selected_rules: tuple[str, ...]
    processed_count: int
    problem_count: int
    problem_rows: int


def normalize_column(column_name: str) -> str:
    normalized = column_name.strip().upper()
    column_index_from_string(normalized)
    return normalized


def cell_text(value: object) -> str:
    return "" if value is None else str(value)


def build_default_output_path(input_file: str | Path) -> Path:
    return build_prefixed_output_path(input_file, "target_text_check_")


def normalize_rules(rules: Iterable[str] | None) -> tuple[str, ...]:
    requested = SUPPORTED_RULES if rules is None else rules
    selected = tuple(
        dict.fromkeys(RULE_ALIASES.get(rule, rule) for rule in requested)
    )
    unknown = [rule for rule in selected if rule not in SUPPORTED_RULES]
    if unknown:
        raise ValueError("不支持的 Target 文本规范规则：" + "、".join(unknown))
    if not selected:
        raise ValueError("Target 文本规范检查至少需要选择一项规则。")
    return selected


def _format_unique(values: Iterable[str]) -> str:
    return "、".join(dict.fromkeys(values))


def _find_abnormal_punctuation(text: str) -> TextIssue | None:
    matches = [
        match.group(0) for match in _REPEATED_PUNCTUATION_PATTERN.finditer(text)
    ]
    # Three ASCII dots and the Unicode ellipsis character are accepted forms.
    # All other repeated punctuation sequences covered by the pattern are
    # treated as abnormal. Exclamation and question marks are intentionally
    # absent from the pattern.
    abnormal_matches = [match for match in matches if match != "..."]
    if not abnormal_matches:
        return None
    return TextIssue(
        rule=ABNORMAL_PUNCTUATION_RULE,
        issue_type=RULE_LABELS[ABNORMAL_PUNCTUATION_RULE],
        description="Target 中存在异常或重复标点符号。",
        matched_content=_format_unique(abnormal_matches),
    )


def _find_consecutive_spaces(text: str) -> TextIssue | None:
    matches = [
        match.group(0) for match in _CONSECUTIVE_SPACE_RUN_PATTERN.finditer(text)
    ]
    if not matches:
        return None
    lengths = _format_unique(f"{len(match)} 个空格" for match in matches)
    return TextIssue(
        rule=CONSECUTIVE_SPACES_RULE,
        issue_type=RULE_LABELS[CONSECUTIVE_SPACES_RULE],
        description="Target 中存在连续空格。",
        matched_content=lengths,
    )


def _find_leading_trailing_spaces(text: str) -> TextIssue | None:
    leading_count = len(text) - len(text.lstrip(" "))
    trailing_count = len(text) - len(text.rstrip(" "))
    if not leading_count and not trailing_count:
        return None

    if leading_count == len(text):
        matched_content = f"首尾 {leading_count} 个空格"
    else:
        positions = []
        if leading_count:
            positions.append(f"开头 {leading_count} 个空格")
        if trailing_count:
            positions.append(f"结尾 {trailing_count} 个空格")
        matched_content = "、".join(positions)

    return TextIssue(
        rule=LEADING_TRAILING_SPACES_RULE,
        issue_type=RULE_LABELS[LEADING_TRAILING_SPACES_RULE],
        description="Target 开头或结尾存在普通空格。",
        matched_content=matched_content,
    )


def _find_mixed_width(text: str) -> TextIssue | None:
    characters = set(text)
    mixed_families: list[str] = []
    matched_pairs: list[str] = []
    for family_name, halfwidth, fullwidth in _WIDTH_FAMILIES:
        half_matches = sorted(characters & halfwidth)
        full_matches = sorted(characters & fullwidth)
        if not half_matches or not full_matches:
            continue
        mixed_families.append(family_name)
        matched_pairs.append(
            f"{family_name}（半角 {''.join(half_matches)} / 全角 {''.join(full_matches)}）"
        )
    if not mixed_families:
        return None
    return TextIssue(
        rule=MIXED_WIDTH_RULE,
        issue_type=RULE_LABELS[MIXED_WIDTH_RULE],
        description="Target 中存在同类字符的全半角混用。",
        matched_content="；".join(matched_pairs),
    )


_RULE_CHECKERS = {
    ABNORMAL_PUNCTUATION_RULE: _find_abnormal_punctuation,
    CONSECUTIVE_SPACES_RULE: _find_consecutive_spaces,
    LEADING_TRAILING_SPACES_RULE: _find_leading_trailing_spaces,
    MIXED_WIDTH_RULE: _find_mixed_width,
}


def find_text_issues(
    value: object,
    *,
    rules: Iterable[str] | None = None,
) -> tuple[TextIssue, ...]:
    selected_rules = normalize_rules(rules)
    if not isinstance(value, str) or not value:
        return ()
    return tuple(
        issue
        for rule in selected_rules
        if (issue := _RULE_CHECKERS[rule](value)) is not None
    )


def process_excel(
    input_file: str | Path,
    source_column: str,
    target_column: str,
    sheet: str | None = None,
    start_row: int = 2,
    rules: Iterable[str] | None = None,
    output_file: str | Path | None = None,
) -> CheckSummary:
    if start_row < 1:
        raise ValueError("开始行必须大于等于 1。")

    input_path = Path(input_file).expanduser().resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"输入文件不存在: {input_path}")
    source_column = normalize_column(source_column)
    target_column = normalize_column(target_column)
    validate_distinct_source_target_columns(source_column, target_column)
    rules = normalize_rules(rules)
    output_path = (
        Path(output_file).expanduser().resolve()
        if output_file
        else build_default_output_path(input_path).resolve()
    )

    workbook = load_workbook_for_editing(input_path)
    try:
        summary = process_workbook(
            workbook=workbook,
            output_path=output_path,
            source_column=source_column,
            target_column=target_column,
            sheet=sheet,
            start_row=start_row,
            rules=rules,
        )
        workbook.save(output_path)
        return summary
    finally:
        workbook.close()


def process_workbook(
    *,
    workbook,
    output_path: Path,
    source_column: str,
    target_column: str,
    sheet: str | None = None,
    start_row: int = 2,
    rules: Iterable[str] | None = None,
    format_output: bool = True,
) -> CheckSummary:
    """Run the selected checks against an already-open workbook without saving it."""
    if start_row < 1:
        raise ValueError("开始行必须大于等于 1。")

    source_column = normalize_column(source_column)
    target_column = normalize_column(target_column)
    validate_distinct_source_target_columns(source_column, target_column)
    selected_rules = normalize_rules(rules)
    worksheet = workbook[sheet] if sheet else workbook.active
    processed_count = 0
    problem_rows: set[int] = set()
    problem_entries: list[tuple[object, ...]] = []
    last_row = find_last_value_row(
        worksheet,
        (source_column, target_column),
        start_row=start_row,
    )
    for row_index in range(start_row, last_row + 1):
        source_value = worksheet[f"{source_column}{row_index}"].value
        target_value = worksheet[f"{target_column}{row_index}"].value
        issues = find_text_issues(target_value, rules=selected_rules)
        processed_count += 1
        for issue in issues:
            problem_rows.add(row_index)
            problem_entries.append(
                (
                    row_index,
                    cell_text(source_value),
                    cell_text(target_value),
                    issue.description,
                    issue.issue_type,
                    issue.matched_content,
                )
            )

    write_output_table(
        workbook,
        current_sheet_name=worksheet.title,
        sheet_name=PROBLEM_SHEET_NAME,
        headers=PROBLEM_BASE_HEADERS + ("问题类型", "命中内容"),
        rows=problem_entries,
        row_link_target_column=target_column,
        format_output=format_output,
    )
    return CheckSummary(
        output_path=output_path,
        worksheet_title=worksheet.title,
        source_column=source_column,
        target_column=target_column,
        start_row=start_row,
        selected_rules=selected_rules,
        processed_count=processed_count,
        problem_count=len(problem_entries),
        problem_rows=len(problem_rows),
    )


__all__ = [
    "ABNORMAL_ELLIPSIS_RULE",
    "ABNORMAL_PUNCTUATION_RULE",
    "CONSECUTIVE_SPACES_RULE",
    "LEADING_TRAILING_SPACES_RULE",
    "MIXED_WIDTH_RULE",
    "PROBLEM_SHEET_NAME",
    "RULE_LABELS",
    "SUPPORTED_RULE_INPUTS",
    "SUPPORTED_RULES",
    "CheckSummary",
    "TextIssue",
    "build_default_output_path",
    "find_text_issues",
    "normalize_rules",
    "process_excel",
    "process_workbook",
]
