"""Discover source containment and report missing reference translations."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re

from openpyxl.utils import column_index_from_string

from tools.consistency_text import normalized_text_with_offset, normalize_consistency_text
from tools.excel_output import (
    PROBLEM_BASE_HEADERS,
    find_last_value_row,
    validate_distinct_source_target_columns,
    write_output_table,
)
from tools.term_matching import span_matches_mode


PROBLEM_SHEET_NAME = "子串译文一致性"
DETAIL_LIMIT = 10
EXCERPT_LIMIT = 120
# Remove common structural tokens only when deciding whether a reference has
# useful text. Actual matching uses the complete normalized segment.
STRUCTURAL_TOKEN = re.compile(
    r"<[^<>]*>|\{[^{}]*\}|\[[^\[\]]*\]|"
    r"%(?:\d+\$)?[-+#0 ]*\d*(?:\.\d+)?[a-zA-Z%]"
)


@dataclass(frozen=True)
class CheckSummary:
    output_path: Path
    worksheet_title: str
    source_column: str
    target_column: str
    start_row: int
    total_rows_checked: int
    reference_count: int
    skipped_conflicting_sources: int
    problem_count: int
    problem_rows: int


@dataclass(frozen=True)
class Occurrence:
    row: int
    source: str
    target: str
    source_offset: int


def _cell_text(value: object) -> str:
    return "" if value is None else str(value)


def _usable_reference(source: str, target: str) -> bool:
    # Two letters (including CJK) excludes single characters, numeric-only
    # segments and punctuation. A one-character translation remains valid.
    return (
        sum(char.isalpha() for char in STRUCTURAL_TOKEN.sub("", source)) >= 2
        and any(char.isalpha() for char in STRUCTURAL_TOKEN.sub("", target))
    )


def _build_index(patterns):
    # The C extension is a declared dependency; defer loading until execution.
    import ahocorasick

    automaton = ahocorasick.Automaton()
    for text in patterns:
        automaton.add_word(text, text)
    if len(automaton):
        automaton.make_automaton()
    return automaton


def _matches(text: str, index):
    if not len(index):
        return
    for last, pattern in index.iter(text):
        start, end = last + 1 - len(pattern), last + 1
        if span_matches_mode(text, start, end, pattern, "hybrid-boundary"):
            yield pattern, start, end


def _excerpt(text: str) -> str:
    return text if len(text) <= EXCERPT_LIMIT else text[:EXCERPT_LIMIT] + "…（已截断）"


def process_workbook(
    *,
    workbook,
    output_path: Path,
    source_column: str,
    target_column: str,
    sheet: str | None = None,
    start_row: int = 2,
    format_output: bool = True,
) -> CheckSummary:
    """Inspect one worksheet without saving or modifying its business cells."""
    if start_row < 1:
        raise ValueError("开始行必须大于等于 1。")
    source_column = source_column.strip().upper()
    target_column = target_column.strip().upper()
    column_index_from_string(source_column)
    column_index_from_string(target_column)
    validate_distinct_source_target_columns(source_column, target_column)
    worksheet = workbook[sheet] if sheet else workbook.active
    last_row = find_last_value_row(worksheet, (source_column, target_column), start_row=start_row)
    groups: dict[str, dict[str, list[Occurrence]]] = {}
    for row in range(start_row, last_row + 1):
        source = _cell_text(worksheet[f"{source_column}{row}"].value)
        source_key, source_offset = normalized_text_with_offset(source)
        if source_key:
            target = _cell_text(worksheet[f"{target_column}{row}"].value)
            target_key = normalize_consistency_text(target)
            groups.setdefault(source_key, {}).setdefault(target_key, []).append(
                Occurrence(row, source, target, source_offset)
            )

    references: dict[str, tuple[str, list[Occurrence]]] = {}
    conflicting = 0
    for source, variants in groups.items():
        if len(variants) != 1:
            conflicting += 1
            continue
        target, rows = next(iter(variants.items()))
        if _usable_reference(source, target):
            references[source] = (target, rows)

    source_index = _build_index(references)
    target_index = _build_index({target.casefold() for target, _ in references.values()})
    problem_entries = []
    problem_count = 0
    for source, variants in groups.items():
        # Keep only one position per child within this parent. No global graph
        # or cross-product of duplicate row numbers is materialized.
        children: dict[str, tuple[int, int]] = {}
        for child, start, end in _matches(source, source_index):
            if len(child) < len(source):
                children.setdefault(child, (start, end))
        if not children:
            continue
        for target, rows in variants.items():
            if not target.strip():
                continue
            present = {pattern for pattern, _, _ in _matches(target.casefold(), target_index)}
            details = []
            missing = 0
            for child, (start, end) in children.items():
                reference_target, reference_rows = references[child]
                if reference_target.casefold() in present:
                    continue
                missing += 1
                if len(details) < DETAIL_LIMIT:
                    row_text = "、".join(str(item.row) for item in reference_rows[:DETAIL_LIMIT])
                    if len(reference_rows) > DETAIL_LIMIT:
                        row_text += f"…（共 {len(reference_rows)} 行，行号已截断）"
                    first = reference_rows[0]
                    reference_detail = (
                        f"参考行 {row_text}；Source：{_excerpt(first.source)}；"
                        f"参考 Target：{_excerpt(first.target)}；"
                    )
                    details.append((reference_detail, start, end))
            if missing:
                for occurrence in rows:
                    description = (
                        f"疑似子串译文不一致：{missing} 个原文片段的参考译文未匹配到，请复核。"
                        "语序、词形或上下文差异可能属于合理翻译。\n"
                        + "\n".join(
                            f"{detail}原文位置：第 {start + occurrence.source_offset + 1}"
                            f"–{end + occurrence.source_offset} 字符"
                            for detail, start, end in details
                        )
                    )
                    if missing > DETAIL_LIMIT:
                        description += f"\n仅展示前 {DETAIL_LIMIT} 条，另 {missing - DETAIL_LIMIT} 条未展示。"
                    problem_entries.append((
                        occurrence.row, occurrence.source, occurrence.target, description,
                    ))
                problem_count += missing * len(rows)

    problem_entries.sort(key=lambda entry: entry[0])
    write_output_table(
        workbook,
        current_sheet_name=worksheet.title,
        sheet_name=PROBLEM_SHEET_NAME,
        headers=PROBLEM_BASE_HEADERS,
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
        total_rows_checked=max(0, last_row - start_row + 1),
        reference_count=len(references),
        skipped_conflicting_sources=conflicting,
        problem_count=problem_count,
        problem_rows=len(problem_entries),
    )
