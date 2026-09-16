"""Discover source containment and report missing reference translations."""

from __future__ import annotations

from dataclasses import dataclass
from collections.abc import Iterable
from pathlib import Path
import re

from openpyxl.utils import column_index_from_string

from tools.consistency_text import normalize_consistency_text
from tools.excel_output import (
    PROBLEM_BASE_HEADERS,
    existing_cell_value,
    value_row_numbers,
    validate_distinct_source_target_columns,
    write_output_table,
)
from tools.term_matching import needs_left_boundary, needs_right_boundary, span_matches_mode


PROBLEM_SHEET_NAME = "子串译文一致性"
DETAIL_LIMIT = 10
EXCERPT_LIMIT = 120
DEFAULT_MIN_CJK_CHARS = 3
DEFAULT_MIN_OTHER_CHARS = 2
MAX_MIN_CHARS = 1_000_000
MIN_POSITION_BUDGET = 256
POSITIONS_PER_CHARACTER = 2
CJK_LETTER = re.compile(
    r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff\U00020000-\U000323af"
    r"\u3040-\u30ff\u31f0-\u31ff\uff66-\uff9f"
    r"\u1100-\u11ff\u3130-\u318f\ua960-\ua97f\uac00-\ud7af\ud7b0-\ud7ff]"
)
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


def _cell_text(value: object) -> str:
    return "" if value is None else str(value)


def validate_minimum_characters(min_cjk_chars: int, min_other_chars: int) -> None:
    for value in (min_cjk_chars, min_other_chars):
        if type(value) is not int or not 1 <= value <= MAX_MIN_CHARS:
            raise ValueError(f"子串最小有效字符数必须为 1–{MAX_MIN_CHARS} 的整数。")


def _usable_reference(source: str, target: str, min_cjk_chars: int, min_other_chars: int) -> bool:
    letters = "".join(char for char in STRUCTURAL_TOKEN.sub("", source) if char.isalpha())
    minimum = min_cjk_chars if CJK_LETTER.search(letters) else min_other_chars
    # A one-character translation remains valid regardless of Source minimum.
    return (
        len(letters) >= minimum
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


def _unique_patterns(text: str, index, boundary_patterns: dict[str, re.Pattern[str]]) -> list[str]:
    """Keep first valid matches, switching strategy when positions are dense.

    AC yields by end position, longest first at equal ends. The fallback finds
    each unseen pattern's first boundary-valid occurrence, then restores that
    exact order so report details (including their first-ten limit) stay stable.
    The budget limits position enumeration, never the number of checked terms.
    """
    if not len(index):
        return []
    found: dict[str, int] = {}
    budget = max(MIN_POSITION_BUDGET, POSITIONS_PER_CHARACTER * len(text))
    for count, (last, pattern) in enumerate(index.iter(text), 1):
        if pattern not in found and span_matches_mode(
            text, last + 1 - len(pattern), last + 1, pattern, "hybrid-boundary"
        ):
            found[pattern] = last + 1
        if count >= budget:
            for candidate in index.keys():
                if candidate in found or len(candidate) > len(text):
                    continue
                left, right = needs_left_boundary(candidate), needs_right_boundary(candidate)
                if left or right:
                    regex = boundary_patterns.get(candidate)
                    if regex is None:
                        expression = (r"(?<![A-Za-z0-9_])" if left else "") + re.escape(candidate)
                        expression += r"(?![A-Za-z0-9_])" if right else ""
                        regex = boundary_patterns[candidate] = re.compile(expression)
                    match = regex.search(text)
                    if match:
                        found[candidate] = match.end()
                else:
                    start = text.find(candidate)
                    if start >= 0:
                        found[candidate] = start + len(candidate)
            return sorted(found, key=lambda item: (found[item], -len(item)))
    return list(found)


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
    min_cjk_chars: int = DEFAULT_MIN_CJK_CHARS,
    min_other_chars: int = DEFAULT_MIN_OTHER_CHARS,
    checked_source_terms: Iterable[str] = (),
) -> CheckSummary:
    """Inspect one worksheet without saving or modifying its business cells."""
    if start_row < 1:
        raise ValueError("开始行必须大于等于 1。")
    validate_minimum_characters(min_cjk_chars, min_other_chars)
    source_column = source_column.strip().upper()
    target_column = target_column.strip().upper()
    column_index_from_string(source_column)
    column_index_from_string(target_column)
    validate_distinct_source_target_columns(source_column, target_column)
    worksheet = workbook[sheet] if sheet else workbook.active
    row_numbers = value_row_numbers(worksheet, (source_column, target_column), start_row=start_row)
    last_row = row_numbers[-1] if row_numbers else start_row - 1
    source_index = column_index_from_string(source_column)
    target_index = column_index_from_string(target_column)
    # Only whole Source equality excludes a row. Never remove term spans from
    # otherwise useful sentences. Casefold matches terminology's case policy.
    excluded_sources = {normalize_consistency_text(term).casefold() for term in checked_source_terms}
    groups: dict[str, dict[str, list[Occurrence]]] = {}
    for row in row_numbers:
        source = _cell_text(existing_cell_value(worksheet, row, source_index))
        source_key = normalize_consistency_text(source)
        if source_key and source_key.casefold() not in excluded_sources:
            target = _cell_text(existing_cell_value(worksheet, row, target_index))
            target_key = normalize_consistency_text(target)
            groups.setdefault(source_key, {}).setdefault(target_key, []).append(
                Occurrence(row, source, target)
            )

    references: dict[str, tuple[str, list[Occurrence]]] = {}
    conflicting = 0
    for source, variants in groups.items():
        if len(variants) != 1:
            conflicting += 1
            continue
        target, rows = next(iter(variants.items()))
        if _usable_reference(source, target, min_cjk_chars, min_other_chars):
            references[source] = (target, rows)

    source_index = _build_index(references)
    target_index = _build_index({target.casefold() for target, _ in references.values()})
    boundary_patterns: dict[str, re.Pattern[str]] = {}
    problem_entries = []
    problem_count = 0
    for source, variants in groups.items():
        # Keep each child once within this parent. No global graph
        # or cross-product of duplicate row numbers is materialized.
        children = [child for child in _unique_patterns(source, source_index, boundary_patterns)
                    if len(child) < len(source)]
        if not children:
            continue
        for target, rows in variants.items():
            if not target.strip():
                continue
            present = set(_unique_patterns(target.casefold(), target_index, boundary_patterns))
            details = []
            missing = 0
            for child in children:
                reference_target, reference_rows = references[child]
                if reference_target.casefold() in present:
                    continue
                missing += 1
                if len(details) < DETAIL_LIMIT:
                    row_text = "、".join(str(item.row) for item in reference_rows[:DETAIL_LIMIT])
                    if len(reference_rows) > DETAIL_LIMIT:
                        row_text += f"…（共 {len(reference_rows)} 行，行号已截断）"
                    first = reference_rows[0]
                    details.append(
                        f"“{_excerpt(first.source)}” → “{_excerpt(first.target)}”"
                        f"（参考原表第 {row_text} 行）"
                    )
            if missing:
                description = "\n".join(details)
                if missing > DETAIL_LIMIT:
                    description += f"\n另 {missing - DETAIL_LIMIT} 条未展示。"
                for occurrence in rows:
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
