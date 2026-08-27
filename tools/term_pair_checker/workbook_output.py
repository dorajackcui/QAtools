"""Workbook output helpers for term pair checking."""

from __future__ import annotations

from collections.abc import Iterable
from pathlib import Path
from typing import TYPE_CHECKING

from tools.excel_output import (
    PROBLEM_BASE_HEADERS,
    ROW_PROBLEM_SEPARATOR,
    build_prefixed_output_path,
    format_row_problem_text,
    join_unique_text,
    rebuild_output_sheet,
    write_output_table,
)

if TYPE_CHECKING:
    from tools.term_pair_checker.extract_terms_from_excel import (
        ProblemEntry,
        RecordedTermPair,
    )


TERM_SHEET_NAME = "术语表"
PROBLEM_SHEET_NAME = "问题列"
LEGACY_TERM_SHEET_NAMES = ("术语表（无mark）", "术语汇总")


def build_default_output_path(input_path: Path) -> Path:
    return build_prefixed_output_path(input_path, "term_pair_check_")


def build_row_problem_summaries(problem_entries: Iterable["ProblemEntry"]) -> dict[int, str]:
    summaries_by_row: dict[int, list[str]] = {}
    for problem_entry in problem_entries:
        summary = format_row_problem_text(
            problem_entry.problem_source_term,
            problem_entry.expected_target_term,
            problem_entry.description,
        )
        if summary:
            summaries_by_row.setdefault(problem_entry.row_index, []).append(summary)
    return {
        row_index: ROW_PROBLEM_SEPARATOR.join(summaries)
        for row_index, summaries in summaries_by_row.items()
    }


def write_term_sheet(workbook, worksheet_title: str, term_pairs: Iterable["RecordedTermPair"]) -> None:
    term_sheet = rebuild_output_sheet(workbook, worksheet_title, TERM_SHEET_NAME)
    term_sheet["A1"] = "source术语"
    term_sheet["B1"] = "target术语"
    term_sheet["C1"] = "source术语（无mark）"
    term_sheet["D1"] = "target术语（无mark）"
    term_sheet["E1"] = "术语来源"
    for row_index, term_pair in enumerate(term_pairs, start=2):
        for column, value in zip(
            ("A", "B", "C", "D", "E"),
            (
                term_pair.source_display_text,
                term_pair.target_display_text,
                term_pair.source_plain_text,
                term_pair.target_plain_text,
                term_pair.term_source,
            ),
        ):
            if value:
                term_sheet[f"{column}{row_index}"] = value


def write_problem_sheet(
    workbook,
    worksheet_title: str,
    target_column: str,
    problem_entries: Iterable["ProblemEntry"],
) -> None:
    entries_by_row: dict[int, list["ProblemEntry"]] = {}
    for problem_entry in problem_entries:
        entries_by_row.setdefault(problem_entry.row_index, []).append(problem_entry)

    rows = []
    for source_row, row_entries in entries_by_row.items():
        first_entry = row_entries[0]
        descriptions = (
            format_row_problem_text(
                entry.problem_source_term,
                entry.expected_target_term,
                entry.description,
            )
            for entry in row_entries
        )
        rows.append(
            (
                source_row,
                first_entry.source_snapshot,
                first_entry.target_snapshot,
                join_unique_text(descriptions),
                join_unique_text(entry.problem_source_term for entry in row_entries),
                join_unique_text(entry.expected_target_term for entry in row_entries),
                join_unique_text(entry.term_source for entry in row_entries),
            )
        )

    write_output_table(
        workbook,
        current_sheet_name=worksheet_title,
        sheet_name=PROBLEM_SHEET_NAME,
        headers=PROBLEM_BASE_HEADERS
        + ("source术语", "预期target术语", "术语来源"),
        rows=rows,
        row_link_target_column=target_column,
    )


def delete_legacy_term_sheets(workbook) -> None:
    for sheet_name in LEGACY_TERM_SHEET_NAMES:
        if sheet_name in workbook.sheetnames:
            del workbook[sheet_name]
