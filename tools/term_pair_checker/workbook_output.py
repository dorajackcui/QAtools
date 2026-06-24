"""Workbook output helpers for term pair checking."""

from __future__ import annotations

from collections.abc import Iterable
from pathlib import Path
from typing import TYPE_CHECKING

from tools.excel_output import (
    ROW_PROBLEM_SEPARATOR,
    build_prefixed_output_path,
    format_row_problem_text,
    rebuild_output_sheet,
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
        term_sheet[f"A{row_index}"] = term_pair.source_display_text
        term_sheet[f"B{row_index}"] = term_pair.target_display_text
        term_sheet[f"C{row_index}"] = term_pair.source_plain_text
        term_sheet[f"D{row_index}"] = term_pair.target_plain_text
        term_sheet[f"E{row_index}"] = term_pair.term_source


def write_problem_sheet(workbook, worksheet_title: str, problem_entries: Iterable["ProblemEntry"]) -> None:
    problem_sheet = rebuild_output_sheet(workbook, worksheet_title, PROBLEM_SHEET_NAME)
    problem_sheet["A1"] = "问题行号"
    problem_sheet["B1"] = "问题source术语"
    problem_sheet["C1"] = "预期target术语"
    problem_sheet["D1"] = "术语来源"
    problem_sheet["E1"] = "问题简述"
    problem_sheet["F1"] = "source原文"
    problem_sheet["G1"] = "target原文"
    for row_index, problem_entry in enumerate(problem_entries, start=2):
        problem_sheet[f"A{row_index}"] = problem_entry.row_index
        problem_sheet[f"B{row_index}"] = problem_entry.problem_source_term
        problem_sheet[f"C{row_index}"] = problem_entry.expected_target_term
        problem_sheet[f"D{row_index}"] = problem_entry.term_source
        problem_sheet[f"E{row_index}"] = problem_entry.description
        problem_sheet[f"F{row_index}"] = problem_entry.source_snapshot
        problem_sheet[f"G{row_index}"] = problem_entry.target_snapshot


def delete_legacy_term_sheets(workbook) -> None:
    for sheet_name in LEGACY_TERM_SHEET_NAMES:
        if sheet_name in workbook.sheetnames:
            del workbook[sheet_name]
