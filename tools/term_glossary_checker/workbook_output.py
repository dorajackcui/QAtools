"""Workbook output helpers for glossary-based term checking."""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING

from tools.excel_output import (
    ROW_PROBLEM_SEPARATOR,
    build_prefixed_output_path,
    format_row_problem_text,
    rebuild_output_sheet,
)
from tools.term_matching import normalize_text

if TYPE_CHECKING:
    from tools.term_glossary_checker.check_terms_against_glossary import (
        CheckSummary,
        GlossaryConflict,
    )


PROBLEM_SHEET_NAME = "术语命中问题"
SUMMARY_SHEET_NAME = "检查汇总"
GlossaryProblemEntry = tuple[int, str, str, str, str, str]


def build_default_output_path(data_path: Path) -> Path:
    return build_prefixed_output_path(data_path, "glossary_check_")



def write_problem_sheet(workbook, worksheet_title: str, problem_entries: list[GlossaryProblemEntry]) -> None:
    problem_sheet = rebuild_output_sheet(workbook, worksheet_title, PROBLEM_SHEET_NAME)
    headers = ["行号", "问题类型", "source术语", "期望target术语", "source文本", "target文本"]
    for column_index, header in enumerate(headers, start=1):
        problem_sheet.cell(1, column_index, header)

    sorted_problem_entries = sorted(
        problem_entries,
        key=lambda entry: (
            entry[2] == "",
            normalize_text(entry[2], case_sensitive=False),
            entry[0],
        ),
    )
    for row_index, entry in enumerate(sorted_problem_entries, start=2):
        for column_index, value in enumerate(entry, start=1):
            problem_sheet.cell(row_index, column_index, value)


def build_row_problem_summaries(problem_entries: list[GlossaryProblemEntry]) -> dict[int, str]:
    summaries_by_row: dict[int, list[str]] = {}
    for row_index, problem_type, source_term, expected_target_term, _source_text, _target_text in problem_entries:
        summary = format_row_problem_text(source_term, expected_target_term, problem_type)
        if summary:
            summaries_by_row.setdefault(row_index, []).append(summary)
    return {
        row_index: ROW_PROBLEM_SEPARATOR.join(summaries)
        for row_index, summaries in summaries_by_row.items()
    }


def write_summary_sheet(
    workbook,
    worksheet_title: str,
    summary: "CheckSummary",
    conflicts: list["GlossaryConflict"],
    glossary_source_column: str,
    glossary_target_column: str,
    data_source_column: str,
    data_target_column: str,
    start_row: int,
    match_mode: str,
) -> None:
    summary_sheet = rebuild_output_sheet(workbook, worksheet_title, SUMMARY_SHEET_NAME)
    summary_sheet["A1"] = "统计项"
    summary_sheet["B1"] = "值"

    summary_rows = [
        ("术语表工作表", summary.glossary_sheet_title),
        ("检查工作表", summary.data_sheet_title),
        ("术语表source列", glossary_source_column),
        ("术语表target列", glossary_target_column),
        ("检查source列", data_source_column),
        ("检查target列", data_target_column),
        ("开始行", start_row),
        ("大小写模式", "严格区分" if summary.case_sensitive else "忽略大小写"),
        ("匹配模式", "混合边界" if match_mode == "hybrid-boundary" else "纯包含"),
        ("总行数", summary.total_rows_checked),
        ("命中术语行数", summary.matched_rows),
        ("问题行数", summary.problem_rows),
        ("问题条数", summary.problem_count),
        ("术语表条数", summary.glossary_term_count),
        ("冲突术语数", summary.conflict_count),
    ]
    for row_index, (label, value) in enumerate(summary_rows, start=2):
        summary_sheet.cell(row_index, 1, label)
        summary_sheet.cell(row_index, 2, value)

    summary_sheet["D1"] = "冲突source术语"
    summary_sheet["E1"] = "候选target术语"
    if conflicts:
        for row_index, conflict in enumerate(conflicts, start=2):
            summary_sheet.cell(row_index, 4, conflict.source_term)
            summary_sheet.cell(row_index, 5, " / ".join(conflict.target_terms))
    else:
        summary_sheet["D2"] = "无"
