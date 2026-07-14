#!/usr/bin/env python3
"""Run multiple Excel checks in sequence and keep results in one workbook."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from tools.excel_output import build_prefixed_output_path
from tools.tag_placeholder_checker.check_tags_and_placeholders import (
    process_excel as run_tag_check_excel,
)
from tools.term_pair_checker.extract_terms_from_excel import (
    process_excel as run_term_pair_check_excel,
)


@dataclass(frozen=True)
class WorkflowSummary:
    output_path: Path
    worksheet_title: str
    source_column: str
    target_column: str
    start_row: int
    ran_term_pair_check: bool
    ran_tag_check: bool
    term_count: int
    term_problem_count: int
    tag_problem_count: int


def build_default_output_path(input_file: str | Path) -> Path:
    return build_prefixed_output_path(input_file, "workflow_check_")


def run_workflow(
    *,
    input_file: str | Path,
    source_column: str,
    target_column: str,
    output_file: str | Path | None = None,
    sheet: str | None = None,
    start_row: int = 2,
    run_term_pair_check: bool = True,
    term_mark_styles: Iterable[str] | None = None,
    term_history_tb_file: str | Path | None = None,
    term_history_sheet: str | None = None,
    term_history_source_column: str | None = None,
    term_history_target_column: str | None = None,
    term_history_start_row: int = 2,
    run_tag_check: bool = True,
    tag_token_types: tuple[str, ...] | list[str] | None = None,
) -> WorkflowSummary:
    if not run_term_pair_check and not run_tag_check:
        raise ValueError("请至少选择一个 workflow 任务。")

    input_path = Path(input_file).expanduser().resolve()
    output_path = (
        Path(output_file).expanduser().resolve()
        if output_file
        else build_default_output_path(input_path)
    )

    current_input_path = input_path
    worksheet_title = sheet or ""
    normalized_source_column = source_column.strip().upper()
    normalized_target_column = target_column.strip().upper()
    term_count = 0
    term_problem_count = 0
    tag_problem_count = 0

    if run_term_pair_check:
        (
            worksheet_title,
            normalized_source_column,
            normalized_target_column,
            saved_path,
            term_count,
            term_problem_count,
        ) = run_term_pair_check_excel(
            input_file=current_input_path,
            source_column=source_column,
            target_column=target_column,
            sheet=sheet,
            start_row=start_row,
            mark_styles=term_mark_styles,
            history_tb_file=term_history_tb_file,
            history_sheet=term_history_sheet,
            history_source_column=term_history_source_column,
            history_target_column=term_history_target_column,
            history_start_row=term_history_start_row,
            output_file=output_path,
        )
        current_input_path = saved_path

    if run_tag_check:
        tag_summary = run_tag_check_excel(
            input_file=current_input_path,
            source_column=source_column,
            target_column=target_column,
            sheet=sheet,
            start_row=start_row,
            token_types=tag_token_types,
            output_file=output_path,
        )
        worksheet_title = tag_summary.worksheet_title
        tag_problem_count = tag_summary.problem_count

    return WorkflowSummary(
        output_path=output_path,
        worksheet_title=worksheet_title,
        source_column=normalized_source_column,
        target_column=normalized_target_column,
        start_row=start_row,
        ran_term_pair_check=run_term_pair_check,
        ran_tag_check=run_tag_check,
        term_count=term_count,
        term_problem_count=term_problem_count,
        tag_problem_count=tag_problem_count,
    )
