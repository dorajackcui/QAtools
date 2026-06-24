#!/usr/bin/env python3
"""Shared helpers for generated Excel output files."""

from __future__ import annotations

from collections.abc import Mapping
from pathlib import Path

from openpyxl.utils import column_index_from_string


ROW_PROBLEM_COLUMN_HEADER = "术语QA问题"
ROW_PROBLEM_SEPARATOR = "；"


def build_prefixed_output_path(input_file: str | Path, prefix: str) -> Path:
    input_path = Path(input_file).expanduser()
    return input_path.with_name(f"{prefix}{input_path.name}")


def format_row_problem_text(source_term: str, expected_target_term: str, description: str) -> str:
    source = source_term.strip()
    expected_target = expected_target_term.strip()
    description = description.strip()

    if source and expected_target:
        problem_subject = f"{source} -> {expected_target}"
    elif source:
        problem_subject = source
    elif expected_target:
        problem_subject = f"-> {expected_target}"
    else:
        return description

    if not description:
        return problem_subject
    return f"{problem_subject}：{description}"


def insert_row_problem_column(
    worksheet,
    target_column: str,
    row_problem_texts: Mapping[int, str],
) -> None:
    target_column_index = column_index_from_string(target_column.strip().upper())
    problem_column_index = target_column_index + 1

    worksheet.insert_cols(problem_column_index)
    worksheet.cell(1, problem_column_index, ROW_PROBLEM_COLUMN_HEADER)
    for row_index, problem_text in sorted(row_problem_texts.items()):
        if problem_text:
            worksheet.cell(row_index, problem_column_index, problem_text)
