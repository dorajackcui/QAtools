#!/usr/bin/env python3
"""Shared helpers for generated Excel output files."""

from __future__ import annotations

from collections.abc import Mapping
from pathlib import Path

from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import column_index_from_string, get_column_letter


ROW_PROBLEM_COLUMN_HEADER = "术语QA问题"
ROW_PROBLEM_SEPARATOR = "；"
PROBLEM_BASE_HEADERS = ("行号", "source原文", "target原文", "问题描述")
HEADER_FILL = PatternFill(fill_type="solid", fgColor="D9EAF7")


def rebuild_output_sheet(workbook, current_sheet_name: str, sheet_name: str):
    if current_sheet_name == sheet_name:
        raise ValueError(f"数据工作表名称不能为 {sheet_name}")
    if sheet_name in workbook.sheetnames:
        del workbook[sheet_name]
    return workbook.create_sheet(title=sheet_name)


def join_unique_text(values, separator: str = ROW_PROBLEM_SEPARATOR) -> str:
    """Join non-empty text values once, preserving their first-seen order."""
    unique_values: list[str] = []
    seen: set[str] = set()
    for value in values:
        text = str(value).strip() if value is not None else ""
        if text and text not in seen:
            seen.add(text)
            unique_values.append(text)
    return separator.join(unique_values)


def write_output_table(
    workbook,
    *,
    current_sheet_name: str,
    sheet_name: str,
    headers,
    rows,
):
    """Write a consistently formatted output table and return its worksheet."""
    worksheet = rebuild_output_sheet(workbook, current_sheet_name, sheet_name)
    normalized_headers = tuple(headers)
    for column_index, header in enumerate(normalized_headers, start=1):
        cell = worksheet.cell(1, column_index, header)
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(vertical="center", wrap_text=True)

    max_lengths = [len(str(header)) for header in normalized_headers]
    for row_index, row in enumerate(rows, start=2):
        for column_index, value in enumerate(row, start=1):
            cell = worksheet.cell(row_index, column_index, value)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if value is not None:
                max_lengths[column_index - 1] = max(
                    max_lengths[column_index - 1],
                    max(len(line) for line in str(value).splitlines() or [""]),
                )

    preferred_widths = {1: 10, 2: 52, 3: 52, 4: 48}
    for column_index, max_length in enumerate(max_lengths, start=1):
        width = preferred_widths.get(column_index, min(max(max_length + 2, 12), 30))
        worksheet.column_dimensions[get_column_letter(column_index)].width = width

    worksheet.freeze_panes = "A2"
    if normalized_headers:
        last_column = get_column_letter(len(normalized_headers))
        worksheet.auto_filter.ref = f"A1:{last_column}{max(worksheet.max_row, 1)}"
    return worksheet


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
