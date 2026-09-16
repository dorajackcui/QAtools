#!/usr/bin/env python3
"""Find identical target text associated with different source text."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from openpyxl.utils import column_index_from_string

from tools.consistency_text import normalize_consistency_text
from tools.report_text import EXCEL_CELL_TEXT_LIMIT, format_consistency_variants
from tools.excel_output import (
    PROBLEM_BASE_HEADERS,
    build_prefixed_output_path,
    existing_cell_value,
    value_row_numbers,
    load_workbook_for_editing,
    validate_distinct_source_target_columns,
    validate_report_output_path,
    write_output_table,
)


PROBLEM_SHEET_NAME = "同Target不同Source"


@dataclass(frozen=True)
class TargetOccurrence:
    row_index: int
    source_text: str
    target_text: str = ""


@dataclass(frozen=True)
class CheckSummary:
    output_path: Path
    worksheet_title: str
    source_column: str
    target_column: str
    start_row: int
    total_rows_checked: int
    non_empty_target_rows: int
    repeated_target_count: int
    inconsistent_target_count: int
    problem_rows: int


def normalize_column(column_name: str) -> str:
    normalized = column_name.strip().upper()
    column_index_from_string(normalized)
    return normalized


def cell_text(value: object) -> str:
    return "" if value is None else str(value)


def build_default_output_path(input_file: str | Path) -> Path:
    return build_prefixed_output_path(input_file, "target_consistency_check_")


def process_excel(
    input_file: str | Path,
    source_column: str,
    target_column: str,
    sheet: str | None = None,
    start_row: int = 2,
    output_file: str | Path | None = None,
) -> CheckSummary:
    """Report non-empty target strings that map to multiple exact source strings."""

    if start_row < 1:
        raise ValueError("开始行必须大于等于 1。")

    input_path = Path(input_file).expanduser().resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"输入文件不存在: {input_path}")
    source_column = normalize_column(source_column)
    target_column = normalize_column(target_column)
    validate_distinct_source_target_columns(source_column, target_column)
    output_path = (
        Path(output_file).expanduser().resolve()
        if output_file
        else build_default_output_path(input_path)
    )

    validate_report_output_path(input_path, output_path)
    workbook = load_workbook_for_editing(input_path)
    try:
        summary = process_workbook(
            workbook=workbook,
            output_path=output_path,
            source_column=source_column,
            target_column=target_column,
            sheet=sheet,
            start_row=start_row,
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
    format_output: bool = True,
    include_grouped_rows: bool = True,
) -> CheckSummary:
    """Run the check against an already-open workbook without saving it."""
    if start_row < 1:
        raise ValueError("开始行必须大于等于 1。")

    source_column = normalize_column(source_column)
    target_column = normalize_column(target_column)
    validate_distinct_source_target_columns(source_column, target_column)
    worksheet = workbook[sheet] if sheet else workbook.active
    occurrences_by_target: dict[str, list[TargetOccurrence]] = {}
    row_numbers = value_row_numbers(worksheet, (source_column, target_column), start_row=start_row)
    last_row = row_numbers[-1] if row_numbers else start_row - 1
    source_index = column_index_from_string(source_column)
    target_index = column_index_from_string(target_column)

    for row_index in row_numbers:
        target_text = cell_text(existing_cell_value(worksheet, row_index, target_index))
        target_key = normalize_consistency_text(target_text)
        if not target_key:
            continue
        source_text = cell_text(existing_cell_value(worksheet, row_index, source_index))
        occurrences_by_target.setdefault(target_key, []).append(
            TargetOccurrence(row_index=row_index, source_text=source_text, target_text=target_text)
        )

    repeated_target_count = 0
    inconsistent_target_count = 0
    problem_entries: list[tuple[int, str, str, str, int, str]] = []
    for occurrences in occurrences_by_target.values():
        if len(occurrences) < 2:
            continue
        repeated_target_count += 1
        source_variants: dict[str, str] = {}
        for occurrence in occurrences:
            key = normalize_consistency_text(occurrence.source_text)
            source_variants.setdefault(key, occurrence.source_text if key else "[空原文]")
        if len(source_variants) < 2:
            continue

        inconsistent_target_count += 1
        grouped_rows = (
            "、".join(str(occurrence.row_index) for occurrence in occurrences)[:EXCEL_CELL_TEXT_LIMIT]
            if include_grouped_rows else ""
        )
        description = format_consistency_variants(list(source_variants.values()), "原文")
        for occurrence in occurrences:
            problem_entries.append(
                (
                    occurrence.row_index,
                    occurrence.source_text,
                    occurrence.target_text,
                    description,
                    len(source_variants),
                    grouped_rows,
                )
            )

    write_output_table(
        workbook,
        current_sheet_name=worksheet.title,
        sheet_name=PROBLEM_SHEET_NAME,
        headers=PROBLEM_BASE_HEADERS + ("source版本数", "同组行号"),
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
        non_empty_target_rows=sum(
            len(items) for items in occurrences_by_target.values()
        ),
        repeated_target_count=repeated_target_count,
        inconsistent_target_count=inconsistent_target_count,
        problem_rows=len(problem_entries),
    )
