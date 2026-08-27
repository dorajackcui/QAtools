#!/usr/bin/env python3
"""Find identical target text associated with different source text."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from openpyxl.utils import column_index_from_string

from tools.excel_output import (
    PROBLEM_BASE_HEADERS,
    build_prefixed_output_path,
    find_last_value_row,
    load_workbook_for_editing,
    validate_distinct_source_target_columns,
    write_output_table,
)


PROBLEM_SHEET_NAME = "同Target不同Source"


@dataclass(frozen=True)
class TargetOccurrence:
    row_index: int
    source_text: str


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
) -> CheckSummary:
    """Run the check against an already-open workbook without saving it."""
    if start_row < 1:
        raise ValueError("开始行必须大于等于 1。")

    source_column = normalize_column(source_column)
    target_column = normalize_column(target_column)
    validate_distinct_source_target_columns(source_column, target_column)
    worksheet = workbook[sheet] if sheet else workbook.active
    occurrences_by_target: dict[str, list[TargetOccurrence]] = {}
    last_row = find_last_value_row(
        worksheet,
        (source_column, target_column),
        start_row=start_row,
    )

    for row_index in range(start_row, last_row + 1):
        target_text = cell_text(worksheet[f"{target_column}{row_index}"].value)
        if not target_text.strip():
            continue
        source_text = cell_text(worksheet[f"{source_column}{row_index}"].value)
        occurrences_by_target.setdefault(target_text, []).append(
            TargetOccurrence(row_index=row_index, source_text=source_text)
        )

    repeated_target_count = 0
    inconsistent_target_count = 0
    problem_entries: list[tuple[int, str, str, str, int, str]] = []
    for target_text, occurrences in occurrences_by_target.items():
        if len(occurrences) < 2:
            continue
        repeated_target_count += 1
        source_variants = {occurrence.source_text for occurrence in occurrences}
        if len(source_variants) < 2:
            continue

        inconsistent_target_count += 1
        grouped_rows = "、".join(str(occurrence.row_index) for occurrence in occurrences)
        for occurrence in occurrences:
            problem_entries.append(
                (
                    occurrence.row_index,
                    occurrence.source_text,
                    target_text,
                    f"同一 target 对应 {len(source_variants)} 个不同 source",
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
