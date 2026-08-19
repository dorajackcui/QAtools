#!/usr/bin/env python3
"""Shared helpers for reading Excel workbook metadata in GUI flows."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree
from zipfile import BadZipFile, ZipFile

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException


@dataclass(frozen=True)
class WorkbookSheetChoices:
    sheet_names: tuple[str, ...]
    default_sheet: str | None


@dataclass(frozen=True)
class SourceTargetColumns:
    detected_source_column: str | None
    detected_target_column: str | None


def resolve_workbook_path(input_file: str | Path) -> Path:
    workbook_path = Path(input_file).expanduser().resolve()
    if not workbook_path.exists():
        raise FileNotFoundError(f"输入文件不存在: {workbook_path}")
    return workbook_path


def normalize_header(value: object) -> str:
    return "" if value is None else str(value).strip().casefold()


def list_workbook_sheets(input_file: str | Path) -> WorkbookSheetChoices:
    workbook_path = resolve_workbook_path(input_file)
    try:
        with ZipFile(workbook_path) as archive:
            workbook_xml = archive.read("xl/workbook.xml")
    except (BadZipFile, KeyError) as exc:
        raise InvalidFileException(
            f"无法读取 Excel 工作簿结构: {workbook_path}"
        ) from exc

    namespace = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    root = ElementTree.fromstring(workbook_xml)
    sheet_nodes = root.findall(f"{namespace}sheets/{namespace}sheet")
    sheet_names = tuple(node.attrib.get("name", "") for node in sheet_nodes)
    visible_indexes = [
        index
        for index, node in enumerate(sheet_nodes)
        if node.attrib.get("state", "visible") == "visible"
    ]
    view = root.find(f"{namespace}bookViews/{namespace}workbookView")
    try:
        active_index = int(view.attrib.get("activeTab", "0")) if view is not None else 0
    except ValueError:
        active_index = 0
    if active_index not in visible_indexes:
        active_index = visible_indexes[0] if visible_indexes else 0
    default_sheet = (
        sheet_names[active_index]
        if sheet_names and 0 <= active_index < len(sheet_names)
        else None
    )
    return WorkbookSheetChoices(
        sheet_names=sheet_names,
        default_sheet=default_sheet,
    )


def detect_source_target_columns(
    input_file: str | Path,
    sheet: str | None = None,
) -> SourceTargetColumns:
    workbook_path = resolve_workbook_path(input_file)
    workbook = load_workbook(workbook_path, read_only=True)

    try:
        if sheet and sheet not in workbook.sheetnames:
            raise ValueError(f"工作表不存在: {sheet}")

        worksheet = workbook[sheet] if sheet else workbook.active
        header_row = next(
            worksheet.iter_rows(min_row=1, max_row=1, values_only=True),
            (),
        )

        source_columns: list[str] = []
        target_columns: list[str] = []
        for column_index, value in enumerate(header_row, start=1):
            normalized_value = normalize_header(value)
            if normalized_value == "source":
                source_columns.append(get_column_letter(column_index))
            elif normalized_value == "target":
                target_columns.append(get_column_letter(column_index))

        return SourceTargetColumns(
            detected_source_column=source_columns[0] if len(source_columns) == 1 else None,
            detected_target_column=target_columns[0] if len(target_columns) == 1 else None,
        )
    finally:
        workbook.close()
