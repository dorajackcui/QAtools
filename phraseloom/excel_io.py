from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook

from .models import RowItem
from .template_engine import parse_template


def _read_source_rows(
    input_path: Path, source_col: str | int, target_col: str | int | None
) -> list[RowItem]:
    wb = load_workbook(input_path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    source_index = _resolve_column(ws, source_col)
    target_index = _resolve_column(ws, target_col) if target_col is not None else None

    rows: list[RowItem] = []
    seen_source = False
    blank_source_run = 0
    for row_number, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        source_value = _cell_value(row, source_index)
        if source_value is None or str(source_value).strip() == "":
            if seen_source:
                blank_source_run += 1
                if blank_source_run >= 1000:
                    break
            continue
        seen_source = True
        blank_source_run = 0
        source = str(source_value).strip()
        target_value = _cell_value(row, target_index) if target_index else ""
        existing_target = "" if target_value is None else str(target_value).strip()
        rows.append(
            RowItem(row_number, source, existing_target, parse_template(source), tuple(row))
        )

    return rows


def _load_translated_units(path: Path) -> dict[tuple[str, str], str]:
    wb = load_workbook(path, read_only=True, data_only=True)
    if "tm_pairs" in wb.sheetnames:
        ws = wb["tm_pairs"]
        headers = [
            str(cell.value).strip() if cell.value is not None else "" for cell in ws[1]
        ]
        try:
            type_index = headers.index("unit_type")
            source_index = headers.index("source_unit")
            target_index = headers.index("target_unit")
        except ValueError:
            return {}

        units: dict[tuple[str, str], str] = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            unit_type = row[type_index] if type_index < len(row) else None
            source_unit = row[source_index] if source_index < len(row) else None
            target_unit = row[target_index] if target_index < len(row) else None
            if unit_type and source_unit and target_unit:
                units[(str(unit_type), str(source_unit))] = str(target_unit)
        return units

    if "translation_units" in wb.sheetnames:
        ws = wb["translation_units"]
        headers = [
            str(cell.value).strip() if cell.value is not None else "" for cell in ws[1]
        ]
        try:
            type_index = headers.index("unit_type")
            source_index = headers.index("source_unit")
            target_index = headers.index("target_unit")
        except ValueError:
            return {}

        units: dict[tuple[str, str], str] = {}
        status_index = headers.index("status") if "status" in headers else None
        for row in ws.iter_rows(min_row=2, values_only=True):
            unit_type = row[type_index] if type_index < len(row) else None
            source_unit = row[source_index] if source_index < len(row) else None
            target_unit = row[target_index] if target_index < len(row) else None
            status = (
                row[status_index]
                if status_index is not None and status_index < len(row)
                else None
            )
            if (
                unit_type
                and source_unit
                and target_unit
                and str(status or "").strip().lower() != "skip"
            ):
                units[(str(unit_type), str(source_unit))] = str(target_unit)
        return units

    if "to_translate" in wb.sheetnames:
        units = _load_unit_sheet(wb["to_translate"])
        if "prefilled_units" in wb.sheetnames:
            units.update(_load_unit_sheet(wb["prefilled_units"]))
        return units

    if "template_review" not in wb.sheetnames:
        return {}
    ws = wb["template_review"]
    headers = [str(cell.value).strip() if cell.value is not None else "" for cell in ws[1]]
    try:
        source_index = headers.index("source_template")
        target_index = headers.index("target_template")
    except ValueError:
        return {}

    templates: dict[tuple[str, str], str] = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        source_template = row[source_index] if source_index < len(row) else None
        target_template = row[target_index] if target_index < len(row) else None
        if source_template and target_template:
            templates[("template", str(source_template))] = str(target_template)
    return templates


def _load_unit_sheet(ws) -> dict[tuple[str, str], str]:
    headers = [
        str(cell.value).strip() if cell.value is not None else "" for cell in ws[1]
    ]
    try:
        type_index = headers.index("unit_type")
        source_index = headers.index("source_unit")
        target_index = headers.index("target_unit")
    except ValueError:
        return {}

    units: dict[tuple[str, str], str] = {}
    status_index = headers.index("status") if "status" in headers else None
    for row in ws.iter_rows(min_row=2, values_only=True):
        unit_type = row[type_index] if type_index < len(row) else None
        source_unit = row[source_index] if source_index < len(row) else None
        target_unit = row[target_index] if target_index < len(row) else None
        status = (
            row[status_index]
            if status_index is not None and status_index < len(row)
            else None
        )
        if (
            unit_type
            and source_unit
            and target_unit
            and str(status or "").strip().lower() != "skip"
        ):
            units[(str(unit_type), str(source_unit))] = str(target_unit)
    return units


def _resolve_column(ws, col: str | int | None) -> int:
    if col is None:
        raise ValueError("Column cannot be None")
    if isinstance(col, int):
        return col
    if str(col).isdigit():
        return int(str(col))

    wanted = str(col).strip()
    wanted_lower = wanted.lower()
    for cell in ws[1]:
        value = "" if cell.value is None else str(cell.value).strip()
        if value == wanted or value.lower() == wanted_lower:
            return cell.column
    raise ValueError(f"Column {col!r} not found in header row")


def _cell_value(row: tuple[object, ...], one_based_index: int | None) -> object | None:
    if one_based_index is None:
        return None
    zero_based = one_based_index - 1
    return row[zero_based] if zero_based < len(row) else None


def _read_headers(input_path: Path) -> list[str]:
    wb = load_workbook(input_path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    return [
        str(cell.value).strip() if cell.value is not None else f"column_{cell.column}"
        for cell in ws[1]
    ]


def _default_work_dir(input_path: Path) -> Path:
    return input_path.parent / f"{input_path.stem}_l10n"


def _default_extract_output_path(input_path: Path) -> Path:
    return _default_work_dir(input_path) / f"{input_path.stem}_template_pack.xlsx"


def _default_tm_output_path(input_path: Path) -> Path:
    return _default_work_dir(input_path) / f"{input_path.stem}_tm_pairs.xlsx"


def _default_to_translate_output_path(input_path: Path) -> Path:
    return _default_work_dir(input_path) / f"{input_path.stem}_to_translate.xlsx"


def _default_fill_output_path(input_path: Path) -> Path:
    return _default_work_dir(input_path) / f"{input_path.stem}_filled.xlsx"


def _default_legacy_output_path(input_path: Path) -> Path:
    return _default_work_dir(input_path) / f"{input_path.stem}_template_demo.xlsx"


def _write_output_workbook(*args, **kwargs) -> None:
    from ._template_workflow import _write_output_workbook as impl

    return impl(*args, **kwargs)


def _write_to_translate_workbook(*args, **kwargs) -> None:
    from ._template_workflow import _write_to_translate_workbook as impl

    return impl(*args, **kwargs)


def _write_target_column_workbook(*args, **kwargs) -> None:
    from ._template_workflow import _write_target_column_workbook as impl

    return impl(*args, **kwargs)


def _write_tm_workbook(*args, **kwargs) -> None:
    from ._template_workflow import _write_tm_workbook as impl

    return impl(*args, **kwargs)


__all__ = [
    "_cell_value",
    "_default_extract_output_path",
    "_default_fill_output_path",
    "_default_legacy_output_path",
    "_default_tm_output_path",
    "_default_to_translate_output_path",
    "_default_work_dir",
    "_load_translated_units",
    "_load_unit_sheet",
    "_read_headers",
    "_read_source_rows",
    "_resolve_column",
    "_write_output_workbook",
    "_write_target_column_workbook",
    "_write_tm_workbook",
    "_write_to_translate_workbook",
]
