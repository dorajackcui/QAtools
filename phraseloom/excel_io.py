from __future__ import annotations

import json
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .models import RowItem, TranslationUnit
from .template_engine import PLACEHOLDER_RE, parse_template


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


def _write_output_workbook(
    output_path: Path,
    input_path: Path,
    units: list[TranslationUnit],
    result_rows: list[tuple[RowItem, TranslationUnit | None, str | None]],
) -> None:
    wb = Workbook()
    summary = wb.active
    summary.title = "summary"

    template_units = [unit for unit in units if unit.unit_type == "template"]
    segment_units = [unit for unit in units if unit.unit_type == "segment"]
    total_source_segments = len(result_rows)
    prefilled_units = sum(1 for unit in units if unit.target_unit)
    new_units = sum(1 for unit in units if not unit.target_unit)
    filled_rows = sum(1 for _, _, auto in result_rows if auto)
    new_rows = sum(unit.coverage_count for unit in units if not unit.target_unit)
    summary_rows = [
        ("total_source_rows", total_source_segments),
        ("total_translation_units", len(units)),
        ("already_filled_units", prefilled_units),
        ("already_filled_source_rows", filled_rows),
        ("units_to_translate", new_units),
        ("source_rows_to_translate", new_rows),
    ]
    for row in summary_rows:
        summary.append(row)

    review = wb.create_sheet("translation_units")
    review.append(
        [
            "unit_id",
            "unit_type",
            "source_unit",
            "target_unit",
            "coverage_count",
            "unique_source_count",
            "variables",
            "status",
            "sample_sources",
            "row_numbers",
            "target_unit_source",
            "suggested_target_unit",
            "warning",
            "translator_note",
        ]
    )
    for unit in units:
        review.append(
            [
                unit.unit_id,
                unit.unit_type,
                unit.source_unit,
                unit.target_unit or None,
                unit.coverage_count,
                unit.unique_source_count,
                _variables_summary(unit),
                "ready" if unit.target_unit else None,
                "\n".join(dict.fromkeys(item.source for item in unit.items[:10])),
                ",".join(str(item.row_number) for item in unit.items),
                unit.target_unit_source or None,
                unit.suggested_target_unit or None,
                unit.warning or None,
                None,
            ]
        )

    todo = wb.create_sheet("to_translate")
    todo.append(
        [
            "unit_id",
            "unit_type",
            "source_unit",
            "target_unit",
            "coverage_count",
            "sample_sources",
            "variables",
            "warning",
            "translator_note",
        ]
    )
    for unit in units:
        if unit.target_unit:
            continue
        todo.append(
            [
                unit.unit_id,
                unit.unit_type,
                unit.source_unit,
                None,
                unit.coverage_count,
                "\n".join(dict.fromkeys(item.source for item in unit.items[:5])),
                _variables_summary(unit),
                unit.warning or None,
                None,
            ]
        )

    result = wb.create_sheet("source_map")
    result.append(
        [
            "row_number",
            "source",
            "unit_id",
            "unit_type",
            "source_unit",
            "variable_values",
            "target_unit",
            "auto_target",
            "existing_target",
            "fill_status",
            "warning",
        ]
    )
    for row, unit, auto_target in result_rows:
        if unit is None:
            warning = "no translation unit"
            fill_status = "unit_not_found"
        elif not unit.target_unit:
            warning = "fill target_unit in translation_units, then rerun fill"
            fill_status = "missing_target_unit"
        else:
            warning = unit.warning
            fill_status = "filled"

        result.append(
            [
                row.row_number,
                row.source,
                unit.unit_id if unit else None,
                unit.unit_type if unit else None,
                unit.source_unit if unit else None,
                json.dumps(
                    row.match.values if unit and unit.unit_type == "template" else {},
                    ensure_ascii=False,
                ),
                unit.target_unit if unit else None,
                auto_target,
                row.existing_target or None,
                fill_status,
                warning or None,
            ]
        )

    filled = wb.create_sheet("filled_workbook")
    headers = _read_headers(input_path)
    filled.append(
        headers
        + [
            "auto_target",
            "fill_status",
            "unit_id",
            "unit_type",
            "warning",
        ]
    )
    for row, unit, auto_target in result_rows:
        if unit is None:
            warning = "no translation unit"
            fill_status = "unit_not_found"
        elif not unit.target_unit:
            warning = "fill target_unit in to_translate, then rerun fill"
            fill_status = "missing_target_unit"
        else:
            warning = unit.warning
            fill_status = "filled"
        original = list(row.original_values)
        if len(original) < len(headers):
            original.extend([None] * (len(headers) - len(original)))
        filled.append(
            original[: len(headers)]
            + [
                auto_target,
                fill_status,
                unit.unit_id if unit else None,
                unit.unit_type if unit else None,
                warning or None,
            ]
        )

    qa = wb.create_sheet("qa_report")
    qa.append(["check", "count"])
    qa.append(
        [
            "missing_target_unit",
            sum(1 for _, unit, auto in result_rows if unit and not unit.target_unit),
        ]
    )
    qa.append(["filled", sum(1 for _, _, auto in result_rows if auto)])
    qa.append(["warning_units", sum(1 for unit in units if unit.warning)])
    qa.append(["template_units", len(template_units)])
    qa.append(["segment_units", len(segment_units)])

    for ws in wb.worksheets:
        _style_sheet(ws)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def _write_to_translate_workbook(
    output_path: Path, input_path: Path, units: list[TranslationUnit]
) -> None:
    wb = Workbook()
    todo = wb.active
    todo.title = "to_translate"
    todo.append(
        [
            "unit_id",
            "unit_type",
            "source_unit",
            "target_unit",
            "coverage_count",
            "sample_sources",
            "variables",
            "warning",
            "translator_note",
        ]
    )
    for unit in units:
        if unit.target_unit:
            continue
        todo.append(
            [
                unit.unit_id,
                unit.unit_type,
                unit.source_unit,
                None,
                unit.coverage_count,
                "\n".join(dict.fromkeys(item.source for item in unit.items[:5])),
                _variables_summary(unit),
                unit.warning or None,
                None,
            ]
        )

    prefilled = wb.create_sheet("prefilled_units")
    prefilled.append(
        [
            "unit_id",
            "unit_type",
            "source_unit",
            "target_unit",
            "coverage_count",
            "target_unit_source",
        ]
    )
    for unit in units:
        if not unit.target_unit:
            continue
        prefilled.append(
            [
                unit.unit_id,
                unit.unit_type,
                unit.source_unit,
                unit.target_unit,
                unit.coverage_count,
                unit.target_unit_source or None,
            ]
        )
    prefilled.sheet_state = "hidden"

    for ws in wb.worksheets:
        _style_sheet(ws)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def _write_target_column_workbook(
    output_path: Path,
    input_path: Path,
    target_col: str | int,
    result_rows: list[tuple[RowItem, TranslationUnit | None, str | None]],
) -> None:
    wb = load_workbook(input_path)
    ws = wb.worksheets[0]
    target_index = _resolve_or_create_column(ws, target_col)

    for row, _, auto_target in result_rows:
        if auto_target:
            ws.cell(row=row.row_number, column=target_index).value = auto_target

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def _resolve_or_create_column(ws, col: str | int) -> int:
    try:
        return _resolve_column(ws, col)
    except ValueError:
        if isinstance(col, int) or str(col).isdigit():
            raise
        column = ws.max_column + 1
        ws.cell(row=1, column=column).value = str(col)
        return column


def _write_tm_workbook(
    output_path: Path,
    input_path: Path,
    units: list[TranslationUnit],
    rows: list[RowItem],
) -> None:
    wb = Workbook()
    summary = wb.active
    summary.title = "summary"

    template_units = [unit for unit in units if unit.unit_type == "template"]
    segment_units = [unit for unit in units if unit.unit_type == "segment"]
    unique_source_segments = len({row.source for row in rows})
    summary_rows = [
        ("input_file", str(input_path)),
        ("tm_source_segments", len(rows)),
        ("unique_source_segments", unique_source_segments),
        ("duplicate_source_segments", len(rows) - unique_source_segments),
        ("tm_pair_count", len(units)),
        ("template_pair_count", len(template_units)),
        ("segment_pair_count", len(segment_units)),
        ("matched_source_segments", sum(unit.coverage_count for unit in units)),
        ("next_step", "Use this workbook with extract --tm to prefill new translation_units."),
    ]
    for row in summary_rows:
        summary.append(row)

    pairs = wb.create_sheet("tm_pairs")
    pairs.append(
        [
            "tm_id",
            "unit_type",
            "source_unit",
            "target_unit",
            "coverage_count",
            "unique_source_count",
            "variables",
            "sample_sources",
            "sample_targets",
            "row_numbers",
            "warning",
        ]
    )
    for index, unit in enumerate(units, start=1):
        pairs.append(
            [
                f"TM{index:05d}",
                unit.unit_type,
                unit.source_unit,
                unit.target_unit or None,
                unit.coverage_count,
                unit.unique_source_count,
                _variables_summary(unit),
                "\n".join(dict.fromkeys(item.source for item in unit.items[:10])),
                "\n".join(
                    dict.fromkeys(
                        item.existing_target for item in unit.items[:10] if item.existing_target
                    )
                ),
                ",".join(str(item.row_number) for item in unit.items),
                unit.warning or None,
            ]
        )

    tm_map = wb.create_sheet("tm_map")
    tm_map.append(
        [
            "row_number",
            "source",
            "target",
            "unit_type",
            "source_unit",
            "target_unit",
            "variable_values",
        ]
    )
    unit_by_row_number = {
        item.row_number: unit for unit in units for item in unit.items
    }
    for row in rows:
        unit = unit_by_row_number.get(row.row_number)
        tm_map.append(
            [
                row.row_number,
                row.source,
                row.existing_target,
                unit.unit_type if unit else None,
                unit.source_unit if unit else None,
                unit.target_unit if unit else None,
                json.dumps(
                    row.match.values if unit and unit.unit_type == "template" else {},
                    ensure_ascii=False,
                ),
            ]
        )

    qa = wb.create_sheet("qa_report")
    qa.append(["check", "count"])
    qa.append(["tm_pairs", len(units)])
    qa.append(["template_pairs", len(template_units)])
    qa.append(["segment_pairs", len(segment_units)])
    qa.append(["warning_pairs", sum(1 for unit in units if unit.warning)])

    for ws in wb.worksheets:
        _style_sheet(ws)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def _variables_summary(unit: TranslationUnit) -> str | None:
    if unit.unit_type != "template":
        return None
    placeholders = PLACEHOLDER_RE.findall(unit.source_unit)
    if not placeholders:
        return None
    samples: dict[str, list[str]] = {placeholder: [] for placeholder in placeholders}
    for item in unit.items:
        for key, value in item.match.values.items():
            placeholder = "{" + key + "}"
            if placeholder in samples and value not in samples[placeholder]:
                samples[placeholder].append(value)
    return "; ".join(
        f"{placeholder}={','.join(values[:5])}" for placeholder, values in samples.items()
    )


def _style_sheet(ws) -> None:
    header_fill = PatternFill("solid", fgColor="1F2937")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(vertical="top", wrap_text=True)

    ws.freeze_panes = "A2"
    if ws.max_row and ws.max_column:
        ws.auto_filter.ref = ws.dimensions

    for column in ws.columns:
        letter = get_column_letter(column[0].column)
        max_len = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column)
        ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 60)
        for cell in column[1:]:
            cell.alignment = Alignment(vertical="top", wrap_text=True)


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
