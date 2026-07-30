from __future__ import annotations

from pathlib import Path

from . import workbook_schema as schema
from ._entity_extract import _build_entity_split, _build_entity_tm
from ._entity_fill import (
    fill_entity_pack_workbook,
    fill_entity_workbook,
    merge_entity_pack_workbook,
    merge_entity_workbooks,
)
from ._entity_io import (
    _build_entity_pack_workbook,
    _read_tm_pair_rows,
    _read_unit_rows,
    _save_workbook,
    _write_entity_related_workbook,
    _write_entity_tm_workbook,
    _write_non_entity_workbook,
)
from ._entity_prefill import _prefill_entity_sheets, prefill_entity_workbook
from ._entity_types import ClusterProbeStrategy, EntityExtractionStrategy


def _default_entity_work_dir(input_path: Path) -> Path:
    return input_path.parent / f"{input_path.stem}_l10n"


def default_entity_memory_output_path(input_path: str | Path) -> Path:
    input_path = Path(input_path)
    return _default_entity_work_dir(input_path) / f"{input_path.stem}_entity_memory.xlsx"


def default_entity_pack_output_path(input_path: str | Path) -> Path:
    input_path = Path(input_path)
    return _default_entity_work_dir(input_path) / f"{input_path.stem}_entity_pack.xlsx"


def default_entity_filled_pack_output_path(input_path: str | Path) -> Path:
    input_path = Path(input_path)
    return input_path.with_name(f"{input_path.stem}_filled.xlsx")


def default_entity_merged_todo_output_path(input_path: str | Path) -> Path:
    input_path = Path(input_path)
    return input_path.with_name(f"{input_path.stem}_merged_todo.xlsx")


def extract_entity_memory_workbook(
    input_path: str | Path,
    output_path: str | Path,
    *,
    min_group_size: int = 3,
    strategy: EntityExtractionStrategy | None = None,
) -> dict[str, int | str]:
    return extract_entity_tm_workbook(
        input_path,
        output_path,
        min_group_size=min_group_size,
        strategy=strategy,
    )


def split_entity_workbook(
    input_path: str | Path,
    entity_output_path: str | Path,
    non_entity_output_path: str | Path,
    *,
    min_group_size: int = 3,
    strategy: EntityExtractionStrategy | None = None,
) -> dict[str, int | str]:
    input_path = Path(input_path)
    entity_output_path = Path(entity_output_path)
    non_entity_output_path = Path(non_entity_output_path)
    rows, headers = _read_unit_rows(input_path, schema.TO_TRANSLATE_SHEET)
    active_strategy = strategy or ClusterProbeStrategy(min_group_size=min_group_size)
    clusters = active_strategy.find_clusters(
        [(row.source_unit, row.target_unit) for row in rows]
    )
    entity_rows, structures, terms, source_map = _build_entity_split(rows, clusters)
    entity_indices = {row.original_index for row in entity_rows}
    non_entity_rows = [row for row in rows if row.original_index not in entity_indices]

    _write_entity_related_workbook(
        entity_output_path,
        input_path,
        headers,
        entity_rows,
        structures,
        terms,
        source_map,
    )
    _write_non_entity_workbook(non_entity_output_path, input_path, headers, non_entity_rows)
    return {
        "entity_unit_count": len(entity_rows),
        "non_entity_unit_count": len(non_entity_rows),
        "entity_structure_count": len(structures),
        "entity_term_count": len(terms),
        "entity_output_path": str(entity_output_path),
        "non_entity_output_path": str(non_entity_output_path),
    }


def prepare_entity_pack_workbook(
    input_path: str | Path,
    output_path: str | Path,
    *,
    tm_path: str | Path | None = None,
    min_group_size: int = 3,
    strategy: EntityExtractionStrategy | None = None,
) -> dict[str, int | str]:
    input_path = Path(input_path)
    output_path = Path(output_path)
    rows, headers = _read_unit_rows(input_path, schema.TO_TRANSLATE_SHEET)
    active_strategy = strategy or ClusterProbeStrategy(min_group_size=min_group_size)
    clusters = active_strategy.find_clusters(
        [(row.source_unit, row.target_unit) for row in rows]
    )
    entity_rows, structures, terms, source_map = _build_entity_split(rows, clusters)
    entity_indices = {row.original_index for row in entity_rows}
    non_entity_rows = [row for row in rows if row.original_index not in entity_indices]
    wb = _build_entity_pack_workbook(
        input_path,
        headers,
        entity_rows,
        non_entity_rows,
        structures,
        terms,
        source_map,
    )
    prefilled_structure_count = 0
    prefilled_term_count = 0
    if tm_path is not None:
        prefilled_structure_count, prefilled_term_count = _prefill_entity_sheets(
            wb,
            Path(tm_path),
        )
    try:
        _save_workbook(wb, output_path)
    finally:
        wb.close()
    return {
        "related_unit_count": len(entity_rows),
        "non_related_unit_count": len(non_entity_rows),
        "entity_structure_count": len(structures),
        "entity_term_count": len(terms),
        "prefilled_structure_count": prefilled_structure_count,
        "prefilled_term_count": prefilled_term_count,
        "output_path": str(output_path),
    }


def extract_entity_tm_workbook(
    input_path: str | Path,
    output_path: str | Path,
    *,
    min_group_size: int = 3,
    strategy: EntityExtractionStrategy | None = None,
) -> dict[str, int | str]:
    input_path = Path(input_path)
    output_path = Path(output_path)
    rows = _read_tm_pair_rows(input_path)
    active_strategy = strategy or ClusterProbeStrategy(min_group_size=min_group_size)
    clusters = active_strategy.find_clusters(
        [(row.source_unit, row.target_unit) for row in rows]
    )
    structures, terms = _build_entity_tm(rows, clusters, min_group_size)
    _write_entity_tm_workbook(output_path, structures, terms)
    return {
        "entity_structure_count": len(structures),
        "entity_term_count": len(terms),
        "output_path": str(output_path),
    }


__all__ = [
    "ClusterProbeStrategy",
    "EntityExtractionStrategy",
    "_default_entity_work_dir",
    "default_entity_filled_pack_output_path",
    "default_entity_memory_output_path",
    "default_entity_merged_todo_output_path",
    "default_entity_pack_output_path",
    "extract_entity_memory_workbook",
    "extract_entity_tm_workbook",
    "fill_entity_pack_workbook",
    "fill_entity_workbook",
    "merge_entity_pack_workbook",
    "merge_entity_workbooks",
    "prepare_entity_pack_workbook",
    "prefill_entity_workbook",
    "split_entity_workbook",
]
