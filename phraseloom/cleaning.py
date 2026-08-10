from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from typing import Iterable

from .models import RowItem
from .template_engine import is_candidate_template


@dataclass(frozen=True)
class CleanStringUnit:
    unit_id: str
    unit_type: str
    source: str
    items: tuple[RowItem, ...]


def build_clean_string_units(
    rows: list[RowItem],
    *,
    min_template_variants: int = 2,
) -> list[CleanStringUnit]:
    """Deduplicate rows and merge eligible numeric variants into templates."""
    template_groups = _eligible_template_groups(rows, min_template_variants)
    assigned_segments = {
        (item.row_number, item.segment_index)
        for items in template_groups.values()
        for item in items
    }

    units = [
        CleanStringUnit(
            unit_id=f"T{index:04d}",
            unit_type="template",
            source=source,
            items=tuple(items),
        )
        for index, (source, items) in enumerate(
            sorted(template_groups.items(), key=lambda entry: (-len(entry[1]), entry[0])),
            start=1,
        )
    ]

    segment_groups: dict[str, list[RowItem]] = defaultdict(list)
    for row in rows:
        if (row.row_number, row.segment_index) not in assigned_segments:
            segment_groups[row.source].append(row)
    units.extend(
        CleanStringUnit(
            unit_id=f"S{index:04d}",
            unit_type="segment",
            source=source,
            items=tuple(items),
        )
        for index, (source, items) in enumerate(
            sorted(segment_groups.items(), key=lambda entry: (-len(entry[1]), entry[0])),
            start=1,
        )
    )
    return units


def _eligible_template_groups(
    rows: Iterable[RowItem],
    min_template_variants: int,
) -> dict[str, list[RowItem]]:
    groups: dict[str, list[RowItem]] = defaultdict(list)
    for row in rows:
        if is_candidate_template(row.match):
            groups[row.match.template].append(row)
    return {
        template: items
        for template, items in groups.items()
        if len({item.source for item in items}) >= min_template_variants
    }


__all__ = ["CleanStringUnit", "build_clean_string_units"]
