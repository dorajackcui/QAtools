from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class TemplateMatch:
    text: str
    template: str
    values: dict[str, str]


@dataclass(frozen=True)
class RowItem:
    row_number: int
    source: str
    match: TemplateMatch
    original_values: tuple[object, ...]
    raw_source: str
    raw_segment: str
    raw_existing_target: str
    segment_index: int
    segment_count: int
    segment_prefix: str
    segment_suffix: str

__all__ = [
    "RowItem",
    "TemplateMatch",
]
