from __future__ import annotations

from dataclasses import dataclass

from .tag_engine import TagToken


@dataclass(frozen=True)
class TemplateMatch:
    text: str
    template: str
    values: dict[str, str]


@dataclass(frozen=True)
class RowItem:
    row_number: int
    source: str
    existing_target: str
    match: TemplateMatch
    original_values: tuple[object, ...]
    raw_source: str = ""
    raw_existing_target: str = ""
    tag_tokens: tuple[TagToken, ...] = ()
    tag_warnings: tuple[str, ...] = ()
    target_tag_warnings: tuple[str, ...] = ()


@dataclass(frozen=True)
class TranslationUnit:
    unit_id: str
    unit_type: str
    source_unit: str
    coverage_count: int
    unique_source_count: int
    items: tuple[RowItem, ...]
    target_unit: str
    target_unit_source: str
    suggested_target_unit: str
    warning: str


@dataclass(frozen=True)
class RowFillResult:
    row: RowItem
    unit: TranslationUnit | None
    auto_target: str | None
    warning: str = ""


@dataclass(frozen=True)
class EntityOccurrence:
    row_number: int
    source: str
    target: str
    entity: str
    entity_token_count: int


@dataclass(frozen=True)
class EntityCluster:
    source_pattern: str
    coverage_count: int
    unique_source_count: int
    unique_entity_count: int
    entity_values: tuple[str, ...]
    confidence: float
    risk: str
    sample_sources: tuple[str, ...]
    sample_targets: tuple[str, ...]
    row_numbers: tuple[int, ...]

__all__ = [
    "EntityCluster",
    "EntityOccurrence",
    "RowFillResult",
    "RowItem",
    "TemplateMatch",
    "TranslationUnit",
]
