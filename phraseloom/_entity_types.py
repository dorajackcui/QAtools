from __future__ import annotations

from dataclasses import dataclass
from typing import Protocol

from . import workbook_schema as schema
from ._entity_cluster_probe import find_entity_clusters
from .models import EntityCluster


ENTITY_STRUCTURE_COLUMNS = [
    schema.STRUCTURE_ID_COLUMN,
    schema.SOURCE_STRUCTURE_COLUMN,
    schema.TARGET_STRUCTURE_COLUMN,
    schema.COVERAGE_COUNT_COLUMN,
    schema.CONFIDENCE_COLUMN,
    schema.RISK_COLUMN,
    schema.STATUS_COLUMN,
    schema.SAMPLE_SOURCES_COLUMN,
    schema.ROW_NUMBERS_COLUMN,
    schema.WARNING_COLUMN,
]
ENTITY_PACK_STRUCTURE_COLUMNS = [
    schema.STRUCTURE_ID_COLUMN,
    schema.SOURCE_STRUCTURE_COLUMN,
    schema.TARGET_STRUCTURE_COLUMN,
    schema.COVERAGE_COUNT_COLUMN,
    schema.CONFIDENCE_COLUMN,
    schema.RISK_COLUMN,
    schema.SAMPLE_SOURCES_COLUMN,
    schema.SAMPLE_CONTEXT_COLUMN,
    schema.ROW_NUMBERS_COLUMN,
    schema.WARNING_COLUMN,
]
ENTITY_TERM_COLUMNS = [
    schema.TERM_ID_COLUMN,
    schema.SOURCE_ENTITY_COLUMN,
    schema.TARGET_ENTITY_COLUMN,
    schema.OCCURRENCE_COUNT_COLUMN,
    schema.STRUCTURE_IDS_COLUMN,
    schema.STATUS_COLUMN,
    schema.WARNING_COLUMN,
]
ENTITY_PACK_TERM_COLUMNS = [
    schema.TERM_ID_COLUMN,
    schema.SOURCE_ENTITY_COLUMN,
    schema.TARGET_ENTITY_COLUMN,
    schema.OCCURRENCE_COUNT_COLUMN,
    schema.STRUCTURE_IDS_COLUMN,
    schema.SAMPLE_SOURCES_COLUMN,
    schema.SAMPLE_CONTEXT_COLUMN,
    schema.WARNING_COLUMN,
]
ENTITY_SOURCE_MAP_COLUMNS = [
    schema.ORIGINAL_INDEX_COLUMN,
    schema.UNIT_ID_COLUMN,
    schema.UNIT_TYPE_COLUMN,
    schema.SOURCE_UNIT_COLUMN,
    schema.STRUCTURE_ID_COLUMN,
    schema.ENTITIES_JSON_COLUMN,
    schema.PREVIEW_TARGET_COLUMN,
    schema.FILL_STATUS_COLUMN,
    schema.WARNING_COLUMN,
]


@dataclass(frozen=True)
class UnitRow:
    original_index: int
    values: dict[str, object]

    def _value(self, column: str) -> object:
        for candidate in schema.UNIT_COLUMN_ALIASES.get(column, (column,)):
            value = self.values.get(candidate)
            if value not in (None, ""):
                return value
        return ""

    @property
    def unit_id(self) -> str:
        return str(self.values.get(schema.UNIT_ID_COLUMN) or "")

    @property
    def unit_type(self) -> str:
        return str(self.values.get(schema.UNIT_TYPE_COLUMN) or "")

    @property
    def source_unit(self) -> str:
        return str(self._value(schema.SOURCE_UNIT_COLUMN) or "")

    @property
    def target_unit(self) -> str:
        return str(self._value(schema.TARGET_UNIT_COLUMN) or "")


class EntityExtractionStrategy(Protocol):
    name: str

    def find_clusters(self, rows: list[tuple[str, str]]) -> list[EntityCluster]:
        ...


@dataclass(frozen=True)
class ClusterProbeStrategy:
    min_group_size: int = 3
    max_entity_tokens: int = 4
    min_literal_tokens: int = 3
    top: int = 200
    name: str = "cluster_probe"

    def find_clusters(self, rows: list[tuple[str, str]]) -> list[EntityCluster]:
        return find_entity_clusters(
            rows,
            min_group_size=self.min_group_size,
            max_entity_tokens=self.max_entity_tokens,
            min_literal_tokens=self.min_literal_tokens,
            top=self.top,
        )
