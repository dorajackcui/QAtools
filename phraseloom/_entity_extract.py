from __future__ import annotations

import json
import re
from collections import Counter, defaultdict

from . import workbook_schema as schema
from ._entity_cluster_probe import find_entity_clusters
from ._entity_types import UnitRow
from .models import EntityCluster


def _build_entity_tm(
    rows: list[UnitRow],
    clusters: list[EntityCluster],
    min_group_size: int,
) -> tuple[list[dict[str, object]], list[dict[str, object]]]:
    structures: list[dict[str, object]] = []
    term_occurrences: dict[str, Counter[str]] = defaultdict(Counter)
    term_targets: dict[str, set[str]] = defaultdict(set)
    assigned_indexes: set[int] = set()

    for structure_index, cluster in enumerate(clusters, start=1):
        structure_id = f"TMES{structure_index:04d}"
        occurrences: list[tuple[UnitRow, dict[str, str]]] = []
        for cluster_row_number in cluster.row_numbers:
            row_index = cluster_row_number - 2
            if row_index < 0 or row_index >= len(rows):
                continue
            unit_row = rows[row_index]
            if unit_row.original_index in assigned_indexes:
                continue
            source_entities = _extract_entities(
                cluster.source_pattern,
                unit_row.source_unit,
            )
            if source_entities is None:
                continue
            occurrences.append((unit_row, source_entities))

        if not occurrences:
            continue

        target_structure, target_entity_rows = _derive_target_entity_pattern(
            occurrences,
            min_group_size=min(min_group_size, len(occurrences)),
        )
        for unit_row, source_entities in occurrences:
            assigned_indexes.add(unit_row.original_index)
            for source_entity in source_entities.values():
                term_occurrences[source_entity][structure_id] += 1
            target_entities = target_entity_rows.get(unit_row.original_index, {})
            for source_entity, target_entity in _ordered_entity_pairs(
                source_entities,
                target_entities,
            ):
                if target_entity:
                    term_targets[source_entity].add(target_entity)

        structures.append(
            {
                schema.STRUCTURE_ID_COLUMN: structure_id,
                schema.SOURCE_STRUCTURE_COLUMN: cluster.source_pattern,
                schema.TARGET_STRUCTURE_COLUMN: target_structure,
                schema.COVERAGE_COUNT_COLUMN: len(occurrences),
                schema.CONFIDENCE_COLUMN: cluster.confidence,
                schema.RISK_COLUMN: cluster.risk or None,
                schema.STATUS_COLUMN: "ready" if target_structure else "review",
                schema.SAMPLE_SOURCES_COLUMN: "\n".join(
                    dict.fromkeys(row.source_unit for row, _entities in occurrences[:10])
                ),
                schema.ROW_NUMBERS_COLUMN: ",".join(
                    str(row.original_index) for row, _entities in occurrences
                ),
                schema.WARNING_COLUMN: None
                if target_structure
                else "target entity structure not inferred",
            }
        )

    terms = []
    for index, (source_entity, structures_count) in enumerate(
        sorted(term_occurrences.items()),
        start=1,
    ):
        targets = term_targets.get(source_entity, set())
        target_entity = next(iter(targets)) if len(targets) == 1 else None
        warning = "ambiguous_entity_translation" if len(targets) > 1 else None
        if not targets:
            warning = "target entity not inferred"
        terms.append(
            {
                schema.TERM_ID_COLUMN: f"TMET{index:04d}",
                schema.SOURCE_ENTITY_COLUMN: source_entity,
                schema.TARGET_ENTITY_COLUMN: target_entity,
                schema.OCCURRENCE_COUNT_COLUMN: sum(structures_count.values()),
                schema.STRUCTURE_IDS_COLUMN: ",".join(sorted(structures_count)),
                schema.STATUS_COLUMN: "ready" if target_entity else "review",
                schema.WARNING_COLUMN: warning,
            }
        )
    return structures, terms


def _derive_target_entity_pattern(
    occurrences: list[tuple[UnitRow, dict[str, str]]],
    *,
    min_group_size: int,
) -> tuple[str | None, dict[int, dict[str, str]]]:
    if len(occurrences) < min_group_size:
        return None, {}
    target_rows = [(row.target_unit, "") for row, _entities in occurrences]
    target_clusters = find_entity_clusters(
        target_rows,
        min_group_size=min_group_size,
        max_entity_tokens=4,
        min_literal_tokens=3,
        top=1,
    )
    if not target_clusters:
        return None, {}

    target_cluster = target_clusters[0]
    target_entities_by_index: dict[int, dict[str, str]] = {}
    for target_cluster_row_number in target_cluster.row_numbers:
        occurrence_index = target_cluster_row_number - 2
        if occurrence_index < 0 or occurrence_index >= len(occurrences):
            continue
        unit_row, _source_entities = occurrences[occurrence_index]
        target_entities = _extract_entities(
            target_cluster.source_pattern,
            unit_row.target_unit,
        )
        if target_entities is not None:
            target_entities_by_index[unit_row.original_index] = target_entities
    return target_cluster.source_pattern, target_entities_by_index


def _ordered_entity_pairs(
    source_entities: dict[str, str],
    target_entities: dict[str, str],
) -> list[tuple[str, str | None]]:
    source_values = [
        value for key, value in sorted(source_entities.items(), key=lambda item: item[0])
    ]
    target_values = [
        value for key, value in sorted(target_entities.items(), key=lambda item: item[0])
    ]
    return [
        (source_entity, target_values[index] if index < len(target_values) else None)
        for index, source_entity in enumerate(source_values)
    ]


def _build_entity_split(
    rows: list[UnitRow],
    clusters: list[EntityCluster],
) -> tuple[
    list[UnitRow],
    list[dict[str, object]],
    list[dict[str, object]],
    list[dict[str, object]],
]:
    entity_rows: list[UnitRow] = []
    structures: list[dict[str, object]] = []
    source_map: list[dict[str, object]] = []
    term_occurrences: dict[str, Counter[str]] = defaultdict(Counter)
    term_sample_rows: dict[str, list[UnitRow]] = defaultdict(list)
    assigned_indexes: set[int] = set()

    for structure_index, cluster in enumerate(clusters, start=1):
        structure_id = f"ES{structure_index:04d}"
        structure_rows: list[UnitRow] = []
        structure_occurrences: list[tuple[UnitRow, dict[str, str]]] = []
        for cluster_row_number in cluster.row_numbers:
            row_index = cluster_row_number - 2
            if row_index < 0 or row_index >= len(rows):
                continue
            unit_row = rows[row_index]
            if unit_row.original_index in assigned_indexes:
                continue
            entities = _extract_entities(cluster.source_pattern, unit_row.source_unit)
            if entities is None:
                continue
            assigned_indexes.add(unit_row.original_index)
            structure_occurrences.append((unit_row, entities))

        for unit_row, entities in sorted(
            structure_occurrences,
            key=lambda item: item[0].original_index,
        ):
            entity_rows.append(unit_row)
            structure_rows.append(unit_row)
            for source_entity in entities.values():
                term_occurrences[source_entity][structure_id] += 1
                term_sample_rows[source_entity].append(unit_row)
            source_map.append(
                {
                    schema.ORIGINAL_INDEX_COLUMN: unit_row.original_index,
                    schema.UNIT_ID_COLUMN: unit_row.unit_id,
                    schema.UNIT_TYPE_COLUMN: unit_row.unit_type,
                    schema.SOURCE_UNIT_COLUMN: unit_row.source_unit,
                    schema.STRUCTURE_ID_COLUMN: structure_id,
                    schema.ENTITIES_JSON_COLUMN: json.dumps(
                        entities,
                        ensure_ascii=False,
                        sort_keys=True,
                    ),
                    schema.PREVIEW_TARGET_COLUMN: None,
                    schema.FILL_STATUS_COLUMN: None,
                    schema.WARNING_COLUMN: None,
                }
            )
        if not structure_rows:
            continue
        sample_sources, sample_context = _sample_sources_and_context(structure_rows)
        structures.append(
            {
                schema.STRUCTURE_ID_COLUMN: structure_id,
                schema.SOURCE_STRUCTURE_COLUMN: cluster.source_pattern,
                schema.TARGET_STRUCTURE_COLUMN: None,
                schema.COVERAGE_COUNT_COLUMN: len(structure_rows),
                schema.CONFIDENCE_COLUMN: cluster.confidence,
                schema.RISK_COLUMN: cluster.risk or None,
                schema.STATUS_COLUMN: _default_structure_status(cluster),
                schema.SAMPLE_SOURCES_COLUMN: sample_sources,
                schema.SAMPLE_CONTEXT_COLUMN: sample_context,
                schema.ROW_NUMBERS_COLUMN: ",".join(
                    str(row.original_index) for row in structure_rows
                ),
                schema.WARNING_COLUMN: cluster.risk or None,
            }
        )

    terms = []
    for index, (source_entity, structures_count) in enumerate(
        sorted(term_occurrences.items()),
        start=1,
    ):
        sample_sources, sample_context = _sample_sources_and_context(
            term_sample_rows[source_entity],
        )
        terms.append(
            {
                schema.TERM_ID_COLUMN: f"ET{index:04d}",
                schema.SOURCE_ENTITY_COLUMN: source_entity,
                schema.TARGET_ENTITY_COLUMN: None,
                schema.OCCURRENCE_COUNT_COLUMN: sum(structures_count.values()),
                schema.STRUCTURE_IDS_COLUMN: ",".join(sorted(structures_count)),
                schema.STATUS_COLUMN: "review",
                schema.SAMPLE_SOURCES_COLUMN: sample_sources,
                schema.SAMPLE_CONTEXT_COLUMN: sample_context,
                schema.WARNING_COLUMN: None,
            }
        )
    source_map.sort(key=lambda row: int(row[schema.ORIGINAL_INDEX_COLUMN]))
    return entity_rows, structures, terms, source_map


def _sample_sources_and_context(
    rows: list[UnitRow],
    *,
    limit: int = 10,
) -> tuple[str | None, str | None]:
    sample_sources: list[str] = []
    sample_contexts: list[str] = []
    seen_sources: set[str] = set()
    for row in rows:
        if row.source_unit in seen_sources:
            continue
        seen_sources.add(row.source_unit)
        sample_sources.append(row.source_unit)
        sample_contexts.append(str(row.values.get(schema.CONTEXT_COLUMN) or ""))
        if len(sample_sources) >= limit:
            break
    return "\n".join(sample_sources) or None, "\n".join(sample_contexts) or None


def _default_structure_status(cluster: EntityCluster) -> str:
    if cluster.confidence >= 0.85 and not cluster.risk:
        return "ready"
    return "review"


def _extract_entities(source_structure: str, source_unit: str) -> dict[str, str] | None:
    placeholders = re.findall(r"\{entity\d+\}", source_structure)
    if not placeholders:
        return None
    regex_parts: list[str] = ["^"]
    position = 0
    used_names: set[str] = set()
    for found in re.finditer(r"\{entity\d+\}", source_structure):
        regex_parts.append(re.escape(source_structure[position : found.start()]))
        name = found.group(0)[1:-1]
        if name in used_names:
            regex_parts.append(f"(?P={name})")
        else:
            regex_parts.append(f"(?P<{name}>.+?)")
            used_names.add(name)
        position = found.end()
    regex_parts.append(re.escape(source_structure[position:]))
    regex_parts.append("$")
    matched = re.match("".join(regex_parts), source_unit)
    if not matched:
        return None
    return {name: matched.group(name) for name in sorted(used_names)}
