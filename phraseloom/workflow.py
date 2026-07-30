from __future__ import annotations

import json
import re
from collections import Counter, defaultdict
from pathlib import Path
from typing import Iterable

from . import workbook_schema as schema
from .excel_io import (
    _default_package_fill_output_path,
    _default_restore_audit_output_path,
    _default_to_translate_output_path,
    _load_package_tag_rules,
    _load_translated_units,
    _read_source_rows,
    _translation_package_metadata,
    _write_output_workbook,
    _write_restore_audit_workbook,
    _write_target_column_workbook,
    _write_tm_workbook,
    _write_to_translate_workbook,
    _write_translation_package_target_workbook,
)
from .errors import ConfigError
from .models import RowFillResult, RowItem, TranslationUnit
from .tag_engine import (
    extract_tags,
    is_tag_only_segment,
    is_tag_placeholder,
    restore_tags,
    serialize_known_tags,
    validate_tag_placeholders,
)
from .tag_rules import TagRules, load_tag_rules
from .template_engine import (
    PLACEHOLDER_RE,
    apply_target_template,
    infer_target_template,
    is_candidate_template,
    is_non_translatable_segment,
    parse_template,
)

_is_candidate = is_candidate_template
_is_non_translatable_segment = is_non_translatable_segment


def generate_workbook(
    input_path: str | Path,
    output_path: str | Path,
    *,
    source_col: str | int = "source",
    target_col: str | int | None = "target",
    context_col: str | int | None = None,
    examples: Iterable[tuple[str, str]] = (),
    template_workbook: str | Path | None = None,
    tm_workbook: str | Path | None = None,
    min_group_size: int = 2,
    use_existing_targets: bool = True,
    tag_config: str | Path | None = None,
) -> dict[str, int]:
    input_path = Path(input_path)
    output_path = Path(output_path)
    tag_rules = load_tag_rules(tag_config)

    rows, units, result_rows, autofilled_count = _build_fill_context(
        input_path,
        source_col=source_col,
        target_col=target_col,
        examples=examples,
        template_workbook=template_workbook,
        tm_workbook=tm_workbook,
        min_group_size=min_group_size,
        use_existing_targets=use_existing_targets,
        tag_rules=tag_rules,
    )

    _write_output_workbook(
        output_path,
        input_path,
        units,
        result_rows,
        tag_rules=tag_rules,
        context_col=context_col,
    )
    to_translate_path = _default_to_translate_output_path(input_path)
    if not template_workbook or Path(template_workbook).resolve() != to_translate_path.resolve():
        _write_to_translate_workbook(
            to_translate_path,
            input_path,
            units,
            tag_rules=tag_rules,
            source_col=source_col,
            target_col=target_col,
            context_col=context_col,
            min_group_size=min_group_size,
            use_existing_targets=use_existing_targets,
        )

    stats = _workbook_stats(rows, units, autofilled_count)
    stats["to_translate_path"] = str(to_translate_path)
    return stats


def prepare_translation_package(
    input_path: str | Path,
    output_path: str | Path | None = None,
    *,
    source_col: str | int = "source",
    target_col: str | int | None = "target",
    context_col: str | int | None = None,
    tm_workbook: str | Path | None = None,
    min_group_size: int = 2,
    use_existing_targets: bool = False,
    tag_config: str | Path | None = None,
) -> dict[str, int | str]:
    input_path = Path(input_path)
    package_path = (
        Path(output_path)
        if output_path is not None
        else _default_to_translate_output_path(input_path)
    )
    tag_rules = load_tag_rules(tag_config)
    rows, units, _result_rows, autofilled_count = _build_fill_context(
        input_path,
        source_col=source_col,
        target_col=target_col,
        tm_workbook=tm_workbook,
        min_group_size=min_group_size,
        use_existing_targets=use_existing_targets,
        tag_rules=tag_rules,
    )
    _write_to_translate_workbook(
        package_path,
        input_path,
        units,
        tag_rules=tag_rules,
        source_col=source_col,
        target_col=target_col,
        context_col=context_col,
        min_group_size=min_group_size,
        use_existing_targets=use_existing_targets,
    )
    stats: dict[str, int | str] = _workbook_stats(rows, units, autofilled_count)
    stats["to_translate_path"] = str(package_path)
    return stats


def fill_translation_package(
    package_path: str | Path,
    output_path: str | Path | None = None,
    *,
    tag_config: str | Path | None = None,
    audit_output_path: str | Path | None = None,
) -> dict[str, int | str]:
    package_path = Path(package_path)
    metadata = _translation_package_metadata(package_path)
    source_col = _metadata_column(metadata, schema.SOURCE_COLUMN_KEY)
    target_col = _metadata_column(metadata, schema.TARGET_COLUMN_KEY)
    if target_col is None:
        raise ConfigError(
            "Translation package does not define a target column. "
            "Prepare it again and choose a target column."
        )
    try:
        min_group_size = int(metadata.get(schema.MIN_GROUP_SIZE_KEY, 2))
    except (TypeError, ValueError) as exc:
        raise ConfigError("Translation package has an invalid reusable-part setting.") from exc
    tag_rules = (
        load_tag_rules(tag_config)
        if tag_config
        else _load_package_tag_rules(package_path)
    )
    filled_path = (
        Path(output_path)
        if output_path is not None
        else _default_package_fill_output_path(package_path)
    )

    rows, units, result_rows, autofilled_count = _build_fill_context(
        package_path,
        source_col=source_col,
        target_col=target_col,
        template_workbook=package_path,
        min_group_size=min_group_size,
        use_existing_targets=False,
        tag_rules=tag_rules,
    )
    _write_translation_package_target_workbook(
        filled_path,
        package_path,
        target_col,
        result_rows,
    )

    should_write_audit = audit_output_path is not None or _needs_restore_audit(result_rows)
    audit_path: Path | None = None
    if should_write_audit:
        audit_path = (
            Path(audit_output_path)
            if audit_output_path is not None
            else _default_restore_audit_output_path(filled_path)
        )
        _write_restore_audit_workbook(
            audit_path,
            package_path,
            filled_path,
            result_rows,
            tag_rules=tag_rules,
        )

    stats = _workbook_stats(rows, units, autofilled_count)
    stats["output_path"] = str(filled_path)
    if audit_path is not None:
        stats["audit_output_path"] = str(audit_path)
    return stats


def _metadata_column(metadata: dict[str, object], key: str) -> str | int | None:
    value = metadata.get(key)
    if value is None:
        raise ConfigError(f"Translation package is missing metadata: {key}")
    try:
        parsed = json.loads(str(value))
    except (TypeError, ValueError) as exc:
        raise ConfigError(f"Translation package has invalid metadata: {key}") from exc
    if parsed is None or isinstance(parsed, (str, int)):
        return parsed
    raise ConfigError(f"Translation package has invalid metadata: {key}")


def _needs_restore_audit(result_rows: list[RowFillResult]) -> bool:
    return any(
        result_row.unit is None
        or not result_row.unit.target_unit
        or bool(result_row.unit.warning)
        or bool(result_row.warning)
        or bool(result_row.row.tag_warnings)
        or bool(result_row.row.target_tag_warnings)
        for result_row in result_rows
    )


def fill_target_column_workbook(
    input_path: str | Path,
    output_path: str | Path,
    *,
    source_col: str | int = "source",
    target_col: str | int = "target",
    template_workbook: str | Path,
    tm_workbook: str | Path | None = None,
    min_group_size: int = 2,
    tag_config: str | Path | None = None,
    audit_output_path: str | Path | None = None,
) -> dict[str, int | str]:
    input_path = Path(input_path)
    output_path = Path(output_path)
    tag_rules = load_tag_rules(tag_config)

    rows, units, result_rows, autofilled_count = _build_fill_context(
        input_path,
        source_col=source_col,
        target_col=target_col,
        template_workbook=template_workbook,
        tm_workbook=tm_workbook,
        min_group_size=min_group_size,
        use_existing_targets=False,
        tag_rules=tag_rules,
    )
    _write_target_column_workbook(output_path, input_path, target_col, result_rows)
    audit_path = (
        Path(audit_output_path)
        if audit_output_path is not None
        else _default_restore_audit_output_path(output_path)
    )
    _write_restore_audit_workbook(
        audit_path,
        input_path,
        output_path,
        result_rows,
        tag_rules=tag_rules,
    )
    stats = _workbook_stats(rows, units, autofilled_count)
    stats["audit_output_path"] = str(audit_path)
    return stats


def _build_fill_context(
    input_path: Path,
    *,
    source_col: str | int,
    target_col: str | int | None,
    examples: Iterable[tuple[str, str]] = (),
    template_workbook: str | Path | None = None,
    tm_workbook: str | Path | None = None,
    min_group_size: int,
    use_existing_targets: bool,
    tag_rules: TagRules,
) -> tuple[
    list[RowItem],
    list[TranslationUnit],
    list[RowFillResult],
    int,
]:
    rows = _read_source_rows(input_path, source_col, target_col, tag_rules=tag_rules)
    provided_units, provided_sources = _build_provided_units(
        examples, template_workbook, tm_workbook, tag_rules=tag_rules
    )
    units = _build_translation_units(
        rows,
        min_group_size,
        provided_units,
        provided_sources,
        use_existing_targets,
    )

    unit_by_row_number = {
        item.row_number: unit for unit in units for item in unit.items
    }
    result_rows = []
    autofilled_count = 0
    for row in rows:
        unit = unit_by_row_number.get(row.row_number)
        target_template = unit.target_unit if unit else ""
        serialized_target = (
            apply_target_template(target_template, row.match.values)
            if target_template and unit and unit.unit_type == "template"
            else target_template
            if target_template
            else None
        )
        auto_target = None
        warning = ""
        if serialized_target:
            validation = validate_tag_placeholders(serialized_target, row.tag_tokens)
            warning = _merge_warning_parts(*validation.warnings)
            auto_target = restore_tags(serialized_target, row.tag_tokens)
        if auto_target:
            autofilled_count += 1
        result_rows.append(
            RowFillResult(row=row, unit=unit, auto_target=auto_target, warning=warning)
        )

    return rows, units, result_rows, autofilled_count


def _workbook_stats(
    rows: list[RowItem], units: list[TranslationUnit], autofilled_count: int
) -> dict[str, int]:
    template_units = [unit for unit in units if unit.unit_type == "template"]
    segment_units = [unit for unit in units if unit.unit_type == "segment"]
    template_source_segments = sum(unit.coverage_count for unit in template_units)
    segment_source_segments = sum(unit.coverage_count for unit in segment_units)
    unique_source_segments = len({row.source for row in rows})
    prefilled_translation_unit_count = sum(1 for unit in units if unit.target_unit)
    return {
        "row_count": len(rows),
        "unique_source_segments": unique_source_segments,
        "duplicate_source_segments": len(rows) - unique_source_segments,
        "template_count": len(template_units),
        "template_unit_count": len(template_units),
        "template_source_segments": template_source_segments,
        "template_unique_source_segments": sum(
            unit.unique_source_count for unit in template_units
        ),
        "segment_unit_count": len(segment_units),
        "segment_source_segments": segment_source_segments,
        "translation_unit_count": len(units),
        "prefilled_translation_unit_count": prefilled_translation_unit_count,
        "untranslated_translation_unit_count": len(units)
        - prefilled_translation_unit_count,
        "new_translation_unit_count": len(units) - prefilled_translation_unit_count,
        "new_source_segment_count": sum(
            unit.coverage_count for unit in units if not unit.target_unit
        ),
        "tm_unit_hit_rate": _format_rate(prefilled_translation_unit_count, len(units)),
        "tm_row_hit_rate": _format_rate(autofilled_count, len(rows)),
        "clustered_source_segments": template_source_segments,
        "unclustered_source_segments": segment_source_segments,
        "autofilled_count": autofilled_count,
    }


def generate_tm_pairs(
    input_path: str | Path,
    output_path: str | Path,
    *,
    source_col: str | int = "source",
    target_col: str | int = "target",
    context_col: str | int | None = None,
    min_group_size: int = 2,
    tag_config: str | Path | None = None,
) -> dict[str, int]:
    input_path = Path(input_path)
    output_path = Path(output_path)
    tag_rules = load_tag_rules(tag_config)
    rows = [
        row
        for row in _read_source_rows(
            input_path, source_col, target_col, tag_rules=tag_rules
        )
        if row.existing_target
    ]
    units = _build_translation_units(
        rows,
        min_group_size,
        provided_units={},
        provided_sources={},
        use_existing_targets=True,
    )
    _write_tm_workbook(
        output_path,
        input_path,
        units,
        rows,
        tag_rules=tag_rules,
        context_col=context_col,
    )

    template_pairs = [unit for unit in units if unit.unit_type == "template"]
    segment_pairs = [unit for unit in units if unit.unit_type == "segment"]
    unique_source_segments = len({row.source for row in rows})
    return {
        "row_count": len(rows),
        "unique_source_segments": unique_source_segments,
        "duplicate_source_segments": len(rows) - unique_source_segments,
        "tm_pair_count": len(units),
        "template_pair_count": len(template_pairs),
        "segment_pair_count": len(segment_pairs),
        "matched_source_segments": sum(unit.coverage_count for unit in units),
    }


def _eligible_groups(
    rows: Iterable[RowItem], min_group_size: int
) -> dict[str, list[RowItem]]:
    grouped: dict[str, list[RowItem]] = defaultdict(list)
    for row in rows:
        if _is_candidate(row.match):
            grouped[row.match.template].append(row)
    return {
        template: items
        for template, items in grouped.items()
        if len({item.source for item in items}) >= min_group_size
    }


def _build_translation_units(
    rows: list[RowItem],
    min_group_size: int,
    provided_units: dict[tuple[str, str], str],
    provided_sources: dict[tuple[str, str], str],
    use_existing_targets: bool,
) -> list[TranslationUnit]:
    template_groups = _eligible_groups(rows, min_group_size)
    assigned_row_numbers: set[int] = set()
    units: list[TranslationUnit] = []

    for row in rows:
        key = ("template", row.match.template)
        if _is_candidate(row.match) and key in provided_units:
            existing = template_groups.setdefault(row.match.template, [])
            if row not in existing:
                existing.append(row)

    for index, (source_unit, items) in enumerate(
        sorted(template_groups.items(), key=lambda item: (-len(item[1]), item[0])),
        start=1,
    ):
        for item in items:
            assigned_row_numbers.add(item.row_number)
        suggested = _suggest_template_target_unit(items) if use_existing_targets else ""
        key = ("template", source_unit)
        target_unit = provided_units.get(key) or suggested
        target_unit_source = (
            provided_sources.get(key, "")
            if key in provided_units
            else "existing_target"
            if suggested
            else ""
        )
        warning = _unit_warning("template", source_unit, target_unit, suggested, items)
        units.append(
            TranslationUnit(
                unit_id=f"T{index:04d}",
                unit_type="template",
                source_unit=source_unit,
                coverage_count=len(items),
                unique_source_count=len({item.source for item in items}),
                items=tuple(items),
                target_unit=target_unit,
                target_unit_source=target_unit_source,
                suggested_target_unit=suggested,
                warning=warning,
            )
        )

    segment_groups: dict[str, list[RowItem]] = defaultdict(list)
    for row in rows:
        if row.row_number not in assigned_row_numbers:
            segment_groups[row.source].append(row)

    for index, (source_unit, items) in enumerate(
        sorted(segment_groups.items(), key=lambda item: (-len(item[1]), item[0])),
        start=1,
    ):
        suggested = _suggest_segment_target_unit(items) if use_existing_targets else ""
        key = ("segment", source_unit)
        if key in provided_units:
            target_unit = provided_units[key]
            target_unit_source = provided_sources.get(key, "")
        elif suggested:
            target_unit = suggested
            target_unit_source = "existing_target"
        elif is_tag_only_segment(source_unit):
            target_unit = source_unit
            target_unit_source = "tag_only"
        elif _is_non_translatable_segment(source_unit):
            target_unit = source_unit
            target_unit_source = "non_translatable"
        else:
            target_unit = ""
            target_unit_source = ""
        warning = _unit_warning("segment", source_unit, target_unit, suggested, items)
        units.append(
            TranslationUnit(
                unit_id=f"S{index:04d}",
                unit_type="segment",
                source_unit=source_unit,
                coverage_count=len(items),
                unique_source_count=1,
                items=tuple(items),
                target_unit=target_unit,
                target_unit_source=target_unit_source,
                suggested_target_unit=suggested,
                warning=warning,
            )
        )

    return units


def _build_provided_units(
    examples: Iterable[tuple[str, str]],
    template_workbook: str | Path | None,
    tm_workbook: str | Path | None = None,
    *,
    tag_rules: TagRules,
) -> tuple[dict[tuple[str, str], str], dict[tuple[str, str], str]]:
    provided: dict[tuple[str, str], str] = {}
    sources: dict[tuple[str, str], str] = {}

    if tm_workbook:
        for key, target_unit in _load_translated_units(
            Path(tm_workbook), tag_rules=tag_rules
        ).items():
            provided[key] = target_unit
            sources[key] = "tm_pairs"

    if template_workbook:
        for key, target_unit in _load_translated_units(
            Path(template_workbook), tag_rules=tag_rules
        ).items():
            provided[key] = target_unit
            sources[key] = "translation_units"

    for source, target in examples:
        source_extraction = extract_tags(source, rules=tag_rules)
        serialized_source = source_extraction.text
        serialized_target = serialize_known_tags(target, source_extraction.tags).text
        match = parse_template(serialized_source)
        target_template = infer_target_template(match.values, serialized_target)
        if target_template:
            key = ("template", match.template)
            provided[key] = target_template
            sources[key] = (
                f"example: {source} => {target}"
                if serialized_source == source and serialized_target == target
                else f"example: {source} => {target} ({serialized_source} => {serialized_target})"
            )
        else:
            key = ("segment", serialized_source)
            provided[key] = serialized_target
            sources[key] = (
                f"example: {source} => {target}"
                if serialized_source == source and serialized_target == target
                else f"example: {source} => {target} ({serialized_source} => {serialized_target})"
            )

    return provided, sources


def _suggest_template_target_unit(items: Iterable[RowItem]) -> str:
    suggestions: list[str] = []
    for item in items:
        if not item.existing_target:
            continue
        inferred = infer_target_template(item.match.values, item.existing_target)
        if inferred:
            suggestions.append(inferred)
    if not suggestions:
        return ""
    return Counter(suggestions).most_common(1)[0][0]


def _suggest_segment_target_unit(items: Iterable[RowItem]) -> str:
    suggestions = [item.existing_target for item in items if item.existing_target]
    if not suggestions:
        return ""
    return Counter(suggestions).most_common(1)[0][0]


def _row_serialization_warnings(items: Iterable[RowItem]) -> list[str]:
    warnings: list[str] = []
    for item in items:
        warnings.extend(item.tag_warnings)
        warnings.extend(item.target_tag_warnings)
    return list(dict.fromkeys(warnings))


def _unit_warning(
    unit_type: str,
    source_unit: str,
    target_unit: str,
    suggested_target_unit: str,
    items: Iterable[RowItem],
) -> str:
    item_list = list(items)
    warnings: list[str] = []
    warnings.extend(_row_serialization_warnings(item_list))
    source_placeholders = {
        placeholder
        for placeholder in PLACEHOLDER_RE.findall(source_unit)
        if not is_tag_placeholder(placeholder)
    }
    target_placeholders = {
        placeholder
        for placeholder in PLACEHOLDER_RE.findall(target_unit)
        if not is_tag_placeholder(placeholder)
    }

    if unit_type == "template" and target_unit and source_placeholders - target_placeholders:
        warnings.append("target_unit is missing source variables")
    if "$" in source_unit:
        warnings.append("price-like text; review manually")
    if re.search(r"\b1\s+(day|time|attempt|task|star|pack)s\b", source_unit):
        warnings.append("plural-sensitive text; review manually")

    inferred = []
    if unit_type == "template":
        for item in item_list:
            if item.existing_target:
                guess = infer_target_template(item.match.values, item.existing_target)
                if guess:
                    inferred.append(guess)
    else:
        inferred = [item.existing_target for item in item_list if item.existing_target]
    if suggested_target_unit and len(set(inferred)) > 1:
        warnings.append("multiple existing target patterns found")

    return "; ".join(warnings)


def _merge_warning_parts(*warnings: str) -> str:
    return "; ".join(warning for warning in warnings if warning)


def _format_rate(numerator: int, denominator: int) -> str:
    if denominator == 0:
        return "0.0%"
    return f"{numerator / denominator * 100:.1f}%"


extract_tm_pairs = generate_tm_pairs
prepare_translation = prepare_translation_package
fill_translation = fill_translation_package

__all__ = [
    "extract_tm_pairs",
    "fill_target_column_workbook",
    "fill_translation",
    "fill_translation_package",
    "generate_tm_pairs",
    "generate_workbook",
    "prepare_translation",
    "prepare_translation_package",
]
