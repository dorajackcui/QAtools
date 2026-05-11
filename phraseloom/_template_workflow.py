from __future__ import annotations

import argparse
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path
from typing import Iterable

from .excel_io import (
    _default_extract_output_path,
    _default_fill_output_path,
    _default_legacy_output_path,
    _default_tm_output_path,
    _default_to_translate_output_path,
    _default_work_dir,
    _load_translated_units,
    _read_source_rows,
    _write_output_workbook,
    _write_target_column_workbook,
    _write_tm_workbook,
    _write_to_translate_workbook,
)
from .models import RowItem, TranslationUnit
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
    source_col: str | int = "英語",
    target_col: str | int | None = "target",
    examples: Iterable[tuple[str, str]] = (),
    template_workbook: str | Path | None = None,
    tm_workbook: str | Path | None = None,
    min_group_size: int = 2,
    use_existing_targets: bool = True,
) -> dict[str, int]:
    input_path = Path(input_path)
    output_path = Path(output_path)

    rows, units, result_rows, autofilled_count = _build_fill_context(
        input_path,
        source_col=source_col,
        target_col=target_col,
        examples=examples,
        template_workbook=template_workbook,
        tm_workbook=tm_workbook,
        min_group_size=min_group_size,
        use_existing_targets=use_existing_targets,
    )

    _write_output_workbook(output_path, input_path, units, result_rows)
    to_translate_path = _default_to_translate_output_path(input_path)
    if not template_workbook or Path(template_workbook).resolve() != to_translate_path.resolve():
        _write_to_translate_workbook(to_translate_path, input_path, units)

    stats = _workbook_stats(rows, units, autofilled_count)
    stats["to_translate_path"] = str(to_translate_path)
    return stats


def fill_target_column_workbook(
    input_path: str | Path,
    output_path: str | Path,
    *,
    source_col: str | int = "英語",
    target_col: str | int = "target",
    template_workbook: str | Path,
    tm_workbook: str | Path | None = None,
    min_group_size: int = 2,
) -> dict[str, int]:
    input_path = Path(input_path)
    output_path = Path(output_path)

    rows, units, result_rows, autofilled_count = _build_fill_context(
        input_path,
        source_col=source_col,
        target_col=target_col,
        template_workbook=template_workbook,
        tm_workbook=tm_workbook,
        min_group_size=min_group_size,
        use_existing_targets=False,
    )
    _write_target_column_workbook(output_path, input_path, target_col, result_rows)
    return _workbook_stats(rows, units, autofilled_count)


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
) -> tuple[
    list[RowItem],
    list[TranslationUnit],
    list[tuple[RowItem, TranslationUnit | None, str | None]],
    int,
]:
    rows = _read_source_rows(input_path, source_col, target_col)
    provided_units, provided_sources = _build_provided_units(
        examples, template_workbook, tm_workbook
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
        auto_target = (
            apply_target_template(target_template, row.match.values)
            if target_template and unit and unit.unit_type == "template"
            else target_template
            if target_template
            else None
        )
        if auto_target:
            autofilled_count += 1
        result_rows.append((row, unit, auto_target))

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
    source_col: str | int = "英語",
    target_col: str | int = "target",
    min_group_size: int = 2,
) -> dict[str, int]:
    input_path = Path(input_path)
    output_path = Path(output_path)
    rows = [
        row
        for row in _read_source_rows(input_path, source_col, target_col)
        if row.existing_target
    ]
    units = _build_translation_units(
        rows,
        min_group_size,
        provided_units={},
        provided_sources={},
        use_existing_targets=True,
    )
    _write_tm_workbook(output_path, input_path, units, rows)

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
) -> tuple[dict[tuple[str, str], str], dict[tuple[str, str], str]]:
    provided: dict[tuple[str, str], str] = {}
    sources: dict[tuple[str, str], str] = {}

    if tm_workbook:
        for key, target_unit in _load_translated_units(Path(tm_workbook)).items():
            provided[key] = target_unit
            sources[key] = "tm_pairs"

    if template_workbook:
        for key, target_unit in _load_translated_units(
            Path(template_workbook)
        ).items():
            provided[key] = target_unit
            sources[key] = "translation_units"

    for source, target in examples:
        match = parse_template(source)
        target_template = infer_target_template(match.values, target)
        if target_template:
            key = ("template", match.template)
            provided[key] = target_template
            sources[key] = f"example: {source} => {target}"
        else:
            key = ("segment", source)
            provided[key] = target
            sources[key] = f"example: {source} => {target}"

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


def _unit_warning(
    unit_type: str,
    source_unit: str,
    target_unit: str,
    suggested_target_unit: str,
    items: Iterable[RowItem],
) -> str:
    warnings: list[str] = []
    source_placeholders = set(PLACEHOLDER_RE.findall(source_unit))
    target_placeholders = set(PLACEHOLDER_RE.findall(target_unit))

    if unit_type == "template" and target_unit and source_placeholders - target_placeholders:
        warnings.append("target_unit is missing source variables")
    if "$" in source_unit:
        warnings.append("price-like text; review manually")
    if re.search(r"\b1\s+(day|time|attempt|task|star|pack)s\b", source_unit):
        warnings.append("plural-sensitive text; review manually")

    inferred = []
    if unit_type == "template":
        for item in items:
            if item.existing_target:
                guess = infer_target_template(item.match.values, item.existing_target)
                if guess:
                    inferred.append(guess)
    else:
        inferred = [item.existing_target for item in items if item.existing_target]
    if suggested_target_unit and len(set(inferred)) > 1:
        warnings.append("multiple existing target patterns found")

    return "; ".join(warnings)


def _format_rate(numerator: int, denominator: int) -> str:
    if denominator == 0:
        return "0.0%"
    return f"{numerator / denominator * 100:.1f}%"


def _parse_examples(raw_examples: Iterable[str]) -> list[tuple[str, str]]:
    examples: list[tuple[str, str]] = []
    for raw in raw_examples:
        if "=" not in raw:
            raise ValueError(f"Example must look like SOURCE=TARGET: {raw!r}")
        source, target = raw.split("=", 1)
        examples.append((source.strip(), target.strip()))
    return examples


def main(argv: list[str] | None = None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)

    if argv and argv[0] in {"-h", "--help"}:
        _print_top_level_help()
        return 0
    if not argv or argv[0] in {"interactive", "wizard"}:
        return run_interactive()
    if argv[0] in {"tm-extract", "extract-tm"}:
        return _main_tm_extract(argv[1:])
    if argv[0] == "extract":
        return _main_extract(argv[1:])
    if argv[0] == "fill":
        return _main_fill(argv[1:])
    return _main_legacy(argv)


def run_interactive() -> int:
    print("Localization Workflow")
    print()
    print("1) Build TM from completed Excel")
    print("2) Prepare translator file for new source")
    print("3) Fill source from translated file")
    print("q) Quit")

    action = _prompt_text("Choose step", default="2").lower()
    if action in {"q", "quit", "exit"}:
        print("Bye.")
        return 0
    if action in {"1", "tm", "tm-extract", "extract-tm", "build"}:
        return _interactive_tm_extract()
    if action in {"2", "extract", "prepare", "p"}:
        return _interactive_extract()
    if action in {"3", "fill", "f"}:
        return _interactive_fill()

    print(f"Unknown step: {action}")
    return 2


def _interactive_tm_extract() -> int:
    input_path = _user_path(_prompt_text("Completed Excel path", required=True))
    source_col = _prompt_text("Source column in completed Excel", default="英語")
    target_col = _prompt_text("Target column in completed Excel", default="target")
    output_path = _user_path(
        _prompt_text("Output tm_pairs workbook", default=str(_default_tm_output_path(input_path)))
    )
    min_group_size = _prompt_int("Minimum variants for a reusable template", default=2)

    stats = generate_tm_pairs(
        input_path,
        output_path,
        source_col=source_col,
        target_col=target_col,
        min_group_size=min_group_size,
    )
    _print_tm_stats(output_path, stats)
    return 0


def _interactive_extract() -> int:
    input_path = _user_path(_prompt_text("New source Excel path", required=True))
    source_col = _prompt_text("Source column in new file", default="英語")
    target_col = _normalize_optional_column(
        _prompt_text("Existing target column (- for none)", default="target")
    )
    tm_workbook_text = _prompt_text("Existing tm_pairs path (- for none)", default="-")
    tm_workbook = (
        _user_path(tm_workbook_text)
        if _normalize_optional_column(tm_workbook_text) is not None
        else None
    )
    output_path = _user_path(
        _prompt_text(
            "Output process workbook",
            default=str(_default_extract_output_path(input_path)),
        )
    )
    min_group_size = _prompt_int("Minimum variants for a reusable template", default=2)
    use_existing_targets = (
        _prompt_yes_no("Use existing target column as template suggestions", default=True)
        if target_col is not None
        else False
    )

    stats = generate_workbook(
        input_path,
        output_path,
        source_col=source_col,
        target_col=target_col,
        tm_workbook=tm_workbook,
        min_group_size=min_group_size,
        use_existing_targets=use_existing_targets,
    )
    _print_stats(output_path, stats)
    return 0


def _interactive_fill() -> int:
    input_path = _user_path(_prompt_text("Original source Excel path", required=True))
    template_workbook = _user_path(
        _prompt_text("Translated to_translate file path", required=True)
    )
    source_col = _prompt_text("Source column in original file", default="英語")
    target_col = _normalize_optional_column(
        _prompt_text("Target column to write/check", default="target")
    )
    mode = _prompt_text("Fill mode: report or target-column", default="report")
    output_path = _user_path(
        _prompt_text("Output filled workbook", default=str(_default_fill_output_path(input_path)))
    )
    min_group_size = _prompt_int("Minimum variants for a reusable template", default=2)

    if mode == "target-column":
        if target_col is None:
            raise ValueError("target-column mode needs a target column")
        stats = fill_target_column_workbook(
            input_path,
            output_path,
            source_col=source_col,
            target_col=target_col,
            template_workbook=template_workbook,
            min_group_size=min_group_size,
        )
    else:
        stats = generate_workbook(
            input_path,
            output_path,
            source_col=source_col,
            target_col=target_col,
            template_workbook=template_workbook,
            min_group_size=min_group_size,
            use_existing_targets=False,
        )
    _print_stats(output_path, stats)
    return 0


def _main_tm_extract(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Extract reusable TM pairs from completed source/target columns."
    )
    parser.add_argument("input", type=Path, help="Completed TM .xlsx file")
    parser.add_argument("-o", "--output", type=Path, help="Output TM pairs .xlsx file")
    parser.add_argument("--source-col", default="英語", help="Header name or 1-based index")
    parser.add_argument("--target-col", default="target", help="Header name or 1-based index")
    parser.add_argument("--min-group-size", type=int, default=2)

    args = parser.parse_args(argv)
    output = args.output or _default_tm_output_path(args.input)
    stats = generate_tm_pairs(
        args.input,
        output,
        source_col=args.source_col,
        target_col=args.target_col,
        min_group_size=args.min_group_size,
    )
    _print_tm_stats(output, stats)
    return 0


def _main_extract(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Extract reusable localization templates from an Excel source column."
    )
    parser.add_argument("input", type=Path, help="Input .xlsx file")
    parser.add_argument("-o", "--output", type=Path, help="Output .xlsx file")
    parser.add_argument("--source-col", default="英語", help="Header name or 1-based index")
    parser.add_argument("--target-col", default="target", help="Header name or 1-based index")
    parser.add_argument(
        "--tm",
        type=Path,
        help="TM pairs workbook used to prefill matching translation units",
    )
    parser.add_argument(
        "--example",
        action="append",
        default=[],
        help='One source=target example, e.g. "VIP10 Paid Pack=VIP10pack"',
    )
    parser.add_argument("--min-group-size", type=int, default=2)
    parser.add_argument(
        "--no-existing-targets",
        action="store_true",
        help="Do not infer target templates from the existing target column",
    )

    args = parser.parse_args(argv)
    output = args.output or _default_extract_output_path(args.input)
    stats = generate_workbook(
        args.input,
        output,
        source_col=args.source_col,
        target_col=_normalize_optional_column(args.target_col),
        tm_workbook=args.tm,
        examples=_parse_examples(args.example),
        min_group_size=args.min_group_size,
        use_existing_targets=not args.no_existing_targets,
    )
    _print_stats(output, stats)
    return 0


def _main_fill(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Fill a source workbook from a translated template pack."
    )
    parser.add_argument("input", type=Path, help="Original input .xlsx file")
    parser.add_argument(
        "--templates",
        required=True,
        type=Path,
        help="Translated template pack workbook",
    )
    parser.add_argument("-o", "--output", type=Path, help="Output .xlsx file")
    parser.add_argument("--source-col", default="英語", help="Header name or 1-based index")
    parser.add_argument("--target-col", default="target", help="Header name or 1-based index")
    parser.add_argument("--min-group-size", type=int, default=2)
    parser.add_argument(
        "--mode",
        choices=["report", "target-column"],
        default="report",
        help="report creates analysis sheets; target-column writes auto targets into the target column of an output copy",
    )

    args = parser.parse_args(argv)
    output = args.output or _default_fill_output_path(args.input)
    target_col = _normalize_optional_column(args.target_col)
    if args.mode == "target-column":
        if target_col is None:
            raise ValueError("target-column mode needs a target column")
        stats = fill_target_column_workbook(
            args.input,
            output,
            source_col=args.source_col,
            target_col=target_col,
            template_workbook=args.templates,
            min_group_size=args.min_group_size,
        )
    else:
        stats = generate_workbook(
            args.input,
            output,
            source_col=args.source_col,
            target_col=target_col,
            template_workbook=args.templates,
            min_group_size=args.min_group_size,
            use_existing_targets=False,
        )
    _print_stats(output, stats)
    return 0


def _main_legacy(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Extract or fill reusable localization templates from an Excel source column."
    )
    parser.add_argument("input", type=Path, help="Input .xlsx file")
    parser.add_argument("-o", "--output", type=Path, help="Output .xlsx file")
    parser.add_argument("--source-col", default="英語", help="Header name or 1-based index")
    parser.add_argument("--target-col", default="target", help="Header name or 1-based index")
    parser.add_argument(
        "--templates",
        type=Path,
        help="Previous output workbook whose template_review sheet has target_template filled",
    )
    parser.add_argument(
        "--example",
        action="append",
        default=[],
        help='One source=target example, e.g. "VIP10 Paid Pack=VIP10pack"',
    )
    parser.add_argument("--min-group-size", type=int, default=2)
    parser.add_argument(
        "--no-existing-targets",
        action="store_true",
        help="Do not infer target templates from the existing target column",
    )

    args = parser.parse_args(argv)
    output = args.output or _default_legacy_output_path(args.input)
    stats = generate_workbook(
        args.input,
        output,
        source_col=args.source_col,
        target_col=_normalize_optional_column(args.target_col),
        examples=_parse_examples(args.example),
        template_workbook=args.templates,
        min_group_size=args.min_group_size,
        use_existing_targets=not args.no_existing_targets,
    )
    _print_stats(output, stats)
    return 0


def _prompt_text(prompt: str, *, default: str | None = None, required: bool = False) -> str:
    suffix = f" [{default}]" if default is not None else ""
    while True:
        value = input(f"{prompt}{suffix}: ").strip()
        if value:
            return value
        if default is not None:
            return default
        if not required:
            return ""
        print("This value is required.")


def _prompt_int(prompt: str, *, default: int) -> int:
    while True:
        raw = _prompt_text(prompt, default=str(default))
        try:
            value = int(raw)
        except ValueError:
            print("Please enter a number.")
            continue
        if value < 1:
            print("Please enter a number greater than 0.")
            continue
        return value


def _prompt_yes_no(prompt: str, *, default: bool) -> bool:
    default_text = "Y/n" if default else "y/N"
    while True:
        raw = input(f"{prompt} [{default_text}]: ").strip().lower()
        if not raw:
            return default
        if raw in {"y", "yes"}:
            return True
        if raw in {"n", "no"}:
            return False
        print("Please answer y or n.")


def _user_path(value: str | Path) -> Path:
    text = str(value).strip()
    if len(text) >= 2 and text[0] == text[-1] and text[0] in {"'", '"'}:
        text = text[1:-1].strip()
    return Path(text).expanduser()


def _normalize_optional_column(value: str | int | None) -> str | int | None:
    if value is None:
        return None
    if isinstance(value, int):
        return value
    normalized = value.strip()
    if normalized.lower() in {"", "-", "none", "no", "skip"}:
        return None
    return normalized


def _print_stats(output: Path, stats: dict[str, int]) -> None:
    print(f"Wrote: {output}")
    if "to_translate_path" in stats:
        print(f"To-translate workbook: {stats['to_translate_path']}")
    print(f"Units to translate: {stats['new_translation_unit_count']}")
    print(f"Source rows to translate: {stats['new_source_segment_count']}")
    print(f"Already filled units: {stats['prefilled_translation_unit_count']}")
    print(f"Already filled source rows: {stats['autofilled_count']}")
    print(f"Total translation units: {stats['translation_unit_count']}")
    print(f"Total source rows: {stats['row_count']}")


def _print_tm_stats(output: Path, stats: dict[str, int]) -> None:
    print(f"Wrote: {output}")
    print(f"TM source segments: {stats['row_count']}")
    print(f"Unique source segments: {stats['unique_source_segments']}")
    print(f"Duplicate source segments: {stats['duplicate_source_segments']}")
    print(f"TM pairs: {stats['tm_pair_count']}")
    print(f"Template pairs: {stats['template_pair_count']}")
    print(f"Segment pairs: {stats['segment_pair_count']}")


def _print_top_level_help() -> None:
    print("Localization Workflow")
    print()
    print("Interactive:")
    print("  python template_demo.py")
    print("  python template_demo.py interactive")
    print()
    print("Steps:")
    print("  1) Build TM from completed Excel")
    print("  2) Prepare translator file for new source")
    print("  3) Fill source from translated file")
    print()
    print("Commands:")
    print("  python template_demo.py tm-extract COMPLETED_TM.xlsx [options]")
    print("  python template_demo.py extract SOURCE.xlsx [options]")
    print("  python template_demo.py fill SOURCE.xlsx --templates TEMPLATE_PACK.xlsx [options]")
    print()
    print("Legacy:")
    print("  python template_demo.py SOURCE.xlsx [options]")


if __name__ == "__main__":
    raise SystemExit(main())
