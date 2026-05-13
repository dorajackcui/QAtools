from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Iterable

from .errors import ConfigError, PhraseLoomError
from .entity_workflow import (
    default_entity_filled_pack_output_path,
    default_entity_memory_output_path,
    default_entity_merged_todo_output_path,
    default_entity_pack_output_path,
    extract_entity_memory_workbook,
    extract_entity_tm_workbook,
    fill_entity_pack_workbook,
    fill_entity_workbook,
    merge_entity_pack_workbook,
    merge_entity_workbooks,
    prepare_entity_pack_workbook,
    prefill_entity_workbook,
    split_entity_workbook,
)
from .excel_io import (
    _default_extract_output_path,
    _default_fill_output_path,
    _default_legacy_output_path,
    _default_tm_output_path,
)
from .interactive import (
    _normalize_optional_column,
    run_interactive,
)
from .workflow import (
    fill_target_column_workbook,
    generate_tm_pairs,
    generate_workbook,
)


def _parse_examples(raw_examples: Iterable[str]) -> list[tuple[str, str]]:
    examples: list[tuple[str, str]] = []
    for raw in raw_examples:
        if "=" not in raw:
            raise ConfigError(f"Example must look like SOURCE=TARGET: {raw!r}")
        source, target = raw.split("=", 1)
        examples.append((source.strip(), target.strip()))
    return examples


def main(argv: list[str] | None = None) -> int:
    try:
        return _dispatch(argv)
    except PhraseLoomError as error:
        print(str(error), file=sys.stderr)
        return 1


def _dispatch(argv: list[str] | None = None) -> int:
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
    if argv[0] == "entity-tm":
        return _main_entity_tm(argv[1:])
    if argv[0] == "entity-prepare":
        return _main_entity_prepare(argv[1:])
    if argv[0] == "entity-split":
        return _main_entity_split(argv[1:])
    if argv[0] == "entity-prefill":
        return _main_entity_prefill(argv[1:])
    if argv[0] == "entity-extract-tm":
        return _main_entity_extract_tm(argv[1:])
    if argv[0] == "entity-fill":
        return _main_entity_fill(argv[1:])
    if argv[0] == "entity-fill-pack":
        return _main_entity_fill_pack(argv[1:])
    if argv[0] == "entity-merge-pack":
        return _main_entity_merge_pack(argv[1:])
    if argv[0] == "entity-merge":
        return _main_entity_merge(argv[1:])
    return _main_legacy(argv)


def _main_tm_extract(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Extract reusable TM pairs from completed source/target columns."
    )
    parser.add_argument("input", type=Path, help="Completed TM .xlsx file")
    parser.add_argument("-o", "--output", type=Path, help="Output TM pairs .xlsx file")
    parser.add_argument("--source-col", default="source", help="Header name or 1-based index")
    parser.add_argument("--target-col", default="target", help="Header name or 1-based index")
    parser.add_argument("--min-group-size", type=int, default=2)
    parser.add_argument(
        "--tag-config",
        type=Path,
        help="TOML file defining which tag-like spans are protected",
    )

    args = parser.parse_args(argv)
    output = args.output or _default_tm_output_path(args.input)
    stats = generate_tm_pairs(
        args.input,
        output,
        source_col=args.source_col,
        target_col=args.target_col,
        min_group_size=args.min_group_size,
        tag_config=args.tag_config,
    )
    _print_tm_stats(output, stats)
    return 0


def _main_extract(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Extract reusable localization templates from an Excel source column."
    )
    parser.add_argument("input", type=Path, help="Input .xlsx file")
    parser.add_argument("-o", "--output", type=Path, help="Output .xlsx file")
    parser.add_argument("--source-col", default="source", help="Header name or 1-based index")
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
        "--tag-config",
        type=Path,
        help="TOML file defining which tag-like spans are protected",
    )
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
        tag_config=args.tag_config,
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
    parser.add_argument("--source-col", default="source", help="Header name or 1-based index")
    parser.add_argument("--target-col", default="target", help="Header name or 1-based index")
    parser.add_argument("--min-group-size", type=int, default=2)
    parser.add_argument(
        "--tag-config",
        type=Path,
        help="TOML file defining which tag-like spans are protected",
    )
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
            raise ConfigError("target-column mode needs a target column")
        stats = fill_target_column_workbook(
            args.input,
            output,
            source_col=args.source_col,
            target_col=target_col,
            template_workbook=args.templates,
            min_group_size=args.min_group_size,
            tag_config=args.tag_config,
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
            tag_config=args.tag_config,
        )
    _print_stats(output, stats)
    return 0


def _main_entity_tm(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Build reusable entity memory from a TM reusable-units workbook."
    )
    parser.add_argument("input", type=Path, help="TM reusable units .xlsx file")
    parser.add_argument("-o", "--output", type=Path, help="Entity memory output .xlsx")
    parser.add_argument("--min-group-size", type=int, default=3)
    args = parser.parse_args(argv)
    output = args.output or default_entity_memory_output_path(args.input)
    stats = extract_entity_memory_workbook(
        args.input,
        output,
        min_group_size=args.min_group_size,
    )
    _print_entity_extract_tm_stats(stats)
    return 0


def _main_entity_prepare(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Prepare one source entity pack with related and non-related units."
    )
    parser.add_argument("input", type=Path, help="Translator todo .xlsx file")
    parser.add_argument("-o", "--output", type=Path, help="Entity pack output .xlsx")
    parser.add_argument("--tm", type=Path, help="Entity memory workbook used to prefill the pack")
    parser.add_argument("--min-group-size", type=int, default=3)
    args = parser.parse_args(argv)
    output = args.output or default_entity_pack_output_path(args.input)
    stats = prepare_entity_pack_workbook(
        args.input,
        output,
        tm_path=args.tm,
        min_group_size=args.min_group_size,
    )
    _print_entity_prepare_stats(stats)
    return 0


def _main_entity_split(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Split a preprocessed todo workbook into entity and non-entity workbooks."
    )
    parser.add_argument("input", type=Path, help="Translator todo .xlsx file")
    parser.add_argument("--entity-output", type=Path, help="Entity-related output .xlsx")
    parser.add_argument(
        "--non-entity-output",
        type=Path,
        help="Non-entity output .xlsx",
    )
    parser.add_argument("--min-group-size", type=int, default=3)
    args = parser.parse_args(argv)
    entity_output = args.entity_output or args.input.with_name(
        f"{args.input.stem}_entity_related.xlsx"
    )
    non_entity_output = args.non_entity_output or args.input.with_name(
        f"{args.input.stem}_not_entity_related.xlsx"
    )
    stats = split_entity_workbook(
        args.input,
        entity_output,
        non_entity_output,
        min_group_size=args.min_group_size,
    )
    _print_entity_split_stats(stats)
    return 0


def _main_entity_prefill(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Prefill entity structures and terms from an entity TM workbook."
    )
    parser.add_argument("input", type=Path, help="Entity-related workbook")
    parser.add_argument("--tm", required=True, type=Path, help="Entity TM workbook")
    parser.add_argument("-o", "--output", type=Path, help="Prefilled entity workbook")
    args = parser.parse_args(argv)
    output = args.output or args.input.with_name(f"{args.input.stem}_prefilled.xlsx")
    stats = prefill_entity_workbook(args.input, args.tm, output)
    _print_entity_prefill_stats(stats)
    return 0


def _main_entity_extract_tm(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Extract reusable entity structures and terms from a TM pairs workbook."
    )
    parser.add_argument("input", type=Path, help="TM reusable units .xlsx file")
    parser.add_argument("-o", "--output", type=Path, help="Entity TM output workbook")
    parser.add_argument("--min-group-size", type=int, default=3)
    args = parser.parse_args(argv)
    output = args.output or args.input.with_name(f"{args.input.stem}_entity_tm.xlsx")
    stats = extract_entity_tm_workbook(
        args.input,
        output,
        min_group_size=args.min_group_size,
    )
    _print_entity_extract_tm_stats(stats)
    return 0


def _main_entity_fill(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Fill ready entity rows back into an entity-related todo workbook."
    )
    parser.add_argument("input", type=Path, help="Entity-related workbook")
    parser.add_argument("-o", "--output", type=Path, help="Filled entity workbook")
    args = parser.parse_args(argv)
    output = args.output or args.input.with_name(f"{args.input.stem}_filled.xlsx")
    stats = fill_entity_workbook(args.input, output)
    _print_entity_fill_stats(stats)
    return 0


def _main_entity_fill_pack(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Fill completed entity structures and terms back into related_units."
    )
    parser.add_argument("input", type=Path, help="Source entity pack .xlsx file")
    output_group = parser.add_mutually_exclusive_group()
    output_group.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Filled entity pack output .xlsx",
    )
    output_group.add_argument(
        "--in-place",
        action="store_true",
        help="Update the input pack instead of writing a new file",
    )
    args = parser.parse_args(argv)
    output = args.input if args.in_place else (
        args.output or default_entity_filled_pack_output_path(args.input)
    )
    stats = fill_entity_pack_workbook(args.input, output)
    _print_entity_fill_stats(stats)
    return 0


def _main_entity_merge_pack(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Merge related_units and non_related_units into a normal translator todo."
    )
    parser.add_argument("input", type=Path, help="Filled source entity pack .xlsx file")
    parser.add_argument("-o", "--output", type=Path, help="Merged todo output .xlsx")
    args = parser.parse_args(argv)
    output = args.output or default_entity_merged_todo_output_path(args.input)
    stats = merge_entity_pack_workbook(args.input, output)
    _print_entity_merge_stats(stats)
    return 0


def _main_entity_merge(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Merge filled entity rows and non-entity rows into a todo workbook."
    )
    parser.add_argument("--entity", required=True, type=Path, help="Filled entity workbook")
    parser.add_argument(
        "--non-entity",
        required=True,
        type=Path,
        help="Non-entity workbook",
    )
    parser.add_argument("-o", "--output", type=Path, help="Merged todo workbook")
    args = parser.parse_args(argv)
    output = args.output or args.entity.with_name(f"{args.entity.stem}_merged_todo.xlsx")
    stats = merge_entity_workbooks(args.entity, args.non_entity, output)
    _print_entity_merge_stats(stats)
    return 0


def _main_legacy(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Extract or fill reusable localization templates from an Excel source column."
    )
    parser.add_argument("input", type=Path, help="Input .xlsx file")
    parser.add_argument("-o", "--output", type=Path, help="Output .xlsx file")
    parser.add_argument("--source-col", default="source", help="Header name or 1-based index")
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
        "--tag-config",
        type=Path,
        help="TOML file defining which tag-like spans are protected",
    )
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
        tag_config=args.tag_config,
    )
    _print_stats(output, stats)
    return 0


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


def _print_entity_split_stats(stats: dict[str, int | str]) -> None:
    print(f"Entity workbook: {stats['entity_output_path']}")
    print(f"Non-entity workbook: {stats['non_entity_output_path']}")
    print(f"Entity units: {stats['entity_unit_count']}")
    print(f"Non-entity units: {stats['non_entity_unit_count']}")
    print(f"Entity structures: {stats['entity_structure_count']}")
    print(f"Entity terms: {stats['entity_term_count']}")


def _print_entity_prefill_stats(stats: dict[str, int | str]) -> None:
    print(f"Wrote: {stats['output_path']}")
    print(f"Prefilled structures: {stats['prefilled_structure_count']}")
    print(f"Prefilled terms: {stats['prefilled_term_count']}")


def _print_entity_prepare_stats(stats: dict[str, int | str]) -> None:
    print(f"Wrote: {stats['output_path']}")
    print(f"Related units: {stats['related_unit_count']}")
    print(f"Non-related units: {stats['non_related_unit_count']}")
    print(f"Entity structures: {stats['entity_structure_count']}")
    print(f"Entity terms: {stats['entity_term_count']}")
    print(f"Prefilled structures: {stats['prefilled_structure_count']}")
    print(f"Prefilled terms: {stats['prefilled_term_count']}")


def _print_entity_extract_tm_stats(stats: dict[str, int | str]) -> None:
    print(f"Wrote: {stats['output_path']}")
    print(f"Entity structures: {stats['entity_structure_count']}")
    print(f"Entity terms: {stats['entity_term_count']}")


def _print_entity_fill_stats(stats: dict[str, int | str]) -> None:
    print(f"Wrote: {stats['output_path']}")
    print(f"Filled entity units: {stats['filled_entity_unit_count']}")


def _print_entity_merge_stats(stats: dict[str, int | str]) -> None:
    print(f"Wrote: {stats['output_path']}")
    print(f"Merged units: {stats['merged_unit_count']}")


def _print_top_level_help() -> None:
    print("Localization Workflow")
    print()
    print("Interactive:")
    print("  phraseloom")
    print("  phraseloom interactive")
    print()
    print("Steps:")
    print("  1) Build TM from completed Excel")
    print("  2) Prepare translator file for new source")
    print("  3) Fill source from translated file")
    print()
    print("Commands:")
    print("  phraseloom tm-extract COMPLETED_TM.xlsx [options]")
    print("  phraseloom extract SOURCE.xlsx [options]")
    print("  phraseloom fill SOURCE.xlsx --templates TEMPLATE_PACK.xlsx [options]")
    print()
    print("Entity workflow:")
    print("  phraseloom entity-tm TM_REUSABLE_UNITS.xlsx [options]")
    print("  phraseloom entity-prepare TRANSLATOR_WORKBOOK.xlsx [options]")
    print("  phraseloom entity-fill-pack ENTITY_PACK.xlsx [options]")
    print("  phraseloom entity-merge-pack FILLED_ENTITY_PACK.xlsx [options]")
    print()
    print("Advanced entity commands:")
    print("  phraseloom entity-split TRANSLATOR_WORKBOOK.xlsx [options]")
    print("  phraseloom entity-extract-tm TM_REUSABLE_UNITS.xlsx [options]")
    print("  phraseloom entity-prefill ENTITY.xlsx --tm ENTITY_TM.xlsx [options]")
    print("  phraseloom entity-fill ENTITY.xlsx [options]")
    print("  phraseloom entity-merge --entity ENTITY.xlsx --non-entity NON_ENTITY.xlsx [options]")
    print()
    print("Legacy:")
    print("  python template_demo.py tm-extract COMPLETED_TM.xlsx [options]")
    print("  python template_demo.py extract SOURCE.xlsx [options]")
    print("  python template_demo.py fill SOURCE.xlsx --templates TEMPLATE_PACK.xlsx [options]")
    print("  python template_demo.py SOURCE.xlsx [options]")


__all__ = [
    "main",
    "_parse_examples",
    "_main_tm_extract",
    "_main_extract",
    "_main_entity_extract_tm",
    "_main_entity_fill",
    "_main_entity_fill_pack",
    "_main_entity_merge",
    "_main_entity_merge_pack",
    "_main_entity_prefill",
    "_main_entity_prepare",
    "_main_entity_split",
    "_main_entity_tm",
    "_main_fill",
    "_main_legacy",
    "_print_entity_prepare_stats",
    "_print_stats",
    "_print_tm_stats",
    "_print_top_level_help",
]


if __name__ == "__main__":
    raise SystemExit(main())
