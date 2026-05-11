from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Iterable

from .errors import ConfigError, PhraseLoomError
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
    return _main_legacy(argv)


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
            raise ConfigError("target-column mode needs a target column")
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
    "_main_fill",
    "_main_legacy",
    "_print_stats",
    "_print_tm_stats",
    "_print_top_level_help",
]


if __name__ == "__main__":
    raise SystemExit(main())
