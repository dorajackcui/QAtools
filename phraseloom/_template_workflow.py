from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Iterable

from .excel_io import (
    _default_extract_output_path,
    _default_fill_output_path,
    _default_legacy_output_path,
    _default_tm_output_path,
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
