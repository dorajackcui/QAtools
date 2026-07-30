from __future__ import annotations

import argparse
from pathlib import Path
import sys

from .errors import ConfigError, PhraseLoomError
from .interactive import run_interactive
from .strings_workflow import export_strings_workbook, restore_strings_workbook


def main(argv: list[str] | None = None) -> int:
    try:
        return _dispatch(argv)
    except PhraseLoomError as error:
        print(str(error), file=sys.stderr)
        return 1


def _dispatch(argv: list[str] | None = None) -> int:
    arguments = list(sys.argv[1:] if argv is None else argv)
    if not arguments:
        return run_interactive()

    command = arguments[0]
    if command in {"-h", "--help"}:
        _print_top_level_help()
        return 0
    if command == "gui":
        from .gui import main as gui_main

        return gui_main()
    if command == "interactive":
        return run_interactive()
    if command == "export":
        return _export(arguments[1:])
    if command == "restore":
        return _restore(arguments[1:])
    raise ConfigError(
        f"Unknown command: {command!r}. Use 'phraseloom --help' for available commands."
    )


def _export(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        prog="phraseloom export",
        description="Export clean, deduplicated untranslated Strings.",
    )
    parser.add_argument("input", type=Path, help="Source .xlsx file")
    parser.add_argument("-o", "--output", type=Path, help="Strings output .xlsx")
    parser.add_argument("--source-col", default="source")
    parser.add_argument("--target-col", default="target")
    parser.add_argument(
        "--context-col",
        help="Optional context header or 1-based index; auto-detects 'context'",
    )
    parser.add_argument(
        "--tag-config",
        type=Path,
        help="TOML file defining protected tags and placeholders",
    )
    parser.add_argument(
        "--group-similar",
        action="store_true",
        help="Place structurally similar cleaned strings together at the end",
    )
    args = parser.parse_args(argv)
    stats = export_strings_workbook(
        args.input,
        args.output,
        source_col=args.source_col,
        target_col=args.target_col,
        context_col=args.context_col,
        group_similar=args.group_similar,
        tag_config=args.tag_config,
    )
    _print_export_stats(stats)
    return 0


def _restore(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        prog="phraseloom restore",
        description="Restore translated Strings into the embedded source workbook.",
    )
    parser.add_argument("input", type=Path, help="Translated Strings .xlsx file")
    parser.add_argument("-o", "--output", type=Path, help="Translated output .xlsx")
    args = parser.parse_args(argv)
    stats = restore_strings_workbook(args.input, args.output)
    _print_restore_stats(stats)
    return 0


def _print_export_stats(stats: dict[str, int | str]) -> None:
    print(f"Strings workbook: {stats['output_path']}")
    print(f"Strings to translate: {stats['string_count']}")
    print(f"Pending source rows: {stats['pending_row_count']}")
    print(f"Rows consolidated by cleaning: {stats['duplicate_row_count']}")
    print(
        f"Similar groups: {stats['group_count']}"
        if stats["grouping_enabled"]
        else "Similar grouping: off"
    )
    print(f"Existing targets skipped: {stats['completed_row_count']}")
    print(
        "Non-translatable rows auto-completed: "
        f"{stats['auto_completed_row_count']}"
    )


def _print_restore_stats(stats: dict[str, int | str]) -> None:
    print(f"Translated workbook: {stats['output_path']}")
    print(f"Restored source rows: {stats['restored_row_count']}")
    print(f"Issues: {stats['issue_count']}")
    if "audit_output_path" in stats:
        print(f"Review workbook: {stats['audit_output_path']}")


def _print_top_level_help() -> None:
    print("PhraseLoom Strings Workflow")
    print()
    print("Steps:")
    print("  1) Export untranslated Strings")
    print("  2) Restore translated Strings")
    print()
    print("Commands:")
    print("  phraseloom gui")
    print("  phraseloom export SOURCE.xlsx [options]")
    print("  phraseloom restore SOURCE_STRINGS.xlsx [options]")
    print("  phraseloom interactive")


__all__ = ["main", "_dispatch", "_print_top_level_help"]


if __name__ == "__main__":
    raise SystemExit(main())
