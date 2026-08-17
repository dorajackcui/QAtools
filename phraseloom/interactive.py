from __future__ import annotations

from pathlib import Path

from .strings_workflow import export_strings_workbook, restore_strings_workbook
from .workbook_io import read_headers


def run_interactive() -> int:
    print("PhraseLoom Strings Workflow")
    print()
    print("1) Export untranslated Strings")
    print("2) Restore translated Strings")
    print("q) Quit")

    action = _prompt_text("Choose step", default="1").lower()
    if action in {"q", "quit", "exit"}:
        print("Bye.")
        return 0
    if action in {"1", "export", "e"}:
        return _export()
    if action in {"2", "restore", "r"}:
        return _restore()
    print(f"Unknown step: {action}")
    return 2


def _export() -> int:
    input_path = _user_path(_prompt_text("Source Excel path", required=True))
    headers = read_headers(input_path)
    source_column = _find_header(headers, "source") or _prompt_text(
        f"Source column (available: {', '.join(headers)})",
        required=True,
    )
    target_column = _find_header(headers, "target") or "target"
    split_lines = _prompt_yes_no("Split multiline Source cells", default=True)
    group_similar = _prompt_yes_no("Group similar cleaned strings", default=False)
    stats = export_strings_workbook(
        input_path,
        source_col=source_column,
        target_col=target_column,
        split_lines=split_lines,
        group_similar=group_similar,
    )
    _print_export_stats(stats)
    return 0


def _restore() -> int:
    package_path = _user_path(
        _prompt_text("Translated Strings workbook", required=True)
    )
    stats = restore_strings_workbook(package_path)
    _print_restore_stats(stats)
    return 0


def _find_header(headers: list[str], wanted: str) -> str | None:
    wanted_lower = wanted.lower()
    return next(
        (header for header in headers if header.lower() == wanted_lower),
        None,
    )


def _prompt_text(
    prompt: str,
    *,
    default: str | None = None,
    required: bool = False,
) -> str:
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


def _prompt_yes_no(prompt: str, *, default: bool) -> bool:
    default_text = "Y/n" if default else "y/N"
    while True:
        value = input(f"{prompt} [{default_text}]: ").strip().lower()
        if not value:
            return default
        if value in {"y", "yes"}:
            return True
        if value in {"n", "no"}:
            return False
        print("Please answer y or n.")


def _user_path(value: str | Path) -> Path:
    text = str(value).strip()
    if len(text) >= 2 and text[0] == text[-1] and text[0] in {"'", '"'}:
        text = text[1:-1].strip()
    return Path(text).expanduser()


def _print_export_stats(stats: dict[str, int | str]) -> None:
    print(f"Strings workbook: {stats['output_path']}")
    print(f"Strings to translate: {stats['string_count']}")
    print(
        f"Similar groups: {stats['group_count']}"
        if stats["grouping_enabled"]
        else "Similar grouping: off"
    )
    print(
        "Multiline Source splitting: on"
        if stats["line_splitting_enabled"]
        else "Multiline Source splitting: off"
    )
    print(f"Existing targets skipped: {stats['completed_row_count']}")


def _print_restore_stats(stats: dict[str, int | str]) -> None:
    print(f"Translated workbook: {stats['output_path']}")
    print(f"Restored source rows: {stats['restored_row_count']}")
    print(f"Issues: {stats['issue_count']}")
    if "audit_output_path" in stats:
        print(f"Review workbook: {stats['audit_output_path']}")


__all__ = ["run_interactive"]
