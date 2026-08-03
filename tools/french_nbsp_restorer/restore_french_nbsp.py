#!/usr/bin/env python3
"""Restore French non-breaking spaces in Excel target text."""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from pathlib import Path

if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from openpyxl.utils import column_index_from_string

from tools.excel_output import (
    build_prefixed_output_path,
    find_last_value_row,
    load_workbook_for_editing,
)


NBSP = "\u00a0"
FRENCH_SPACING_CHARS = " \t\u00a0\u202f"
FRENCH_PRECEDING_NBSP_CHARS = ";:?!%"
OPEN_GUILLEMET_PATTERN = re.compile(rf"«[{FRENCH_SPACING_CHARS}]*")
CLOSE_GUILLEMET_PATTERN = re.compile(rf"[{FRENCH_SPACING_CHARS}]*»")
URL_SCHEME_PATTERN = re.compile(r"[A-Za-z][A-Za-z0-9+.-]*$")
URL_TOKEN_PATTERN = re.compile(r"[A-Za-z][A-Za-z0-9+.-]*://")


@dataclass(frozen=True)
class RestoreSummary:
    output_path: Path
    worksheet_title: str
    target_column: str
    result_column: str | None
    start_row: int
    processed_count: int
    changed_count: int


def normalize_column(column_name: str) -> str:
    normalized = column_name.strip().upper()
    column_index_from_string(normalized)
    return normalized


def build_default_output_path(input_file: str | Path) -> Path:
    return build_prefixed_output_path(input_file, "french_nbsp_restore_")


def _is_protected_colon(text: str, index: int) -> bool:
    previous_char = text[index - 1] if index > 0 else ""
    next_char = text[index + 1] if index + 1 < len(text) else ""
    if previous_char.isdigit() and next_char.isdigit():
        return True

    scheme_match = URL_SCHEME_PATTERN.search(text[:index])
    return bool(scheme_match and text[index + 1 : index + 3] == "//")


def _is_inside_url_token(text: str, index: int) -> bool:
    start = index
    while start > 0 and not text[start - 1].isspace():
        start -= 1

    end = index
    while end + 1 < len(text) and not text[end + 1].isspace():
        end += 1

    token = text[start : end + 1]
    token_index = index - start
    url_match = URL_TOKEN_PATTERN.search(token)
    return bool(url_match and token_index >= url_match.start())


def _restore_preceding_nbsp_spacing(text: str) -> str:
    result: list[str] = []
    for index, char in enumerate(text):
        if char not in FRENCH_PRECEDING_NBSP_CHARS:
            result.append(char)
            continue

        if _is_inside_url_token(text, index):
            result.append(char)
            continue

        if char == ":" and _is_protected_colon(text, index):
            result.append(char)
            continue

        while result and result[-1] in FRENCH_SPACING_CHARS:
            result.pop()
        if result and not result[-1].isspace():
            result.append(NBSP)
        result.append(char)

    return "".join(result)


def _restore_guillemet_spacing(text: str) -> str:
    text = OPEN_GUILLEMET_PATTERN.sub(f"«{NBSP}", text)
    return CLOSE_GUILLEMET_PATTERN.sub(f"{NBSP}»", text)


def restore_french_nbsp(value: object) -> object:
    if not isinstance(value, str):
        return value

    text = _restore_guillemet_spacing(value)
    return _restore_preceding_nbsp_spacing(text)


def process_excel(
    input_file: str | Path,
    target_column: str,
    result_column: str | None = None,
    sheet: str | None = None,
    start_row: int = 2,
    output_file: str | Path | None = None,
) -> RestoreSummary:
    input_path = Path(input_file).expanduser().resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"输入文件不存在: {input_path}")
    if start_row < 1:
        raise ValueError("开始行必须大于等于 1。")

    target_column = normalize_column(target_column)
    normalized_result_column = normalize_column(result_column) if result_column else None
    destination_column = normalized_result_column or target_column
    output_path = (
        Path(output_file).expanduser().resolve()
        if output_file
        else build_default_output_path(input_path)
    )

    workbook = load_workbook_for_editing(input_path)
    worksheet = workbook[sheet] if sheet else workbook.active
    last_row = find_last_value_row(
        worksheet,
        (target_column, destination_column),
        start_row=start_row,
    )

    processed_count = 0
    changed_count = 0
    for row_index in range(start_row, last_row + 1):
        original_value = worksheet[f"{target_column}{row_index}"].value
        restored_value = restore_french_nbsp(original_value)
        worksheet[f"{destination_column}{row_index}"] = restored_value
        processed_count += 1
        if restored_value != original_value:
            changed_count += 1

    workbook.save(output_path)
    return RestoreSummary(
        output_path=output_path,
        worksheet_title=worksheet.title,
        target_column=target_column,
        result_column=normalized_result_column,
        start_row=start_row,
        processed_count=processed_count,
        changed_count=changed_count,
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="恢复法语 target 文本中的 non-breaking space，并输出新的 Excel 文件。"
    )
    parser.add_argument("input_file", nargs="?", help="输入 Excel 文件路径，例如 input.xlsx")
    parser.add_argument(
        "-s",
        "--sheet",
        help="工作表名称，不填则默认处理当前活动工作表",
    )
    parser.add_argument(
        "-t",
        "--target-column",
        help="需要修复的 target 列，例如 B",
    )
    parser.add_argument(
        "-r",
        "--result-column",
        help="可选，写入修复后完整 target 的结果列；不填则直接修复 target 列",
    )
    parser.add_argument(
        "--start-row",
        type=int,
        default=None,
        help="开始处理的行号，默认 2",
    )
    parser.add_argument(
        "-o",
        "--output",
        help="输出 Excel 文件路径，默认生成 french_nbsp_restore_<原文件名>",
    )
    return parser.parse_args()


def prompt_if_missing(args: argparse.Namespace) -> argparse.Namespace:
    interactive_mode = sys.stdin.isatty()

    if not args.input_file and not interactive_mode:
        raise ValueError("缺少输入文件路径，请传入 input_file 参数。")
    if not args.input_file:
        args.input_file = input("请输入 Excel 文件路径: ").strip()
    if not args.sheet and interactive_mode and len(sys.argv) == 1:
        args.sheet = input("请输入工作表名称（直接回车使用当前活动工作表）: ").strip() or None
    if not args.target_column and not interactive_mode:
        raise ValueError("缺少 target 列，请使用 -t 或 --target-column 指定。")
    if not args.target_column:
        args.target_column = input("请输入 target 列（例如 B）: ").strip().upper()
    if not args.result_column and interactive_mode and len(sys.argv) == 1:
        args.result_column = input("请输入结果列（直接回车则修复 target 列）: ").strip().upper() or None
    if args.start_row is None:
        if interactive_mode and len(sys.argv) == 1:
            start_row_text = input("请输入开始处理的行号（默认 2）: ").strip()
            args.start_row = int(start_row_text) if start_row_text else 2
        else:
            args.start_row = 2
    return args


def main() -> None:
    args = prompt_if_missing(parse_args())
    summary = process_excel(
        input_file=args.input_file,
        target_column=args.target_column,
        result_column=args.result_column,
        sheet=args.sheet,
        start_row=args.start_row,
        output_file=args.output,
    )

    print("处理完成。")
    print(f"工作表: {summary.worksheet_title}")
    print(f"target 列: {summary.target_column}")
    print(f"结果列: {summary.result_column or '直接修复 target 列'}")
    print(f"开始行: {summary.start_row}")
    print(f"处理行数: {summary.processed_count}")
    print(f"修复行数: {summary.changed_count}")
    print(f"输出文件: {summary.output_path}")


if __name__ == "__main__":
    main()
