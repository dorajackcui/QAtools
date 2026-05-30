#!/usr/bin/env python3
"""Check whether Excel target cells contain Chinese characters."""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter


PROBLEM_SHEET_NAME = "中文检查问题"
RESULT_HEADER = "中文检查"
CHINESE_MARKER = "含中文"
CHINESE_PATTERN = re.compile(
    r"[\u3400-\u4DBF\u4E00-\u9FFF\uF900-\uFAFF"
    r"\u3000-\u303F\uFF01-\uFF0F\uFF1A-\uFF20\uFF3B-\uFF40\uFF5B-\uFF65"
    r"\u00B7\u2014\u2018\u2019\u201C\u201D\u2026]"
)


@dataclass(frozen=True)
class CheckSummary:
    output_path: Path
    worksheet_title: str
    target_column: str
    result_column: str
    start_row: int
    processed_count: int
    matched_count: int


def normalize_column(column_name: str) -> str:
    normalized = column_name.strip().upper()
    column_index_from_string(normalized)
    return normalized


def build_default_output_path(input_file: str | Path) -> Path:
    input_path = Path(input_file).expanduser().resolve()
    return input_path.with_name(f"{input_path.stem}_chinese_target_checked{input_path.suffix}")


def build_default_result_column(target_column: str) -> str:
    return get_column_letter(column_index_from_string(target_column) + 1)


def contains_chinese(value: object) -> bool:
    return isinstance(value, str) and bool(CHINESE_PATTERN.search(value))


def extract_chinese_characters(value: object) -> str:
    if not isinstance(value, str):
        return ""
    return "".join(CHINESE_PATTERN.findall(value))


def process_excel(
    input_file: str | Path,
    target_column: str,
    result_column: str | None = None,
    sheet: str | None = None,
    start_row: int = 2,
    output_file: str | Path | None = None,
) -> CheckSummary:
    input_path = Path(input_file).expanduser().resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"输入文件不存在: {input_path}")
    if start_row < 1:
        raise ValueError("开始行必须大于等于 1。")

    target_column = normalize_column(target_column)
    normalized_result_column = (
        normalize_column(result_column)
        if result_column
        else build_default_result_column(target_column)
    )
    output_path = (
        Path(output_file).expanduser().resolve()
        if output_file
        else input_path
    )

    workbook = load_workbook(input_path)
    worksheet = workbook[sheet] if sheet else workbook.active
    if PROBLEM_SHEET_NAME in workbook.sheetnames and worksheet.title != PROBLEM_SHEET_NAME:
        del workbook[PROBLEM_SHEET_NAME]
    should_insert_result_column = (
        result_column is None
        and worksheet[f"{normalized_result_column}1"].value != RESULT_HEADER
    )
    if should_insert_result_column:
        worksheet.insert_cols(column_index_from_string(normalized_result_column))
    worksheet[f"{normalized_result_column}1"] = RESULT_HEADER

    processed_count = 0
    matched_count = 0
    for row_index in range(start_row, worksheet.max_row + 1):
        target_value = worksheet[f"{target_column}{row_index}"].value
        chinese_characters = extract_chinese_characters(target_value)
        has_chinese = bool(chinese_characters)
        worksheet[f"{normalized_result_column}{row_index}"] = CHINESE_MARKER if has_chinese else None
        processed_count += 1
        if has_chinese:
            matched_count += 1

    workbook.save(output_path)
    return CheckSummary(
        output_path=output_path,
        worksheet_title=worksheet.title,
        target_column=target_column,
        result_column=normalized_result_column,
        start_row=start_row,
        processed_count=processed_count,
        matched_count=matched_count,
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="检查 Excel target 列是否包含中文字符或中文/全角标点，并在结果列标记。"
    )
    parser.add_argument("input_file", nargs="?", help="输入 Excel 文件路径，例如 input.xlsx")
    parser.add_argument("-s", "--sheet", help="工作表名称，不填则默认处理当前活动工作表")
    parser.add_argument("-t", "--target-column", help="需要检查的 target 列，例如 B")
    parser.add_argument(
        "-r",
        "--result-column",
        help="可选，写入中文检查标记的结果列；不填则默认在 target 右侧新增一列",
    )
    parser.add_argument("--start-row", type=int, default=None, help="开始处理的行号，默认 2")
    parser.add_argument(
        "-o",
        "--output",
        help="输出 Excel 文件路径；不填则直接修改原文件",
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
        args.result_column = input("请输入结果列（直接回车则在 target 右侧新增一列）: ").strip().upper() or None
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
    print(f"结果列: {summary.result_column}")
    print(f"开始行: {summary.start_row}")
    print(f"处理行数: {summary.processed_count}")
    print(f"含中文行数: {summary.matched_count}")
    print(f"输出文件: {summary.output_path}")


if __name__ == "__main__":
    main()
