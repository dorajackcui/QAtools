#!/usr/bin/env python3
"""Validate source/target term pairs extracted from an Excel worksheet."""

from __future__ import annotations

import argparse
import re
import sys
from collections.abc import Iterable
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string


SUPPORTED_MARKS = ("【】", "[]", "<>")
TERM_SHEET_NAME = "术语表"
PROBLEM_SHEET_NAME = "问题列"
DEFAULT_MARK_STYLES = ("【】",)
MARK_PATTERNS: dict[str, tuple[re.Pattern[str], ...]] = {
    "【】": (re.compile(r"【([^【】]+)】"),),
    "[]": (
        re.compile(r"\[([^\[\]]+)\]"),
        re.compile(r"［([^［］]+)］"),
    ),
    "<>": (
        re.compile(r"<([^<>]+)>"),
        re.compile(r"＜([^＜＞]+)＞"),
    ),
}


def normalize_mark_styles(
    mark_styles: Iterable[str] | None = None,
    mark_style: str | None = None,
) -> tuple[str, ...]:
    if mark_styles is None:
        raw_mark_styles: list[str] = [mark_style] if mark_style else list(DEFAULT_MARK_STYLES)
    elif isinstance(mark_styles, str):
        raw_mark_styles = [mark_styles]
    else:
        raw_mark_styles = [style for style in mark_styles if style]

    invalid_mark_styles = [style for style in raw_mark_styles if style not in SUPPORTED_MARKS]
    if invalid_mark_styles:
        raise ValueError(f"不支持的 mark 类型: {'、'.join(invalid_mark_styles)}")

    normalized_mark_styles = tuple(style for style in SUPPORTED_MARKS if style in raw_mark_styles)
    if not normalized_mark_styles:
        raise ValueError("请至少选择一种 mark 类型。")
    return normalized_mark_styles


def extract_terms(
    text: object,
    mark_styles: Iterable[str] | None = None,
    mark_style: str | None = None,
) -> list[str]:
    if text is None:
        return []

    normalized_mark_styles = normalize_mark_styles(mark_styles=mark_styles, mark_style=mark_style)
    text_value = str(text)
    matches: list[tuple[int, int, str]] = []

    for current_mark_style in normalized_mark_styles:
        for pattern in MARK_PATTERNS[current_mark_style]:
            for match in pattern.finditer(text_value):
                matches.append((match.start(), match.end(), match.group(0)))

    matches.sort(key=lambda item: (item[0], item[1]))
    return [term for _, _, term in matches]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="从 Excel 的 source/target 两列提取术语对，并生成术语表和问题列。"
    )
    parser.add_argument("input_file", nargs="?", help="输入 Excel 文件路径，例如 input.xlsx")
    parser.add_argument(
        "-s",
        "--sheet",
        help="工作表名称，不填则默认处理当前活动工作表",
    )
    parser.add_argument(
        "-c",
        "--source-column",
        help="source 列，例如 A 或 C",
    )
    parser.add_argument(
        "-t",
        "--target-column",
        help="target 列，例如 B 或 D",
    )
    parser.add_argument(
        "--start-row",
        type=int,
        default=None,
        help="开始处理的行号，默认通过交互输入，留空时使用 2",
    )
    parser.add_argument(
        "--mark-style",
        action="append",
        choices=SUPPORTED_MARKS,
        default=None,
        help="术语包裹符号，可重复传入，例如 --mark-style [] --mark-style <>",
    )
    parser.add_argument(
        "-o",
        "--output",
        help="输出 Excel 文件路径，默认生成 <原文件名>_term_pairs.xlsx",
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
    if not args.source_column and not interactive_mode:
        raise ValueError("缺少 source 列，请使用 -c 或 --source-column 指定。")
    if not args.source_column:
        args.source_column = input("请输入 source 列（例如 A）: ").strip().upper()
    if not args.target_column and not interactive_mode:
        raise ValueError("缺少 target 列，请使用 -t 或 --target-column 指定。")
    if not args.target_column:
        args.target_column = input("请输入 target 列（例如 B）: ").strip().upper()
    if args.start_row is None:
        if interactive_mode and len(sys.argv) == 1:
            start_row_text = input("请输入开始处理的行号（默认 2）: ").strip()
            args.start_row = int(start_row_text) if start_row_text else 2
        else:
            args.start_row = 2
    if args.mark_style is None:
        if interactive_mode and len(sys.argv) == 1:
            mark_style_text = input(
                "请输入 mark 类型（可多选，逗号分隔，如 【】,[],<>；默认 【】）: "
            ).strip()
            if mark_style_text:
                mark_styles = [style.strip() for style in mark_style_text.split(",") if style.strip()]
                args.mark_style = normalize_mark_styles(mark_styles=mark_styles)
            else:
                args.mark_style = DEFAULT_MARK_STYLES
        else:
            args.mark_style = DEFAULT_MARK_STYLES
    else:
        args.mark_style = normalize_mark_styles(mark_styles=args.mark_style)
    return args


def normalize_column(column_name: str) -> str:
    normalized = column_name.strip().upper()
    column_index_from_string(normalized)
    return normalized


def build_default_output_path(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}_term_pairs{input_path.suffix}")


def rebuild_output_sheet(workbook, current_sheet_name: str, sheet_name: str):
    if current_sheet_name == sheet_name:
        raise ValueError(f"数据工作表名称不能为 {sheet_name}")
    if sheet_name in workbook.sheetnames:
        del workbook[sheet_name]
    return workbook.create_sheet(title=sheet_name)


def format_term_list(terms: list[str]) -> str:
    return "、".join(terms) if terms else "无"


def process_excel(
    input_file: str | Path,
    source_column: str,
    target_column: str,
    sheet: str | None = None,
    start_row: int = 2,
    mark_styles: Iterable[str] | None = None,
    mark_style: str | None = None,
    output_file: str | Path | None = None,
) -> tuple[str, str, str, Path, int, int]:
    input_path = Path(input_file).expanduser().resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"输入文件不存在: {input_path}")

    source_column = normalize_column(source_column)
    target_column = normalize_column(target_column)
    normalized_mark_styles = normalize_mark_styles(mark_styles=mark_styles, mark_style=mark_style)
    output_path = (
        Path(output_file).expanduser().resolve()
        if output_file
        else build_default_output_path(input_path)
    )

    workbook = load_workbook(input_path)
    worksheet = workbook[sheet] if sheet else workbook.active

    term_mapping: dict[str, str] = {}
    problem_entries: list[tuple[int, str, str]] = []
    problem_row_set: set[int] = set()

    for row_index in range(start_row, worksheet.max_row + 1):
        source_terms = extract_terms(
            worksheet[f"{source_column}{row_index}"].value,
            mark_styles=normalized_mark_styles,
        )
        target_terms = extract_terms(
            worksheet[f"{target_column}{row_index}"].value,
            mark_styles=normalized_mark_styles,
        )

        if not source_terms and not target_terms:
            continue

        row_has_problem = False
        problem_type = ""
        problem_description = ""

        if len(source_terms) != len(target_terms):
            row_has_problem = True
            problem_type = "术语数量不一致"
            problem_description = (
                f"术语数量不一致：source={format_term_list(source_terms)}；"
                f"target={format_term_list(target_terms)}"
            )
        else:
            pending_new_mappings: list[tuple[str, str]] = []
            for source_term, target_term in zip(source_terms, target_terms):
                existing_target = term_mapping.get(source_term)
                if existing_target is None:
                    pending_new_mappings.append((source_term, target_term))
                elif existing_target != target_term:
                    row_has_problem = True
                    problem_type = "术语未对齐"
                    problem_description = (
                        f"术语未对齐：source={source_term}；"
                        f"预期target={existing_target}；实际target={target_term}"
                    )
                    break

            if not row_has_problem:
                for source_term, target_term in pending_new_mappings:
                    if source_term not in term_mapping:
                        term_mapping[source_term] = target_term

        if row_has_problem and row_index not in problem_row_set:
            problem_row_set.add(row_index)
            problem_entries.append((row_index, problem_type, problem_description))

    term_sheet = rebuild_output_sheet(workbook, worksheet.title, TERM_SHEET_NAME)
    term_sheet["A1"] = "source术语"
    term_sheet["B1"] = "target术语"
    for row_index, (source_term, target_term) in enumerate(term_mapping.items(), start=2):
        term_sheet[f"A{row_index}"] = source_term
        term_sheet[f"B{row_index}"] = target_term

    problem_sheet = rebuild_output_sheet(workbook, worksheet.title, PROBLEM_SHEET_NAME)
    problem_sheet["A1"] = "问题行号"
    problem_sheet["B1"] = "问题类型"
    problem_sheet["C1"] = "问题简述"
    for row_index, (excel_row, problem_type, description) in enumerate(problem_entries, start=2):
        problem_sheet[f"A{row_index}"] = excel_row
        problem_sheet[f"B{row_index}"] = problem_type
        problem_sheet[f"C{row_index}"] = description

    if "术语汇总" in workbook.sheetnames:
        del workbook["术语汇总"]

    workbook.save(output_path)

    return (
        worksheet.title,
        source_column,
        target_column,
        output_path,
        len(term_mapping),
        len(problem_entries),
    )


def main() -> None:
    args = prompt_if_missing(parse_args())

    (
        worksheet_title,
        source_column,
        target_column,
        output_path,
        term_count,
        problem_count,
    ) = process_excel(
        input_file=args.input_file,
        source_column=args.source_column,
        target_column=args.target_column,
        sheet=args.sheet,
        start_row=args.start_row,
        mark_styles=args.mark_style,
        output_file=args.output,
    )

    print("处理完成。")
    print(f"工作表: {worksheet_title}")
    print(f"source 列: {source_column}")
    print(f"target 列: {target_column}")
    print(f"mark 类型: {'、'.join(args.mark_style)}")
    print(f"术语表条目数: {term_count}")
    print(f"问题行数: {problem_count}")
    print(f"输出文件: {output_path}")


if __name__ == "__main__":
    main()
