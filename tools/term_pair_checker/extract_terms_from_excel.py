#!/usr/bin/env python3
"""Validate source/target term pairs extracted from an Excel worksheet."""

from __future__ import annotations

import argparse
import re
import sys
from collections.abc import Iterable
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

from tools.term_matching import (
    TermMappingEntry,
    build_matcher,
    find_row_terms,
    normalize_text,
    text_contains_term,
)


SUPPORTED_MARKS = ("【】", "[]", "<>")
TERM_SHEET_NAME = "术语表"
PROBLEM_SHEET_NAME = "问题列"
DEFAULT_MARK_STYLES = ("【】",)
PAIR_CHECK_MATCH_MODE = "hybrid-boundary"
PAIR_CHECK_CASE_SENSITIVE = False
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


@dataclass(frozen=True)
class ExtractedTerm:
    display_text: str
    plain_text: str
    start: int
    end: int


@dataclass(frozen=True)
class RecordedTermPair:
    source_display_text: str
    target_display_text: str
    source_plain_text: str
    target_plain_text: str


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
    return [
        extracted_term.display_text
        for extracted_term in extract_term_details(
            text,
            mark_styles=mark_styles,
            mark_style=mark_style,
        )
    ]


def extract_term_details(
    text: object,
    mark_styles: Iterable[str] | None = None,
    mark_style: str | None = None,
) -> list[ExtractedTerm]:
    if text is None:
        return []

    normalized_mark_styles = normalize_mark_styles(mark_styles=mark_styles, mark_style=mark_style)
    text_value = str(text)
    matches: list[ExtractedTerm] = []

    for current_mark_style in normalized_mark_styles:
        for pattern in MARK_PATTERNS[current_mark_style]:
            for match in pattern.finditer(text_value):
                matches.append(
                    ExtractedTerm(
                        display_text=match.group(0),
                        plain_text=match.group(1).strip(),
                        start=match.start(),
                        end=match.end(),
                    )
                )

    matches.sort(key=lambda item: (item.start, item.end))
    return matches


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


def strip_supported_marks(text: object, mark_styles: Iterable[str] | None = None) -> str:
    text_value = "" if text is None else str(text)
    if not text_value:
        return ""

    normalized_mark_styles = normalize_mark_styles(
        mark_styles=SUPPORTED_MARKS if mark_styles is None else mark_styles
    )
    extracted_terms = extract_term_details(text_value, mark_styles=normalized_mark_styles)
    if not extracted_terms:
        return text_value

    parts: list[str] = []
    last_index = 0
    for extracted_term in extracted_terms:
        parts.append(text_value[last_index : extracted_term.start])
        parts.append(extracted_term.plain_text)
        last_index = extracted_term.end
    parts.append(text_value[last_index:])
    return "".join(parts)


def build_term_mapping_entries(term_pairs: Iterable[RecordedTermPair]) -> list[TermMappingEntry]:
    entries = [
        TermMappingEntry(
            source_term=term_pair.source_plain_text,
            target_term=term_pair.target_plain_text,
            normalized_source=normalize_text(
                term_pair.source_plain_text,
                case_sensitive=PAIR_CHECK_CASE_SENSITIVE,
            ),
            normalized_target=normalize_text(
                term_pair.target_plain_text,
                case_sensitive=PAIR_CHECK_CASE_SENSITIVE,
            ),
        )
        for term_pair in term_pairs
    ]
    entries.sort(key=lambda entry: (len(entry.normalized_source), entry.normalized_source), reverse=True)
    return entries


def append_problem(
    problem_entries: list[tuple[int, str, str]],
    problem_row_set: set[int],
    row_index: int,
    problem_type: str,
    problem_description: str,
) -> None:
    if row_index in problem_row_set:
        return
    problem_row_set.add(row_index)
    problem_entries.append((row_index, problem_type, problem_description))


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

    term_mapping: dict[str, RecordedTermPair] = {}
    problem_entries: list[tuple[int, str, str]] = []
    problem_row_set: set[int] = set()

    for row_index in range(start_row, worksheet.max_row + 1):
        source_terms = extract_term_details(
            worksheet[f"{source_column}{row_index}"].value,
            mark_styles=normalized_mark_styles,
        )
        target_terms = extract_term_details(
            worksheet[f"{target_column}{row_index}"].value,
            mark_styles=normalized_mark_styles,
        )

        if not source_terms and not target_terms:
            continue

        row_has_problem = False
        problem_type = ""
        problem_description = ""

        if len(source_terms) != len(target_terms):
            append_problem(
                problem_entries,
                problem_row_set,
                row_index,
                "术语数量不一致",
                (
                    "术语数量不一致："
                    f"source={format_term_list([term.display_text for term in source_terms])}；"
                    f"target={format_term_list([term.display_text for term in target_terms])}"
                ),
            )
            row_has_problem = True
        else:
            pending_new_mappings: list[RecordedTermPair] = []
            for source_term, target_term in zip(source_terms, target_terms):
                existing_term_pair = term_mapping.get(source_term.plain_text)
                if existing_term_pair is None:
                    pending_new_mappings.append(
                        RecordedTermPair(
                            source_display_text=source_term.display_text,
                            target_display_text=target_term.display_text,
                            source_plain_text=source_term.plain_text,
                            target_plain_text=target_term.plain_text,
                        )
                    )
                elif existing_term_pair.target_plain_text != target_term.plain_text:
                    row_has_problem = True
                    append_problem(
                        problem_entries,
                        problem_row_set,
                        row_index,
                        "术语未对齐",
                        (
                            f"术语未对齐：source={source_term.plain_text}；"
                            f"预期target={existing_term_pair.target_plain_text}；"
                            f"实际target={target_term.plain_text}；"
                            "术语对示例="
                            f"{existing_term_pair.source_display_text} -> "
                            f"{existing_term_pair.target_display_text}"
                        ),
                    )
                    break

            if not row_has_problem:
                for term_pair in pending_new_mappings:
                    term_mapping.setdefault(term_pair.source_plain_text, term_pair)

    matcher = None
    if term_mapping:
        matcher = build_matcher(build_term_mapping_entries(term_mapping.values()))

    for row_index in range(start_row, worksheet.max_row + 1):
        if matcher is None or row_index in problem_row_set:
            continue

        source_text = strip_supported_marks(worksheet[f"{source_column}{row_index}"].value)
        target_text = strip_supported_marks(worksheet[f"{target_column}{row_index}"].value)
        matched_entries = find_row_terms(
            source_text,
            matcher,
            case_sensitive=PAIR_CHECK_CASE_SENSITIVE,
            match_mode=PAIR_CHECK_MATCH_MODE,
        )
        if not matched_entries:
            continue

        normalized_target_text = normalize_text(target_text, case_sensitive=PAIR_CHECK_CASE_SENSITIVE)
        for entry in matched_entries:
            if text_contains_term(
                normalized_target_text,
                entry.normalized_target,
                match_mode=PAIR_CHECK_MATCH_MODE,
            ):
                continue

            example_term_pair = term_mapping[entry.source_term]
            append_problem(
                problem_entries,
                problem_row_set,
                row_index,
                "术语未对齐",
                (
                    f"术语未对齐：source={entry.source_term}；"
                    f"预期target={entry.target_term}；"
                    "术语对示例="
                    f"{example_term_pair.source_display_text} -> "
                    f"{example_term_pair.target_display_text}"
                ),
            )
            break

    term_sheet = rebuild_output_sheet(workbook, worksheet.title, TERM_SHEET_NAME)
    term_sheet["A1"] = "source术语"
    term_sheet["B1"] = "target术语"
    for row_index, term_pair in enumerate(term_mapping.values(), start=2):
        term_sheet[f"A{row_index}"] = term_pair.source_display_text
        term_sheet[f"B{row_index}"] = term_pair.target_display_text

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
