#!/usr/bin/env python3
"""Validate source/target term pairs extracted from an Excel worksheet."""

from __future__ import annotations

import argparse
import json
import re
import sys
from collections.abc import Iterable
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter

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
DEFAULT_EXCLUSION_CONFIG_NAME = "false_positive_exclusions.json"
PAIR_CHECK_MATCH_MODE = "hybrid-boundary"
PAIR_CHECK_CASE_SENSITIVE = False
HISTORY_EMPTY_ROW_STOP_THRESHOLD = 1000
TERM_SOURCE_HISTORY = "历史TB"
TERM_SOURCE_BATCH = "本批次新增"
HISTORY_SOURCE_EXACT_HEADERS = {"source"}
HISTORY_TARGET_EXACT_HEADERS = {"target"}
HISTORY_SOURCE_TERM_HEADERS = {"source术语"}
HISTORY_TARGET_TERM_HEADERS = {"target术语"}
HISTORY_SOURCE_NO_MARK_HEADERS = {
    "source术语（无mark）",
    "source术语(无mark)",
    "source（无mark）",
    "source(无mark)",
}
HISTORY_TARGET_NO_MARK_HEADERS = {
    "target术语（无mark）",
    "target术语(无mark)",
    "target（无mark）",
    "target(无mark)",
}
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
    term_source: str = TERM_SOURCE_BATCH


@dataclass(frozen=True)
class HistoryTbColumns:
    sheet_title: str
    source_column: str | None
    target_column: str | None


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
    exclusion_patterns: Iterable[str] | None = None,
    exclusion_config_file: str | Path | None = None,
) -> list[str]:
    return [
        extracted_term.display_text
        for extracted_term in extract_term_details(
            text,
            mark_styles=mark_styles,
            mark_style=mark_style,
            exclusion_patterns=exclusion_patterns,
            exclusion_config_file=exclusion_config_file,
        )
    ]


def build_default_exclusion_config_path() -> Path:
    return Path(__file__).with_name(DEFAULT_EXCLUSION_CONFIG_NAME)


def normalize_exclusion_patterns(exclusion_patterns: Iterable[str] | None) -> tuple[str, ...]:
    if exclusion_patterns is None:
        raw_patterns: list[str] = []
    elif isinstance(exclusion_patterns, str):
        raw_patterns = [exclusion_patterns]
    else:
        raw_patterns = [pattern.strip() for pattern in exclusion_patterns if pattern and pattern.strip()]
    return tuple(raw_patterns)


def load_exclusion_patterns_from_file(config_file: str | Path | None = None) -> tuple[str, ...]:
    config_path = (
        Path(config_file).expanduser().resolve()
        if config_file
        else build_default_exclusion_config_path().resolve()
    )
    if not config_path.exists():
        raise FileNotFoundError(f"误判排除配置文件不存在: {config_path}")

    try:
        config_data = json.loads(config_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"误判排除配置文件不是有效 JSON: {config_path} ({exc})") from exc

    patterns = config_data.get("patterns")
    if not isinstance(patterns, list) or any(not isinstance(pattern, str) for pattern in patterns):
        raise ValueError(f"误判排除配置格式错误: {config_path}，需要 JSON 对象中的 patterns 字符串数组。")
    return normalize_exclusion_patterns(patterns)


def resolve_exclusion_patterns(
    exclusion_patterns: Iterable[str] | None,
    exclusion_config_file: str | Path | None = None,
) -> tuple[str, ...]:
    if exclusion_patterns is not None:
        return normalize_exclusion_patterns(exclusion_patterns)
    return load_exclusion_patterns_from_file(exclusion_config_file)


def compile_exclusion_patterns(
    exclusion_patterns: Iterable[str] | None,
    exclusion_config_file: str | Path | None = None,
) -> tuple[re.Pattern[str], ...]:
    normalized_patterns = resolve_exclusion_patterns(exclusion_patterns, exclusion_config_file)
    compiled_patterns: list[re.Pattern[str]] = []
    for pattern in normalized_patterns:
        try:
            compiled_patterns.append(re.compile(pattern, re.IGNORECASE))
        except re.error as exc:
            raise ValueError(f"误判排除正则无效: {pattern} ({exc})") from exc
    return tuple(compiled_patterns)


def should_exclude_term(
    display_text: str,
    plain_text: str,
    exclusion_regexes: Iterable[re.Pattern[str]],
) -> bool:
    return any(
        regex.search(plain_text) or regex.search(display_text)
        for regex in exclusion_regexes
    )


def extract_term_details(
    text: object,
    mark_styles: Iterable[str] | None = None,
    mark_style: str | None = None,
    exclusion_patterns: Iterable[str] | None = None,
    exclusion_config_file: str | Path | None = None,
) -> list[ExtractedTerm]:
    if text is None:
        return []

    normalized_mark_styles = normalize_mark_styles(mark_styles=mark_styles, mark_style=mark_style)
    exclusion_regexes = compile_exclusion_patterns(exclusion_patterns, exclusion_config_file)
    text_value = str(text)
    matches: list[ExtractedTerm] = []

    for current_mark_style in normalized_mark_styles:
        for pattern in MARK_PATTERNS[current_mark_style]:
            for match in pattern.finditer(text_value):
                display_text = match.group(0)
                plain_text = match.group(1).strip()
                if should_exclude_term(display_text, plain_text, exclusion_regexes):
                    continue
                matches.append(
                    ExtractedTerm(
                        display_text=display_text,
                        plain_text=plain_text,
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
        "--exclusion-config",
        help=(
            "误判排除 JSON 配置文件路径；默认读取工具目录下的 "
            f"{DEFAULT_EXCLUSION_CONFIG_NAME}"
        ),
    )
    parser.add_argument(
        "--history-tb",
        help="历史 TB Excel 文件路径；可选，命中历史 source 时优先使用历史 target。",
    )
    parser.add_argument(
        "--history-sheet",
        help="历史 TB 工作表名称；默认优先使用“术语表”，否则使用活动工作表。",
    )
    parser.add_argument(
        "--history-source-column",
        help="历史 TB source 列；不填则自动识别 source/target 表头。",
    )
    parser.add_argument(
        "--history-target-column",
        help="历史 TB target 列；不填则自动识别 source/target 表头。",
    )
    parser.add_argument(
        "--history-start-row",
        type=int,
        default=2,
        help="历史 TB 开始读取行号，默认 2。",
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


def normalize_header(value: object) -> str:
    return "" if value is None else str(value).strip().casefold().replace(" ", "")


def find_unique_header_column(header_row: tuple[object, ...], candidates: set[str]) -> str | None:
    matches = [
        get_column_letter(column_index)
        for column_index, value in enumerate(header_row, start=1)
        if normalize_header(value) in candidates
    ]
    return matches[0] if len(matches) == 1 else None


def find_priority_header_column(header_row: tuple[object, ...], candidate_groups: tuple[set[str], ...]) -> str | None:
    for candidates in candidate_groups:
        column = find_unique_header_column(header_row, candidates)
        if column:
            return column
    return None


def find_two_column_header_fallback(header_row: tuple[object, ...]) -> tuple[str, str] | None:
    nonempty_columns = [
        get_column_letter(column_index)
        for column_index, value in enumerate(header_row, start=1)
        if normalize_header(value)
    ]
    if len(nonempty_columns) != 2:
        return None
    return nonempty_columns[0], nonempty_columns[1]


def detect_history_columns_from_header(header_row: tuple[object, ...]) -> tuple[str | None, str | None]:
    source_column = find_priority_header_column(
        header_row,
        (
            HISTORY_SOURCE_EXACT_HEADERS,
            HISTORY_SOURCE_TERM_HEADERS,
            HISTORY_SOURCE_NO_MARK_HEADERS,
        ),
    )
    target_column = find_priority_header_column(
        header_row,
        (
            HISTORY_TARGET_EXACT_HEADERS,
            HISTORY_TARGET_TERM_HEADERS,
            HISTORY_TARGET_NO_MARK_HEADERS,
        ),
    )
    if source_column and target_column:
        return source_column, target_column

    fallback_columns = find_two_column_header_fallback(header_row)
    if fallback_columns is None:
        return source_column, target_column

    fallback_source_column, fallback_target_column = fallback_columns
    if source_column and not target_column:
        target_column = fallback_target_column if fallback_target_column != source_column else fallback_source_column
    elif target_column and not source_column:
        source_column = fallback_source_column if fallback_source_column != target_column else fallback_target_column
    else:
        source_column, target_column = fallback_columns
    return source_column, target_column


def choose_history_worksheet(workbook, sheet: str | None):
    if sheet:
        if sheet not in workbook.sheetnames:
            raise ValueError(f"历史 TB 工作表不存在: {sheet}")
        return workbook[sheet]
    if TERM_SHEET_NAME in workbook.sheetnames:
        return workbook[TERM_SHEET_NAME]
    return workbook.active


def detect_history_tb_columns(
    history_tb_file: str | Path,
    sheet: str | None = None,
) -> HistoryTbColumns:
    history_path = Path(history_tb_file).expanduser().resolve()
    if not history_path.exists():
        raise FileNotFoundError(f"历史 TB 文件不存在: {history_path}")

    workbook = load_workbook(history_path, read_only=True)
    try:
        worksheet = choose_history_worksheet(workbook, sheet)
        header_row = next(
            worksheet.iter_rows(min_row=1, max_row=1, values_only=True),
            (),
        )
        source_column, target_column = detect_history_columns_from_header(tuple(header_row))
        return HistoryTbColumns(
            sheet_title=worksheet.title,
            source_column=source_column,
            target_column=target_column,
        )
    finally:
        workbook.close()


def normalize_history_source_key(source_term: str) -> str:
    return normalize_text(source_term, case_sensitive=PAIR_CHECK_CASE_SENSITIVE)


def row_value(row: tuple[object, ...], column_index: int) -> object:
    return row[column_index - 1] if len(row) >= column_index else None


def row_values_are_empty(*values: object) -> bool:
    return all(value is None or str(value).strip() == "" for value in values)


def load_history_tb_mapping(
    history_tb_file: str | Path,
    source_column: str | None = None,
    target_column: str | None = None,
    sheet: str | None = None,
    start_row: int = 2,
    exclusion_patterns: Iterable[str] | None = None,
) -> dict[str, RecordedTermPair]:
    history_path = Path(history_tb_file).expanduser().resolve()
    if not history_path.exists():
        raise FileNotFoundError(f"历史 TB 文件不存在: {history_path}")
    if start_row < 1:
        raise ValueError("历史 TB 开始行必须大于等于 1。")

    workbook = load_workbook(history_path, read_only=True)
    try:
        worksheet = choose_history_worksheet(workbook, sheet)
        header_row = next(
            worksheet.iter_rows(min_row=1, max_row=1, values_only=True),
            (),
        )
        detected_source_column, detected_target_column = detect_history_columns_from_header(tuple(header_row))

        source_column = normalize_column(source_column) if source_column else detected_source_column
        target_column = normalize_column(target_column) if target_column else detected_target_column
        if not source_column or not target_column:
            raise ValueError(
                "历史 TB 缺少 source/target 列，请指定 --history-source-column 和 "
                "--history-target-column，或在第 1 行使用 source/target 表头；两列表头的 TB 会默认按第 1 列 source、第 2 列 target 读取。"
            )
        source_column_index = column_index_from_string(source_column)
        target_column_index = column_index_from_string(target_column)
        max_column_index = max(source_column_index, target_column_index)

        history_mapping: dict[str, RecordedTermPair] = {}
        consecutive_empty_rows = 0
        for row in worksheet.iter_rows(
            min_row=start_row,
            max_col=max_column_index,
            values_only=True,
        ):
            raw_source = row_value(row, source_column_index)
            raw_target = row_value(row, target_column_index)
            if row_values_are_empty(raw_source, raw_target):
                consecutive_empty_rows += 1
                if consecutive_empty_rows >= HISTORY_EMPTY_ROW_STOP_THRESHOLD:
                    break
                continue

            consecutive_empty_rows = 0
            source_plain_text = strip_supported_marks(
                raw_source,
                exclusion_patterns=exclusion_patterns,
            ).strip()
            target_plain_text = strip_supported_marks(
                raw_target,
                exclusion_patterns=exclusion_patterns,
            ).strip()
            if not source_plain_text or not target_plain_text:
                continue

            normalized_source = normalize_history_source_key(source_plain_text)
            if not normalized_source:
                continue
            history_mapping.setdefault(
                normalized_source,
                RecordedTermPair(
                    source_display_text=source_plain_text,
                    target_display_text=target_plain_text,
                    source_plain_text=source_plain_text,
                    target_plain_text=target_plain_text,
                    term_source=TERM_SOURCE_HISTORY,
                ),
            )
        return history_mapping
    finally:
        workbook.close()


def build_default_output_path(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}_term_pairs{input_path.suffix}")


def rebuild_output_sheet(workbook, current_sheet_name: str, sheet_name: str):
    if current_sheet_name == sheet_name:
        raise ValueError(f"数据工作表名称不能为 {sheet_name}")
    if sheet_name in workbook.sheetnames:
        del workbook[sheet_name]
    return workbook.create_sheet(title=sheet_name)


def strip_supported_marks(
    text: object,
    mark_styles: Iterable[str] | None = None,
    exclusion_patterns: Iterable[str] | None = None,
    exclusion_config_file: str | Path | None = None,
) -> str:
    text_value = "" if text is None else str(text)
    if not text_value:
        return ""

    normalized_mark_styles = normalize_mark_styles(
        mark_styles=SUPPORTED_MARKS if mark_styles is None else mark_styles
    )
    extracted_terms = extract_term_details(
        text_value,
        mark_styles=normalized_mark_styles,
        exclusion_patterns=exclusion_patterns,
        exclusion_config_file=exclusion_config_file,
    )
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
        if term_pair.target_plain_text
    ]
    entries.sort(key=lambda entry: (len(entry.normalized_source), entry.normalized_source), reverse=True)
    return entries


def merge_term_pair(
    term_mapping: dict[str, RecordedTermPair],
    term_pair: RecordedTermPair,
) -> tuple[bool, RecordedTermPair | None]:
    existing_term_pair = term_mapping.get(term_pair.source_plain_text)
    if existing_term_pair is None:
        term_mapping[term_pair.source_plain_text] = term_pair
        return True, None

    if not existing_term_pair.target_plain_text and term_pair.target_plain_text:
        term_mapping[term_pair.source_plain_text] = term_pair
        return True, None

    if (
        existing_term_pair.target_plain_text
        and term_pair.target_plain_text
        and existing_term_pair.target_plain_text != term_pair.target_plain_text
    ):
        return False, existing_term_pair

    return True, existing_term_pair


def append_problem(
    problem_entries: list[tuple[int, str, str, str, str, str]],
    problem_row_set: set[int],
    row_index: int,
    problem_source_term: str,
    expected_target_term: str,
    problem_description: str,
    source_snapshot: str,
    target_snapshot: str,
) -> None:
    if row_index in problem_row_set:
        return
    problem_row_set.add(row_index)
    problem_entries.append(
        (
            row_index,
            problem_source_term,
            expected_target_term,
            problem_description,
            source_snapshot,
            target_snapshot,
        )
    )


def build_text_snapshot(value: object) -> str:
    return "" if value is None else str(value)


def format_problem_term(terms: Iterable[str]) -> str:
    unique_terms: list[str] = []
    seen_terms: set[str] = set()
    for term in terms:
        normalized_term = term.strip()
        if not normalized_term or normalized_term in seen_terms:
            continue
        seen_terms.add(normalized_term)
        unique_terms.append(normalized_term)
    return "、".join(unique_terms)


def format_expected_target_terms(
    source_terms: Iterable[ExtractedTerm],
    term_mapping: dict[str, RecordedTermPair],
) -> str:
    expected_targets: list[str] = []
    seen_targets: set[str] = set()
    for source_term in source_terms:
        mapped_term = term_mapping.get(source_term.plain_text)
        if mapped_term is None or not mapped_term.target_plain_text:
            continue
        if mapped_term.target_plain_text in seen_targets:
            continue
        seen_targets.add(mapped_term.target_plain_text)
        expected_targets.append(mapped_term.target_plain_text)
    return "、".join(expected_targets)


def row_terms_are_aligned(
    matched_entries: list[TermMappingEntry],
    normalized_target_text: str,
) -> bool:
    for entry in matched_entries:
        if not text_contains_term(
            normalized_target_text,
            entry.normalized_target,
            match_mode=PAIR_CHECK_MATCH_MODE,
        ):
            return False
    return True


def count_mismatch_is_resolved(
    source_terms: list[ExtractedTerm],
    target_terms: list[ExtractedTerm],
    matched_entries: list[TermMappingEntry],
    normalized_target_text: str,
) -> bool:
    if not matched_entries:
        return False
    if not row_terms_are_aligned(matched_entries, normalized_target_text):
        return False

    matched_source_terms = {entry.source_term for entry in matched_entries}
    matched_target_terms = {entry.target_term for entry in matched_entries}
    if any(source_term.plain_text not in matched_source_terms for source_term in source_terms):
        return False
    if any(target_term.plain_text not in matched_target_terms for target_term in target_terms):
        return False
    return True


def build_recorded_term_pair(
    source_term: ExtractedTerm,
    target_term: ExtractedTerm | None,
    history_mapping: dict[str, RecordedTermPair],
) -> RecordedTermPair:
    history_term_pair = history_mapping.get(normalize_history_source_key(source_term.plain_text))
    if history_term_pair is not None:
        return RecordedTermPair(
            source_display_text=source_term.display_text,
            target_display_text=history_term_pair.target_plain_text,
            source_plain_text=source_term.plain_text,
            target_plain_text=history_term_pair.target_plain_text,
            term_source=TERM_SOURCE_HISTORY,
        )

    return RecordedTermPair(
        source_display_text=source_term.display_text,
        target_display_text=target_term.display_text if target_term else "",
        source_plain_text=source_term.plain_text,
        target_plain_text=target_term.plain_text if target_term else "",
        term_source=TERM_SOURCE_BATCH,
    )


def process_excel(
    input_file: str | Path,
    source_column: str,
    target_column: str,
    sheet: str | None = None,
    start_row: int = 2,
    mark_styles: Iterable[str] | None = None,
    mark_style: str | None = None,
    exclusion_patterns: Iterable[str] | None = None,
    exclusion_config_file: str | Path | None = None,
    history_tb_file: str | Path | None = None,
    history_sheet: str | None = None,
    history_source_column: str | None = None,
    history_target_column: str | None = None,
    history_start_row: int = 2,
    output_file: str | Path | None = None,
) -> tuple[str, str, str, Path, int, int]:
    input_path = Path(input_file).expanduser().resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"输入文件不存在: {input_path}")

    source_column = normalize_column(source_column)
    target_column = normalize_column(target_column)
    normalized_mark_styles = normalize_mark_styles(mark_styles=mark_styles, mark_style=mark_style)
    effective_exclusion_patterns = resolve_exclusion_patterns(exclusion_patterns, exclusion_config_file)
    history_mapping = (
        load_history_tb_mapping(
            history_tb_file,
            source_column=history_source_column,
            target_column=history_target_column,
            sheet=history_sheet,
            start_row=history_start_row,
            exclusion_patterns=effective_exclusion_patterns,
        )
        if history_tb_file
        else {}
    )
    output_path = (
        Path(output_file).expanduser().resolve()
        if output_file
        else build_default_output_path(input_path)
    )

    workbook = load_workbook(input_path)
    worksheet = workbook[sheet] if sheet else workbook.active

    term_mapping: dict[str, RecordedTermPair] = {}
    count_mismatch_rows: dict[int, tuple[list[ExtractedTerm], list[ExtractedTerm]]] = {}
    problem_entries: list[tuple[int, str, str, str, str, str]] = []
    problem_row_set: set[int] = set()

    for row_index in range(start_row, worksheet.max_row + 1):
        raw_source_value = worksheet[f"{source_column}{row_index}"].value
        raw_target_value = worksheet[f"{target_column}{row_index}"].value
        source_snapshot = build_text_snapshot(raw_source_value)
        target_snapshot = build_text_snapshot(raw_target_value)

        source_terms = extract_term_details(
            raw_source_value,
            mark_styles=normalized_mark_styles,
            exclusion_patterns=effective_exclusion_patterns,
        )
        target_terms = extract_term_details(
            raw_target_value,
            mark_styles=normalized_mark_styles,
            exclusion_patterns=effective_exclusion_patterns,
        )

        if not source_terms and not target_terms:
            continue

        if len(source_terms) != len(target_terms):
            count_mismatch_rows[row_index] = (source_terms, target_terms)
            for source_term in source_terms[len(target_terms) :]:
                merge_term_pair(
                    term_mapping,
                    build_recorded_term_pair(source_term, None, history_mapping),
                )
        else:
            candidate_term_mapping = dict(term_mapping)
            row_has_problem = False
            for source_term, target_term in zip(source_terms, target_terms):
                merged, existing_term_pair = merge_term_pair(
                    candidate_term_mapping,
                    build_recorded_term_pair(source_term, target_term, history_mapping),
                )
                if not merged:
                    row_has_problem = True
                    append_problem(
                        problem_entries,
                        problem_row_set,
                        row_index,
                        source_term.plain_text,
                        existing_term_pair.target_plain_text,
                        f"target术语不匹配：实际术语 - {target_term.plain_text}",
                        source_snapshot,
                        target_snapshot,
                    )
                    break

            if not row_has_problem:
                term_mapping = candidate_term_mapping

    matcher = None
    if term_mapping:
        matcher = build_matcher(build_term_mapping_entries(term_mapping.values()))

    for row_index in range(start_row, worksheet.max_row + 1):
        if matcher is None or row_index in problem_row_set:
            if row_index in count_mismatch_rows and row_index not in problem_row_set:
                source_terms, target_terms = count_mismatch_rows[row_index]
                append_problem(
                    problem_entries,
                    problem_row_set,
                    row_index,
                    format_problem_term(term.plain_text for term in source_terms),
                    format_expected_target_terms(source_terms, term_mapping),
                    (
                        f"source/target术语数量不一致：{len(source_terms)}（预期数量）- "
                        f"{len(target_terms)}（实际数量）"
                    ),
                    build_text_snapshot(worksheet[f"{source_column}{row_index}"].value),
                    build_text_snapshot(worksheet[f"{target_column}{row_index}"].value),
                )
            continue

        source_text = strip_supported_marks(
            worksheet[f"{source_column}{row_index}"].value,
            exclusion_patterns=effective_exclusion_patterns,
        )
        target_text = strip_supported_marks(
            worksheet[f"{target_column}{row_index}"].value,
            exclusion_patterns=effective_exclusion_patterns,
        )
        matched_entries = find_row_terms(
            source_text,
            matcher,
            case_sensitive=PAIR_CHECK_CASE_SENSITIVE,
            match_mode=PAIR_CHECK_MATCH_MODE,
        )
        if not matched_entries:
            if row_index in count_mismatch_rows:
                source_terms, target_terms = count_mismatch_rows[row_index]
                append_problem(
                    problem_entries,
                    problem_row_set,
                    row_index,
                    format_problem_term(term.plain_text for term in source_terms),
                    format_expected_target_terms(source_terms, term_mapping),
                    (
                        f"source/target术语数量不一致：{len(source_terms)}（预期数量）- "
                        f"{len(target_terms)}（实际数量）"
                    ),
                    build_text_snapshot(worksheet[f"{source_column}{row_index}"].value),
                    build_text_snapshot(worksheet[f"{target_column}{row_index}"].value),
                )
            continue

        normalized_target_text = normalize_text(target_text, case_sensitive=PAIR_CHECK_CASE_SENSITIVE)
        if row_index in count_mismatch_rows:
            source_terms, target_terms = count_mismatch_rows[row_index]
            if count_mismatch_is_resolved(
                source_terms,
                target_terms,
                matched_entries,
                normalized_target_text,
            ):
                continue
            append_problem(
                problem_entries,
                problem_row_set,
                row_index,
                format_problem_term(term.plain_text for term in source_terms),
                format_expected_target_terms(source_terms, term_mapping),
                (
                    f"source/target术语数量不一致：{len(source_terms)}（预期数量）- "
                    f"{len(target_terms)}（实际数量）"
                ),
                build_text_snapshot(worksheet[f"{source_column}{row_index}"].value),
                build_text_snapshot(worksheet[f"{target_column}{row_index}"].value),
            )
            continue

        for entry in matched_entries:
            if text_contains_term(normalized_target_text, entry.normalized_target, match_mode=PAIR_CHECK_MATCH_MODE):
                continue

            append_problem(
                problem_entries,
                problem_row_set,
                row_index,
                entry.source_term,
                entry.target_term,
                (
                    "target缺少预期术语"
                ),
                build_text_snapshot(worksheet[f"{source_column}{row_index}"].value),
                build_text_snapshot(worksheet[f"{target_column}{row_index}"].value),
            )
            break

    term_sheet = rebuild_output_sheet(workbook, worksheet.title, TERM_SHEET_NAME)
    term_sheet["A1"] = "source术语"
    term_sheet["B1"] = "target术语"
    term_sheet["C1"] = "source术语（无mark）"
    term_sheet["D1"] = "target术语（无mark）"
    term_sheet["E1"] = "术语来源"
    for row_index, term_pair in enumerate(term_mapping.values(), start=2):
        term_sheet[f"A{row_index}"] = term_pair.source_display_text
        term_sheet[f"B{row_index}"] = term_pair.target_display_text
        term_sheet[f"C{row_index}"] = term_pair.source_plain_text
        term_sheet[f"D{row_index}"] = term_pair.target_plain_text
        term_sheet[f"E{row_index}"] = term_pair.term_source

    problem_sheet = rebuild_output_sheet(workbook, worksheet.title, PROBLEM_SHEET_NAME)
    problem_sheet["A1"] = "问题行号"
    problem_sheet["B1"] = "问题source术语"
    problem_sheet["C1"] = "预期target术语"
    problem_sheet["D1"] = "问题简述"
    problem_sheet["E1"] = "source原文"
    problem_sheet["F1"] = "target原文"
    sorted_problem_entries = sorted(
        problem_entries,
        key=lambda entry: (
            entry[1] == "",
            normalize_text(entry[1], case_sensitive=False),
            entry[0],
        ),
    )
    for row_index, (
        excel_row,
        problem_source_term,
        expected_target_term,
        description,
        source_snapshot,
        target_snapshot,
    ) in enumerate(
        sorted_problem_entries,
        start=2,
    ):
        problem_sheet[f"A{row_index}"] = excel_row
        problem_sheet[f"B{row_index}"] = problem_source_term
        problem_sheet[f"C{row_index}"] = expected_target_term
        problem_sheet[f"D{row_index}"] = description
        problem_sheet[f"E{row_index}"] = source_snapshot
        problem_sheet[f"F{row_index}"] = target_snapshot

    if "术语表（无mark）" in workbook.sheetnames:
        del workbook["术语表（无mark）"]
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
        exclusion_config_file=args.exclusion_config,
        history_tb_file=args.history_tb,
        history_sheet=args.history_sheet,
        history_source_column=args.history_source_column,
        history_target_column=args.history_target_column,
        history_start_row=args.history_start_row,
        output_file=args.output,
    )

    print("处理完成。")
    print(f"工作表: {worksheet_title}")
    print(f"source 列: {source_column}")
    print(f"target 列: {target_column}")
    print(f"mark 类型: {'、'.join(args.mark_style)}")
    if args.history_tb:
        print(f"历史 TB: {Path(args.history_tb).expanduser().resolve()}")
    print(f"术语表条目数: {term_count}")
    print(f"问题行数: {problem_count}")
    print(f"输出文件: {output_path}")


if __name__ == "__main__":
    main()
