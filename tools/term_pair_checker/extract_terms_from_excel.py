#!/usr/bin/env python3
"""Check Excel terminology from optional marks and an optional history TB."""

from __future__ import annotations

import argparse
import sys
from collections.abc import Iterable
from dataclasses import dataclass
from pathlib import Path

if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from openpyxl.utils import column_index_from_string

from tools.term_matching import (
    TermMappingEntry,
    build_matcher,
    find_row_terms,
    normalize_text,
    term_has_expected_target,
)
from tools.history_tb import (
    HistoryTbColumns,
    detect_history_tb_columns as shared_detect_history_tb_columns,
    iter_history_rows,
)
from tools.excel_output import (
    find_last_value_row,
    insert_row_problem_column,
    load_workbook_for_editing,
    validate_distinct_source_target_columns,
)
from tools.term_pair_checker.term_marks import (
    DEFAULT_MARK_STYLES,
    SUPPORTED_MARKS,
    ExtractedTerm,
    extract_term_details,
    extract_terms,
    normalize_mark_styles,
    resolve_exclusion_patterns,
    should_exclude_term,
    strip_supported_marks,
)
from tools.term_pair_checker.workbook_output import (
    PROBLEM_SHEET_NAME,
    TERM_SHEET_NAME,
    build_default_output_path,
    build_row_problem_summaries,
    delete_legacy_term_sheets,
    write_problem_sheet,
    write_term_sheet,
)


PAIR_CHECK_MATCH_MODE = "hybrid-boundary"
PAIR_CHECK_CASE_SENSITIVE = False
HISTORY_EMPTY_ROW_STOP_THRESHOLD = 1000
TERM_SOURCE_HISTORY = "历史TB"
TERM_SOURCE_BATCH = "本批次新增"


@dataclass(frozen=True)
class RecordedTermPair:
    source_display_text: str
    target_display_text: str
    source_plain_text: str
    target_plain_text: str
    term_source: str = TERM_SOURCE_BATCH


@dataclass(frozen=True)
class ProblemEntry:
    row_index: int
    problem_source_term: str
    expected_target_term: str
    term_source: str
    description: str
    source_snapshot: str
    target_snapshot: str


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "检查 Excel 的 source/target 术语；可从 mark 提取新术语对，也可仅使用历史 TB 检查。"
        )
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
    mark_group = parser.add_mutually_exclusive_group()
    mark_group.add_argument(
        "--mark-style",
        action="append",
        choices=SUPPORTED_MARKS,
        default=None,
        help="术语包裹符号，可重复传入，例如 --mark-style [] --mark-style '【】'",
    )
    mark_group.add_argument(
        "--no-term-mark",
        action="store_true",
        help="不从文本 mark 提取新术语；必须同时提供 --history-tb。",
    )
    parser.add_argument(
        "--exclusion-config",
        help="可选的自定义术语候选排除 JSON 配置文件路径。",
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
        help="输出 Excel 文件路径，默认生成 term_pair_check_<原文件名>",
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
    if getattr(args, "no_term_mark", False):
        args.mark_style = ()
    elif args.mark_style is None:
        if interactive_mode and len(sys.argv) == 1:
            mark_style_text = input(
                "请输入 mark 类型（可多选，逗号分隔，如 【】,[]；默认 【】,[]）: "
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


def detect_history_tb_columns(
    history_tb_file: str | Path,
    sheet: str | None = None,
) -> HistoryTbColumns:
    return shared_detect_history_tb_columns(
        history_tb_file,
        sheet=sheet,
        preferred_sheet=TERM_SHEET_NAME,
        prefer_no_mark=False,
        allow_partial=True,
        require_unique_header_matches=True,
    )


def normalize_history_source_key(source_term: str) -> str:
    return normalize_text(source_term, case_sensitive=PAIR_CHECK_CASE_SENSITIVE)


def normalize_term_key(term: str) -> str:
    return normalize_text(term, case_sensitive=PAIR_CHECK_CASE_SENSITIVE)


def load_history_tb_mapping(
    history_tb_file: str | Path,
    source_column: str | None = None,
    target_column: str | None = None,
    sheet: str | None = None,
    start_row: int = 2,
    exclusion_patterns: Iterable[str] | None = None,
) -> dict[str, RecordedTermPair]:
    _sheet_title, _source_column, _target_column, rows = iter_history_rows(
        history_tb_file,
        sheet=sheet,
        source_column=source_column,
        target_column=target_column,
        start_row=start_row,
        header_row=1,
        preferred_sheet=TERM_SHEET_NAME,
        prefer_no_mark=False,
        empty_row_stop_threshold=HISTORY_EMPTY_ROW_STOP_THRESHOLD,
        require_unique_header_matches=True,
    )

    history_mapping: dict[str, RecordedTermPair] = {}
    for row in rows:
        source_plain_text = strip_supported_marks(
            row.source_text,
            exclusion_patterns=exclusion_patterns,
        ).strip()
        target_plain_text = strip_supported_marks(
            row.target_text,
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
    mapping_key = normalize_term_key(term_pair.source_plain_text)
    existing_term_pair = term_mapping.get(mapping_key)
    if existing_term_pair is None:
        term_mapping[mapping_key] = term_pair
        return True, None

    if not existing_term_pair.target_plain_text and term_pair.target_plain_text:
        term_mapping[mapping_key] = term_pair
        return True, None

    if (
        existing_term_pair.target_plain_text
        and term_pair.target_plain_text
        and normalize_term_key(existing_term_pair.target_plain_text)
        != normalize_term_key(term_pair.target_plain_text)
    ):
        return False, existing_term_pair

    return True, existing_term_pair


def append_problem(
    problem_entries: list[ProblemEntry],
    row_index: int,
    problem_source_term: str,
    expected_target_term: str,
    term_source: str,
    problem_description: str,
    source_snapshot: str,
    target_snapshot: str,
) -> None:
    problem_entries.append(
        ProblemEntry(
            row_index=row_index,
            problem_source_term=problem_source_term,
            expected_target_term=expected_target_term,
            term_source=term_source,
            description=problem_description,
            source_snapshot=source_snapshot,
            target_snapshot=target_snapshot,
        )
    )


def dedupe_problem_entries(problem_entries: Iterable[ProblemEntry]) -> list[ProblemEntry]:
    unique_entries: list[ProblemEntry] = []
    seen_entries: set[ProblemEntry] = set()
    for problem_entry in problem_entries:
        if problem_entry in seen_entries:
            continue
        seen_entries.add(problem_entry)
        unique_entries.append(problem_entry)
    return unique_entries


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
        mapped_term = term_mapping.get(normalize_term_key(source_term.plain_text))
        if mapped_term is None or not mapped_term.target_plain_text:
            continue
        if mapped_term.target_plain_text in seen_targets:
            continue
        seen_targets.add(mapped_term.target_plain_text)
        expected_targets.append(mapped_term.target_plain_text)
    return "、".join(expected_targets)


def format_expected_term_sources(
    source_terms: Iterable[ExtractedTerm],
    term_mapping: dict[str, RecordedTermPair],
) -> str:
    term_sources: list[str] = []
    seen_sources: set[str] = set()
    for source_term in source_terms:
        mapped_term = term_mapping.get(normalize_term_key(source_term.plain_text))
        if mapped_term is None or not mapped_term.term_source:
            continue
        if mapped_term.term_source in seen_sources:
            continue
        seen_sources.add(mapped_term.term_source)
        term_sources.append(mapped_term.term_source)
    return "、".join(term_sources)


def lookup_term_source(source_term: str, term_mapping: dict[str, RecordedTermPair]) -> str:
    term_pair = term_mapping.get(normalize_term_key(source_term))
    return term_pair.term_source if term_pair else ""


def row_terms_are_aligned(
    matched_entries: list[TermMappingEntry],
    normalized_source_text: str,
    normalized_target_text: str,
) -> bool:
    for entry in matched_entries:
        if not term_has_expected_target(
            normalized_source_text,
            normalized_target_text,
            entry,
            match_mode=PAIR_CHECK_MATCH_MODE,
            allow_target_plural_variants=True,
        ):
            return False
    return True


def count_mismatch_is_resolved(
    source_terms: list[ExtractedTerm],
    target_terms: list[ExtractedTerm],
    matched_entries: list[TermMappingEntry],
    normalized_source_text: str,
    normalized_target_text: str,
) -> bool:
    if not matched_entries:
        return False

    if source_terms:
        matched_entries_by_source = {
            normalize_term_key(entry.source_term): entry for entry in matched_entries
        }
        required_entries: list[TermMappingEntry] = []
        for source_term in source_terms:
            matched_entry = matched_entries_by_source.get(
                normalize_term_key(source_term.plain_text)
            )
            if matched_entry is None:
                return False
            required_entries.append(matched_entry)

        # Source terms define the terminology obligations for the row. A translator
        # may additionally mark a target phrase whose source phrase is unmarked; that
        # extra target mark must not turn aligned source terms into a count mismatch.
        return row_terms_are_aligned(
            required_entries,
            normalized_source_text,
            normalized_target_text,
        )

    # Preserve target-only recovery when the corresponding source phrase is unmarked
    # but can be proven through an already-known source -> target mapping.
    matched_entries_by_target = {
        normalize_term_key(entry.target_term): entry for entry in matched_entries
    }
    required_entries = []
    for target_term in target_terms:
        matched_entry = matched_entries_by_target.get(
            normalize_term_key(target_term.plain_text)
        )
        if matched_entry is None:
            return False
        required_entries.append(matched_entry)
    return row_terms_are_aligned(
        required_entries,
        normalized_source_text,
        normalized_target_text,
    )


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


def build_initial_term_mapping(history_mapping: dict[str, RecordedTermPair]) -> dict[str, RecordedTermPair]:
    return {
        normalize_term_key(term_pair.source_plain_text): term_pair
        for term_pair in history_mapping.values()
    }


def add_matched_terms_to_output(
    output_term_mapping: dict[str, RecordedTermPair],
    term_mapping: dict[str, RecordedTermPair],
    matched_entries: Iterable[TermMappingEntry],
) -> None:
    for entry in matched_entries:
        mapping_key = normalize_term_key(entry.source_term)
        term_pair = term_mapping.get(mapping_key)
        if term_pair is None:
            continue
        output_term_mapping.setdefault(mapping_key, term_pair)


def term_source_priority(term_source: str) -> int:
    sources = {source.strip() for source in term_source.split("、") if source.strip()}
    if TERM_SOURCE_HISTORY in sources:
        return 0
    if TERM_SOURCE_BATCH in sources:
        return 1
    return 2


def term_source_sort_key(term_pair: RecordedTermPair) -> int:
    return term_source_priority(term_pair.term_source)


def sorted_output_term_pairs(term_pairs: Iterable[RecordedTermPair]) -> list[RecordedTermPair]:
    return sorted(term_pairs, key=term_source_sort_key)


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
    include_row_problem_column: bool = True,
    *,
    workbook=None,
    save_output: bool = True,
) -> tuple[str, str, str, Path, int, int]:
    if start_row < 1:
        raise ValueError("开始行必须大于等于 1。")

    input_path = Path(input_file).expanduser().resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"输入文件不存在: {input_path}")

    source_column = normalize_column(source_column)
    target_column = normalize_column(target_column)
    validate_distinct_source_target_columns(source_column, target_column)
    if mark_styles is None:
        normalized_mark_styles = normalize_mark_styles(mark_style=mark_style)
    else:
        requested_mark_styles = (
            (mark_styles,)
            if isinstance(mark_styles, str)
            else tuple(style for style in mark_styles if style)
        )
        normalized_mark_styles = (
            normalize_mark_styles(mark_styles=requested_mark_styles)
            if requested_mark_styles
            else normalize_mark_styles(mark_style=mark_style)
            if mark_style
            else ()
        )
    if not normalized_mark_styles and not history_tb_file:
        raise ValueError("不选择术语 mark 时必须提供历史 TB。")
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

    owns_workbook = workbook is None
    if owns_workbook:
        workbook = load_workbook_for_editing(input_path)
    worksheet = workbook[sheet] if sheet else workbook.active
    last_row = find_last_value_row(
        worksheet,
        (source_column, target_column),
        start_row=start_row,
    )

    term_mapping = build_initial_term_mapping(history_mapping)
    output_term_mapping: dict[str, RecordedTermPair] = {}
    count_mismatch_rows: dict[int, tuple[list[ExtractedTerm], list[ExtractedTerm]]] = {}
    conflict_source_terms_by_row: dict[int, set[str]] = {}
    problem_entries: list[ProblemEntry] = []

    for row_index in range(start_row, last_row + 1):
        raw_source_value = worksheet[f"{source_column}{row_index}"].value
        raw_target_value = worksheet[f"{target_column}{row_index}"].value
        source_snapshot = build_text_snapshot(raw_source_value)
        target_snapshot = build_text_snapshot(raw_target_value)

        if normalized_mark_styles:
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
        else:
            source_terms = []
            target_terms = []

        if not source_terms and not target_terms:
            continue

        if len(source_terms) != len(target_terms):
            count_mismatch_rows[row_index] = (source_terms, target_terms)
            for source_term in source_terms[len(target_terms) :]:
                term_pair = build_recorded_term_pair(source_term, None, history_mapping)
                _, existing_term_pair = merge_term_pair(term_mapping, term_pair)
                if not term_pair.target_plain_text and existing_term_pair is None:
                    merge_term_pair(output_term_mapping, term_pair)
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
                    conflict_source_terms_by_row.setdefault(row_index, set()).add(
                        normalize_term_key(source_term.plain_text)
                    )
                    append_problem(
                        problem_entries,
                        row_index,
                        source_term.plain_text,
                        existing_term_pair.target_plain_text,
                        existing_term_pair.term_source,
                        f"target术语不匹配：实际术语 - {target_term.plain_text}",
                        source_snapshot,
                        target_snapshot,
                    )

            if not row_has_problem:
                term_mapping = candidate_term_mapping

    matcher = None
    if term_mapping:
        matcher = build_matcher(build_term_mapping_entries(term_mapping.values()))

    for row_index in range(start_row, last_row + 1):
        if matcher is None:
            if row_index in count_mismatch_rows:
                source_terms, target_terms = count_mismatch_rows[row_index]
                append_problem(
                    problem_entries,
                    row_index,
                    format_problem_term(term.plain_text for term in source_terms),
                    format_expected_target_terms(source_terms, term_mapping),
                    format_expected_term_sources(source_terms, term_mapping),
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
        add_matched_terms_to_output(output_term_mapping, term_mapping, matched_entries)
        if not matched_entries:
            if row_index in count_mismatch_rows:
                source_terms, target_terms = count_mismatch_rows[row_index]
                append_problem(
                    problem_entries,
                    row_index,
                    format_problem_term(term.plain_text for term in source_terms),
                    format_expected_target_terms(source_terms, term_mapping),
                    format_expected_term_sources(source_terms, term_mapping),
                    (
                        f"source/target术语数量不一致：{len(source_terms)}（预期数量）- "
                        f"{len(target_terms)}（实际数量）"
                    ),
                    build_text_snapshot(worksheet[f"{source_column}{row_index}"].value),
                    build_text_snapshot(worksheet[f"{target_column}{row_index}"].value),
                )
            continue

        normalized_source_text = normalize_text(source_text, case_sensitive=PAIR_CHECK_CASE_SENSITIVE)
        normalized_target_text = normalize_text(target_text, case_sensitive=PAIR_CHECK_CASE_SENSITIVE)
        if row_index in count_mismatch_rows:
            source_terms, target_terms = count_mismatch_rows[row_index]
            count_mismatch_resolved = count_mismatch_is_resolved(
                source_terms,
                target_terms,
                matched_entries,
                normalized_source_text,
                normalized_target_text,
            )
            if not count_mismatch_resolved:
                append_problem(
                    problem_entries,
                    row_index,
                    format_problem_term(term.plain_text for term in source_terms),
                    format_expected_target_terms(source_terms, term_mapping),
                    format_expected_term_sources(source_terms, term_mapping),
                    (
                        f"source/target术语数量不一致：{len(source_terms)}（预期数量）- "
                        f"{len(target_terms)}（实际数量）"
                    ),
                    build_text_snapshot(worksheet[f"{source_column}{row_index}"].value),
                    build_text_snapshot(worksheet[f"{target_column}{row_index}"].value),
                )

        conflict_source_terms = conflict_source_terms_by_row.get(row_index, set())
        for entry in matched_entries:
            if normalize_term_key(entry.source_term) in conflict_source_terms:
                continue
            if term_has_expected_target(
                normalized_source_text,
                normalized_target_text,
                entry,
                match_mode=PAIR_CHECK_MATCH_MODE,
                allow_target_plural_variants=True,
            ):
                continue

            append_problem(
                problem_entries,
                row_index,
                entry.source_term,
                entry.target_term,
                lookup_term_source(entry.source_term, term_mapping),
                "target缺少预期术语",
                build_text_snapshot(worksheet[f"{source_column}{row_index}"].value),
                build_text_snapshot(worksheet[f"{target_column}{row_index}"].value),
            )

    problem_entries = dedupe_problem_entries(problem_entries)

    if include_row_problem_column:
        insert_row_problem_column(
            worksheet,
            target_column,
            build_row_problem_summaries(problem_entries),
        )

    write_term_sheet(
        workbook,
        worksheet.title,
        sorted_output_term_pairs(output_term_mapping.values()),
    )

    sorted_problem_entries = sorted(
        problem_entries,
        key=lambda entry: (
            term_source_priority(entry.term_source),
            entry.problem_source_term == "",
            normalize_text(entry.problem_source_term, case_sensitive=False),
            entry.row_index,
        ),
    )
    write_problem_sheet(
        workbook,
        worksheet.title,
        target_column,
        sorted_problem_entries,
    )

    delete_legacy_term_sheets(workbook)

    if save_output:
        workbook.save(output_path)

    result = (
        worksheet.title,
        source_column,
        target_column,
        output_path,
        len(output_term_mapping),
        len(problem_entries),
    )
    if owns_workbook:
        workbook.close()
    return result


def process_workbook(
    *,
    workbook,
    input_file: str | Path,
    output_path: Path,
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
    include_row_problem_column: bool = True,
) -> tuple[str, str, str, Path, int, int]:
    """Run terminology checks against an already-open workbook without saving it."""
    return process_excel(
        input_file=input_file,
        source_column=source_column,
        target_column=target_column,
        sheet=sheet,
        start_row=start_row,
        mark_styles=mark_styles,
        mark_style=mark_style,
        exclusion_patterns=exclusion_patterns,
        exclusion_config_file=exclusion_config_file,
        history_tb_file=history_tb_file,
        history_sheet=history_sheet,
        history_source_column=history_source_column,
        history_target_column=history_target_column,
        history_start_row=history_start_row,
        output_file=output_path,
        include_row_problem_column=include_row_problem_column,
        workbook=workbook,
        save_output=False,
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
    mark_description = "、".join(args.mark_style) if args.mark_style else "未选择（仅历史 TB）"
    print(f"mark 类型: {mark_description}")
    if args.history_tb:
        print(f"历史 TB: {Path(args.history_tb).expanduser().resolve()}")
    print(f"术语表条目数: {term_count}")
    print(f"问题条数: {problem_count}")
    print(f"输出文件: {output_path}")


if __name__ == "__main__":
    main()
