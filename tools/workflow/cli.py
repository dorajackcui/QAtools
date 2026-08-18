from __future__ import annotations

import argparse
from collections.abc import Sequence

from tools.tag_placeholder_checker.check_tags_and_placeholders import (
    SUPPORTED_TOKEN_TYPES,
)
from tools.target_text_checker.check_target_text import (
    SUPPORTED_RULE_INPUTS as TARGET_TEXT_RULES,
)
from tools.term_pair_checker.term_marks import SUPPORTED_MARKS
from tools.workflow.workflow_runner import WorkflowSummary, run_workflow


CHECKS = ("term", "tag", "line-break", "consistency", "chinese", "text")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="qatools qa",
        description="执行所选本地化质量检查并生成统一 Excel 报告。",
    )
    parser.add_argument("input_file", help="输入 Excel 文件")
    parser.add_argument("-s", "--sheet", help="工作表名称；默认使用活动工作表")
    parser.add_argument(
        "-c",
        "--source-column",
        required=True,
        help="source 列，例如 A",
    )
    parser.add_argument(
        "-t",
        "--target-column",
        required=True,
        help="target 列，例如 B",
    )
    parser.add_argument("--start-row", type=int, default=2, help="开始行，默认 2")
    parser.add_argument(
        "-o",
        "--output",
        help="输出 Excel；默认生成 workflow_check_<原文件名>",
    )
    parser.add_argument(
        "--check",
        action="append",
        choices=CHECKS,
        help="只运行指定检查；可重复传入。不传时运行全部检查。",
    )

    term_mark_group = parser.add_mutually_exclusive_group()
    term_mark_group.add_argument(
        "--term-mark-style",
        action="append",
        choices=SUPPORTED_MARKS,
        metavar="MARK",
        help="术语 mark，可重复传入；支持 【】 和 []",
    )
    term_mark_group.add_argument(
        "--no-term-mark",
        action="store_true",
        help="不提取新术语，只使用 --history-tb 检查",
    )
    parser.add_argument("--history-tb", help="历史 TB Excel")
    parser.add_argument("--history-sheet", help="历史 TB 工作表")
    parser.add_argument("--history-source-column", help="历史 TB source 列")
    parser.add_argument("--history-target-column", help="历史 TB target 列")
    parser.add_argument(
        "--history-start-row",
        type=int,
        default=2,
        help="历史 TB 开始行，默认 2",
    )

    parser.add_argument(
        "--tag-token-type",
        action="append",
        choices=SUPPORTED_TOKEN_TYPES,
        help="Tag 检查类型，可重复传入",
    )
    parser.add_argument("--tag-angle-config", help="尖括号 Tag 过滤配置")
    parser.add_argument(
        "--text-rule",
        action="append",
        choices=TARGET_TEXT_RULES,
        help="Target 文本规范规则，可重复传入；默认运行全部规则",
    )
    return parser


def _selected_checks(values: Sequence[str] | None) -> set[str]:
    return set(values or CHECKS)


def _print_summary(summary: WorkflowSummary) -> None:
    print("质量检查完成。")
    print(f"工作表: {summary.worksheet_title}")
    print(f"source / target: {summary.source_column} / {summary.target_column}")
    if summary.ran_term_pair_check:
        print(
            "术语检查: "
            f"{summary.term_problem_rows} 个问题行，"
            f"{summary.term_problem_count} 个问题"
        )
    if summary.ran_tag_check:
        print(
            f"Tag 检查: {summary.tag_problem_rows} 个问题行，"
            f"{summary.tag_problem_count} 个问题"
        )
    if summary.ran_line_break_check:
        print(f"换行数量检查: {summary.line_break_problem_count} 个问题行")
    if summary.ran_source_consistency_check:
        print(
            "同源译文一致性: "
            f"{summary.source_consistency_problem_rows} 个问题行"
        )
    if summary.ran_chinese_target_check:
        print(f"Target 中文检查: {summary.chinese_target_problem_count} 个问题行")
    if summary.ran_target_text_check:
        print(
            "Target 文本规范检查: "
            f"{summary.target_text_problem_rows} 个问题行，"
            f"{summary.target_text_problem_count} 个问题"
        )
    print(f"输出文件: {summary.output_path}")


def main(argv: Sequence[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    checks = _selected_checks(args.check)
    term_mark_styles = () if args.no_term_mark else args.term_mark_style
    summary = run_workflow(
        input_file=args.input_file,
        source_column=args.source_column,
        target_column=args.target_column,
        output_file=args.output,
        sheet=args.sheet,
        start_row=args.start_row,
        run_term_pair_check="term" in checks,
        term_mark_styles=term_mark_styles,
        term_history_tb_file=args.history_tb,
        term_history_sheet=args.history_sheet,
        term_history_source_column=args.history_source_column,
        term_history_target_column=args.history_target_column,
        term_history_start_row=args.history_start_row,
        run_tag_check="tag" in checks,
        tag_token_types=args.tag_token_type,
        tag_angle_config_file=args.tag_angle_config,
        run_line_break_check="line-break" in checks,
        run_source_consistency_check="consistency" in checks,
        run_chinese_target_check="chinese" in checks,
        run_target_text_check="text" in checks,
        target_text_rules=args.text_rule,
    )
    _print_summary(summary)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
