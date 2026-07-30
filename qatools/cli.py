from __future__ import annotations

from dataclasses import dataclass
import importlib
from importlib import metadata
import sys
from typing import Literal, Sequence

from . import __version__


InvocationMode = Literal["argv", "legacy"]


@dataclass(frozen=True)
class CommandSpec:
    name: str
    summary: str
    module: str
    mode: InvocationMode
    aliases: tuple[str, ...] = ()
    legacy_entry: str | None = None


COMMANDS: tuple[CommandSpec, ...] = (
    CommandSpec(
        name="gui",
        summary="打开统一 Toolshub 图形界面",
        module="toolshub_gui",
        mode="argv",
        legacy_entry="python toolshub_gui.py",
    ),
    CommandSpec(
        name="qa",
        summary="一次运行多项质量检查并生成统一报告",
        module="tools.workflow.cli",
        mode="argv",
        aliases=("workflow",),
    ),
    CommandSpec(
        name="phraseloom",
        summary="导出 Strings，并在翻译完成后回填原始 Excel",
        module="phraseloom.cli",
        mode="argv",
        aliases=("strings",),
        legacy_entry="phraseloom",
    ),
    CommandSpec(
        name="term-check",
        summary="检查术语 mark 和历史 TB",
        module="tools.term_pair_checker.extract_terms_from_excel",
        mode="legacy",
        legacy_entry="python tools/term_pair_checker/extract_terms_from_excel.py",
    ),
    CommandSpec(
        name="tag-check",
        summary="检查 Tag、Placeholder、换行标记和 memoQ Tag",
        module="tools.tag_placeholder_checker.check_tags_and_placeholders",
        mode="legacy",
        legacy_entry="python tools/tag_placeholder_checker/check_tags_and_placeholders.py",
    ),
    CommandSpec(
        name="line-break-check",
        summary="检查 source 和 target 的真实换行数量",
        module="tools.line_break_checker.check_line_breaks",
        mode="legacy",
        legacy_entry="python tools/line_break_checker/check_line_breaks.py",
    ),
    CommandSpec(
        name="consistency-check",
        summary="检查相同 source 的 target 是否一致",
        module="tools.source_consistency_checker.check_source_consistency",
        mode="legacy",
        aliases=("source-consistency",),
        legacy_entry="python tools/source_consistency_checker/check_source_consistency.py",
    ),
    CommandSpec(
        name="chinese-check",
        summary="检查 target 中的中文字符和中文标点",
        module="tools.chinese_target_checker.check_chinese_target",
        mode="legacy",
        aliases=("target-chinese",),
        legacy_entry="python tools/chinese_target_checker/check_chinese_target.py",
    ),
    CommandSpec(
        name="split-lines",
        summary="把单元格多行文本连续拆写到结果列",
        module="tools.excel_line_splitter.split_excel_lines",
        mode="legacy",
        legacy_entry="python tools/excel_line_splitter/split_excel_lines.py",
    ),
    CommandSpec(
        name="french-nbsp",
        summary="恢复法语标点所需的 NBSP",
        module="tools.french_nbsp_restorer.restore_french_nbsp",
        mode="legacy",
        legacy_entry="python tools/french_nbsp_restorer/restore_french_nbsp.py",
    ),
    CommandSpec(
        name="xbench",
        summary="把 Xbench QA Report 转换为行级问题表",
        module="tools.xbench_report_transformer.transform_xbench_report",
        mode="legacy",
        aliases=("xbench-transform",),
        legacy_entry="python tools/xbench_report_transformer/transform_xbench_report.py",
    ),
)


def command_map() -> dict[str, CommandSpec]:
    mapping: dict[str, CommandSpec] = {}
    for command in COMMANDS:
        for name in (command.name, *command.aliases):
            if name in mapping:
                raise ValueError(f"重复的 QAtools CLI 命令或别名: {name}")
            mapping[name] = command
    return mapping


def package_version() -> str:
    try:
        return metadata.version("qatools")
    except metadata.PackageNotFoundError:
        return __version__


def format_command_list() -> str:
    width = max(len(command.name) for command in COMMANDS)
    lines = []
    for command in COMMANDS:
        alias_text = (
            f"（别名: {', '.join(command.aliases)}）"
            if command.aliases
            else ""
        )
        lines.append(
            f"  {command.name.ljust(width)}  {command.summary}{alias_text}"
        )
    return "\n".join(lines)


def format_help() -> str:
    return "\n".join(
        (
            "QAtools 统一命令行",
            "",
            "用法:",
            "  qatools <命令> [参数]",
            "  python -m qatools <命令> [参数]",
            "",
            "命令:",
            format_command_list(),
            "",
            "辅助命令:",
            "  qatools list            列出可用命令",
            "  qatools help <命令>     查看命令的完整参数",
            "  qatools --version       显示版本",
            "",
            "示例:",
            "  qatools gui",
            "  qatools qa input.xlsx -c A -t B",
            "  qatools phraseloom export source.xlsx",
            "  qatools tag-check input.xlsx -c A -t B",
        )
    )


def _invoke(command: CommandSpec, arguments: Sequence[str]) -> int:
    module = importlib.import_module(command.module)
    entry = getattr(module, "main")
    if command.mode == "argv":
        result = entry(list(arguments))
        return int(result or 0)

    previous_argv = sys.argv[:]
    sys.argv = [f"qatools {command.name}", *arguments]
    try:
        result = entry()
    finally:
        sys.argv = previous_argv
    return int(result or 0)


def main(argv: Sequence[str] | None = None) -> int:
    arguments = list(sys.argv[1:] if argv is None else argv)
    if not arguments or arguments[0] in {"-h", "--help"}:
        print(format_help())
        return 0
    if arguments[0] in {"-V", "--version"}:
        print(package_version())
        return 0
    if arguments[0] == "list":
        print(format_command_list())
        return 0

    if arguments[0] == "help":
        if len(arguments) == 1:
            print(format_help())
            return 0
        arguments = [arguments[1], "--help", *arguments[2:]]

    invoked_name = arguments[0]
    command = command_map().get(invoked_name)
    if command is None:
        print(
            f"未知命令: {invoked_name!r}。使用 'qatools --help' 查看可用命令。",
            file=sys.stderr,
        )
        return 2

    try:
        return _invoke(command, arguments[1:])
    except SystemExit as exc:
        return int(exc.code or 0)
    except (FileNotFoundError, KeyError, ValueError) as exc:
        print(str(exc), file=sys.stderr)
        return 1


__all__ = [
    "COMMANDS",
    "CommandSpec",
    "command_map",
    "format_command_list",
    "format_help",
    "main",
    "package_version",
]
