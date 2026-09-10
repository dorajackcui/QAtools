"""Shared presentation for noninteractive file operations; no GUI/Excel imports."""
from __future__ import annotations

import argparse
import sys


def column(value: str) -> str:
    normalized = value.strip().upper()
    number = 0
    for character in normalized:
        if not "A" <= character <= "Z":
            raise argparse.ArgumentTypeError("列必须为 A–XFD 的 Excel 列字母。")
        number = number * 26 + ord(character) - ord("A") + 1
    if not 1 <= number <= 16384:
        raise argparse.ArgumentTypeError("列必须为 A–XFD 的 Excel 列字母。")
    return normalized


def header_rows(value: str) -> int:
    try:
        number = int(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("表头行数必须是整数。") from exc
    if not 0 <= number < 1048576:
        raise argparse.ArgumentTypeError("表头行数必须在 0–1048575 之间。")
    return number


def column_count(value: str) -> int:
    number = header_rows(value)
    if not 1 <= number <= 16384:
        raise argparse.ArgumentTypeError("更新列数必须在 1–16384 之间。")
    return number


def add_output_options(parser):
    parser.add_argument("-o", "--output-dir", help="可选；省略或留空时在输入位置操作，填写时必须是尚不存在的新目录")
    parser.add_argument("--quiet", action="store_true", help="仅输出最终摘要和异常，不输出逐文件运行日志")


def execute(operation, *, quiet=False, **kwargs) -> int:
    def log(message):
        print(message, file=sys.stderr, flush=True)

    try:
        summary = operation(log_callback=None if quiet else log, **kwargs)
    except Exception as exc:
        # COM startup and filesystem exceptions should also have a predictable exit.
        log(f"执行失败：{exc}")
        return 1
    print(f"输出目录: {summary.output_dir}")
    print(f"成功: {summary.succeeded_files}；失败: {summary.failed_files}；跳过: {summary.skipped_files}")
    details = getattr(summary, "details", {})
    updates = getattr(summary, "updated_cells", details.get("updated_cells"))
    if updates is not None:
        print(f"实际更新单元格: {updates}")
    for name in ("master_output", "excel_report", "backup_dir"):
        if details.get(name):
            print(f"{name}: {details[name]}")
    if quiet:
        for warning in summary.warnings:
            log(f"警告：{warning}")
    if "master_output" in details and not details["master_output"]:
        return 1  # Reading small tables does not mean Master was saved successfully.
    if summary.failed_files and not summary.succeeded_files:
        return 1
    if summary.failed_files or summary.skipped_files or summary.warnings:
        return 3
    return 0
