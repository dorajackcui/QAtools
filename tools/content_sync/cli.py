"""Noninteractive Master ↔ target adapters; processors own every business rule."""
import argparse

from tools.batch_cli import add_output_options, column, column_count, execute, header_rows


def build_parser():
    parser = argparse.ArgumentParser(prog="qatools content-sync", description="按 Key + 原文同步 Master 与小表；默认原位更新")
    directions = parser.add_subparsers(dest="direction", required=True)
    for direction, description in (("master-to-target", "将 Master 内容同步到小表"),
                                   ("target-to-master", "将小表译文汇总回填 Master")):
        action = directions.add_parser(direction, help=description, description=description)
        action.add_argument("master_file", help="Master .xlsx/.xlsm 路径")
        action.add_argument("target_dir", help="小表目录，递归处理")
        add_output_options(action)
        for side, defaults in (("master", ("B", "C", "D")), ("target", ("A", "B", "C"))):
            for role, default in zip(("key", "source", "content"), defaults):
                action.add_argument(f"--{side}-{role}-column", type=column, default=default,
                                    help=f"{side} {role} 列（默认 {default}）")
            action.add_argument(f"--{side}-sheet", help="工作表名；省略使用活动工作表")
            action.add_argument(f"--{side}-header-rows", type=header_rows, default=1, help="表头行数（默认 1）")
        action.add_argument("--fill-blank-only", action="store_true", help="仅填充空白目标")
        action.add_argument("--allow-blank-write", action="store_true", help="允许空白来源写入；独立于仅填空选项")
        action.add_argument("--workers", type=int, choices=range(1, 5),
                            default=2 if direction == "master-to-target" else 1,
                            help="小表任务数 1–4；默认正向 2、反向 1；Excel 重存时串行")
        if direction == "master-to-target":
            action.add_argument("--column-count", type=column_count, default=1, help="连续更新列数（默认 1）")
            action.add_argument("--compatibility-resave", action="store_true", help="同步后用桌面 Excel 重存已更新的小表")
    return parser


def main(argv=None):
    options = vars(build_parser().parse_args(argv))
    direction = options.pop("direction")
    # Import only after parsing: --help requires neither openpyxl, Qt nor COM.
    from .master_to_target import ColumnMapping, sync_master_to_targets
    from .target_to_master import sync_targets_to_master
    for side in ("master", "target"):
        options[f"{side}_columns"] = ColumnMapping(*(options.pop(f"{side}_{role}_column")
                                                     for role in ("key", "source", "content")))
    operation = sync_master_to_targets if direction == "master-to-target" else sync_targets_to_master
    return execute(operation, **options)


if __name__ == "__main__":
    raise SystemExit(main())
