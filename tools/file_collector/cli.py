"""Noninteractive filename-list extraction; help needs only the standard library."""
import argparse
import sys

from tools.batch_cli import execute


def main(argv=None):
    parser = argparse.ArgumentParser(prog="qatools collect-files", description="按文件名清单提取 Excel 到新目录")
    parser.add_argument("source_dir", help="递归扫描的来源目录")
    parser.add_argument("--names-file", required=True, help="UTF-8 TXT 文件名清单，默认按换行或 Tab 分隔")
    parser.add_argument("-o", "--output-dir", required=True, help="必填；尚不存在且与来源不重叠的新目录")
    parser.add_argument("--preserve-tree", action="store_true", help="保留来源相对路径；默认平铺")
    parser.add_argument("--comma-separated", action="store_true", help="额外按中英文逗号分隔文件名")
    parser.add_argument("--dry-run", action="store_true", help="只预览匹配，不创建目录或复制文件")
    parser.add_argument("--quiet", action="store_true", help="仅输出摘要和问题，不输出逐文件成功日志")
    args = parser.parse_args(argv)
    from .collector import build_copy_plan, execute_copy_plan, read_names_file
    try:
        plan = build_copy_plan(args.source_dir, read_names_file(args.names_file), output_dir=args.output_dir,
                               preserve_tree=args.preserve_tree, comma_separated=args.comma_separated)
    except Exception as exc:
        print(f"执行失败：{exc}", file=sys.stderr)
        return 1
    print(plan.describe())
    if args.dry_run:
        labels = {"ready": "可复制", "conflict": "冲突", "unavailable": "不可读"}
        for entry in plan.entries:
            if not args.quiet or entry.status != "ready":
                print(f"[{labels[entry.status]}] {' / '.join(entry.names)} | {entry.source} → {entry.destination}"
                      + (f" | {entry.message}" if entry.message else ""))
        for name in plan.not_found:
            print(f"[未找到] {name}")
        print(f"仅预览，未写入：{plan.output_dir}")
        return 3 if plan.not_found or plan.ready_count != len(plan.entries) else 0
    def collect(*, log_callback):
        summary = execute_copy_plan(plan, log_callback=log_callback)
        if args.quiet:
            from tools.operation_logs import log_result
            for result in summary.results:
                if result.status in {"failed", "skipped"}:
                    log_result(lambda message: print(message, file=sys.stderr), result)
        return summary

    return execute(collect, quiet=args.quiet)


if __name__ == "__main__":
    raise SystemExit(main())
