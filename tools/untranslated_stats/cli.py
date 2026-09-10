import argparse

from tools.batch_cli import add_output_options, column, execute, header_rows


def main(argv=None):
    parser = argparse.ArgumentParser(prog="qatools untranslated-stats", description="统计未翻译量；默认将统计表写入所选目录")
    parser.add_argument("target_dir", help="工作簿目录，递归处理")
    parser.add_argument("-c", "--source-column", type=column, default="B", help="原文列（默认 B）")
    parser.add_argument("-t", "--target-column", type=column, default="C", help="译文列（默认 C）")
    parser.add_argument("--sheet", help="工作表名；省略使用活动工作表")
    parser.add_argument("--header-rows", type=header_rows, default=1, help="表头行数（默认 1）")
    parser.add_argument("--mode", choices=("chinese_chars", "english_words"), default="chinese_chars", help="计数口径（默认 chinese_chars）")
    add_output_options(parser)
    options = vars(parser.parse_args(argv))
    from .stats import untranslated_stats
    return execute(untranslated_stats, **options)


if __name__ == "__main__":
    raise SystemExit(main())
