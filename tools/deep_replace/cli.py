import argparse

from tools.batch_cli import add_output_options, execute


def main(argv=None):
    parser = argparse.ArgumentParser(prog="qatools deep-replace", description="替换整个同名 Excel 文件；默认原位替换并先备份")
    parser.add_argument("source_dir", help="替换来源目录")
    parser.add_argument("target_dir", help="需要被替换的目标目录")
    add_output_options(parser)
    options = vars(parser.parse_args(argv))
    from .replacer import replace_files
    return execute(replace_files, **options)


if __name__ == "__main__":
    raise SystemExit(main())
