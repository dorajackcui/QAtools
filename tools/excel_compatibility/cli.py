import argparse

from tools.batch_cli import add_output_options, execute


def main(argv=None):
    parser = argparse.ArgumentParser(prog="qatools compatibility", description="用桌面 Excel 按原格式重存；默认原位更新")
    parser.add_argument("folder_path", help="工作簿目录，递归处理")
    add_output_options(parser)
    options = vars(parser.parse_args(argv))
    from .processor import resave_workbooks
    return execute(resave_workbooks, **options)


if __name__ == "__main__":
    raise SystemExit(main())
