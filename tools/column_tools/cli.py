import argparse

from tools.batch_cli import add_output_options, column, execute, header_rows


def main(argv=None):
    parser = argparse.ArgumentParser(prog="qatools columns", description="用桌面 Excel 执行列操作；默认原位更新")
    actions = parser.add_subparsers(dest="action", required=True)
    for name, description in (("clear", "清空表头之后的内容"), ("insert", "插入整列"), ("delete", "删除整列")):
        action = actions.add_parser(name, help=description, description=description)
        action.add_argument("folder_path", help="工作簿目录，递归处理")
        action.add_argument("--column", type=column, default="C", help="操作列（默认 C）")
        action.add_argument("--sheet", help="工作表名；省略使用活动工作表")
        if name == "clear":
            action.add_argument("--header-rows", type=header_rows, default=1, help="保留的表头行数（默认 1）")
        if name == "insert":
            action.add_argument("--inserted-header", default="Translation", help="新列第 1 行标题（默认 Translation）")
        add_output_options(action)
    options = vars(parser.parse_args(argv))
    from .processor import operate_columns
    return execute(operate_columns, **options)


if __name__ == "__main__":
    raise SystemExit(main())
