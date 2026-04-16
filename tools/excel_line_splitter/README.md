# Excel 分行拆列工具

将 Excel 指定列中的单元格内容按回车拆开，并把拆出的内容连续写入结果列，输出为新的 Excel 文件。

## 规则

- 从指定开始行起遍历源列
- 单元格内容会按 `\r\n`、`\n`、`\r` 拆分
- 每一行会先去掉首尾空格
- 空行会被忽略；如果原文本末尾带回车，不会额外写出空结果
- 所有拆出的结果会从结果列开始行起连续堆叠写入
- 不会覆盖原文件，默认输出为 `<原文件名>_split_lines.xlsx`
- 写入前会先清空结果列从开始行起的旧值，避免残留

## 命令行

```bash
python3 tools/excel_line_splitter/split_excel_lines.py input.xlsx -c A -r B --start-row 2
```

常用参数：

- `-c, --source-column`：源列，例如 `A`
- `-r, --result-column`：结果列，例如 `B`
- `-s, --sheet`：工作表名称，可选
- `--start-row`：从第几行开始处理，默认 `2`
- `-o, --output`：输出文件路径，可选

## 图形界面

```bash
python3 tools/excel_line_splitter/split_excel_lines_gui.py
```

界面支持：

- 选择输入 Excel
- 选择输出 Excel
- 自动读取工作表列表，用下拉框选择工作表
- 输入源列、结果列、开始行
- 执行后提示写入条目数和输出文件路径
