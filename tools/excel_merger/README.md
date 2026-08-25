# 合并表格

该工具递归扫描一个目录，把每个 `.xlsx` / `.xlsm` 工作簿当前活动的工作表
依次合并到一个新的 `MergedData` 工作表。

## 合并规则

- 在输出首列添加 `SourceFile`，数据行记录来源文件名。
- 默认保留第一份非空活动工作表的第一行；后续工作表跳过第一行。
- 勾选“保留每个文件的表头”后，每份工作表的第一行都会保留。
- 扫描包含所有子目录，忽略 Excel 临时文件 `~$*`。
- 支持 `.xlsx` 和 `.xlsm`；发现 `.xls` / `.xlsb` 时会统计并跳过。
- 某个文件损坏或无法读取时继续处理其他文件，并在输出旁生成错误日志。
- 输出内容统一写为文本值，不复制源工作表的样式、公式、批注或图片。

默认输出在输入目录的上一级，文件名格式为：

```text
<目录名>_merged_active_sheet_<YYYYMMDD_HHMMSS>.xlsx
```

## 使用

GUI：打开 QAtools，在“其他”区域选择“合并表格”。

CLI：

```bash
qatools merge-sheets ./excel-files
qatools merge-sheets ./excel-files --keep-all-headers -o ./merged.xlsx
```

原 `mergesSheets` 仓库的参数名也可继续使用：

```bash
qatools merge-sheets --folder-path ./excel-files --output-path ./merged.xlsx
```
