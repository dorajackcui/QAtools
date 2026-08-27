# Xbench QA Report 转换

统一 CLI：`qatools xbench --help`。下文中的脚本路径作为兼容入口保留。

## 用途

把 ApSIC Xbench 导出的 QA Report 转换成更适合修改和筛选的五列表格：

```text
文件名, key, source, target, QA问题
```

工具不会修改原始 Xbench 报告，默认生成新的结果 Excel。

## 聚类规则

- `Metadata` 两行或更多：第一行作为 `key`，第二行作为 `文件名`。
- `Metadata` 只有一行且像文件名：`key` 留空，`文件名` 使用这一行，按 `文件名 + source` 聚类。
- `Metadata` 只有一行且不像文件名：这一行作为 `key`，按 `key` 聚类。
- `Metadata` 为空：按 `source` 聚类。

同一组内多个 QA 问题会写入同一个 `QA问题` 单元格，并用中文分号 `；` 连接。

## CLI

```bash
python3 tools/xbench_report_transformer/transform_xbench_report.py Xbench_QA_Report.xlsx
```

指定工作表和输出文件：

```bash
python3 tools/xbench_report_transformer/transform_xbench_report.py Xbench_QA_Report.xlsx \
  -s "Xbench QA" \
  -o xbench_flat.xlsx
```

默认输出文件名为 `xbench_transform_<原文件名>`。

## GUI

单独启动：

```bash
python3 tools/xbench_report_transformer/transform_xbench_report_gui.py
```

也可以从统一入口打开：

```bash
qatools gui
```

GUI 会自动读取工作表列表，输出文件沿用默认命名规则。
