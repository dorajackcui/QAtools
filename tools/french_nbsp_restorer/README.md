# 法语 NBSP 恢复

统一 CLI：`qatools french-nbsp --help`。下文中的脚本路径作为兼容入口保留。

这个工具用于恢复 Excel target 列中的法语 non-breaking space（NBSP），输出新的 Excel 文件，不覆盖原文件。

## 规则

- `;`、`:`、`?`、`!`、`%` 前恢复 NBSP。
- `«` 后和 `»` 前恢复 NBSP。
- 已有普通空格、NBSP 或 narrow NBSP 会统一为 NBSP。
- URL 内标点和 `12:30` 这类时间冒号不会被改写。

## CLI

直接修复 target 列：

```bash
python3 tools/french_nbsp_restorer/restore_french_nbsp.py input.xlsx \
  -s Sheet1 \
  -t B \
  --start-row 2 \
  -o output_french_nbsp_restored.xlsx
```

写入修复后的完整结果列：

```bash
python3 tools/french_nbsp_restorer/restore_french_nbsp.py input.xlsx \
  -s Sheet1 \
  -t B \
  -r C \
  --start-row 2 \
  -o output_french_nbsp_restored.xlsx
```

如果指定结果列，不需要修复的 target 也会复制到结果列。

## GUI

```bash
python3 tools/french_nbsp_restorer/restore_french_nbsp_gui.py
```

也可以通过统一入口打开：

```bash
qatools gui
```

统一 GUI 会自动识别工作表首行的 Target 列；非标准表头可通过
[GUI 表头别名设置](../../README.md#自定义表头识别)配置。
