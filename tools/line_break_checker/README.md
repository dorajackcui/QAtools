# 换行数量检查

统一 CLI：`qatools line-break-check --help`。下文中的脚本路径作为兼容入口保留。

逐行比较 Excel `source` / `target` 单元格中的真实换行数量，定位肉眼不易发现的换行缺失或多出问题。

## 规则

- `LF`（`\n`）、`CRLF`（`\r\n`）和 `CR`（`\r`）都按真实换行处理
- `CRLF` 按一个换行计算
- 文本中的两个普通字符 `\` 和 `n` 不算真实换行
- 非文本或空单元格按 0 个换行计算
- 仅当同一行 source / target 的换行数不同时写入问题表

## CLI

```bash
python3 tools/line_break_checker/check_line_breaks.py input.xlsx \
  -s Sheet1 \
  -c A \
  -t B \
  --start-row 2 \
  -o line_break_check_output.xlsx
```

不指定 `-o/--output` 时，默认生成 `line_break_check_<原文件名>`，不会覆盖输入文件。

## GUI

```bash
python3 tools/line_break_checker/check_line_breaks_gui.py
```

也可以运行 `qatools gui`，在“一键质量检查”中选择“换行数量”。

## 输出

结果工作簿保留原工作表，并新增 `换行数量问题` 工作表，包含：

- `行号 / source原文 / target原文 / 问题描述`
- source / target 换行数及数量差（target 减 source）
