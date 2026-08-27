# 同 Source 不同 Target

统一 CLI：`qatools consistency-check --help`。下文中的脚本路径作为兼容入口保留。

检查 Excel 中完全相同的 `source` 是否对应多个不同 `target`，用于定位同源文本译法不一致问题。

一键质量检查 GUI 和统一报告使用直观名称“同 Source 不同 Target”；现有
`qatools consistency-check` 命令保持兼容。

## 规则

- 按 source 单元格文本精确分组，不忽略大小写、空格或换行差异
- 空 source 和只包含空白字符的 source 跳过
- target 也按单元格文本精确比较
- 空 target 会作为一种译文参与比较；同一 source 同时出现空 target 和非空 target 时会报错
- target 完全相同的重复 source 不报错

## CLI

```bash
python3 tools/source_consistency_checker/check_source_consistency.py input.xlsx \
  -s Sheet1 \
  -c A \
  -t B \
  --start-row 2 \
  -o source_consistency_check_output.xlsx
```

不指定 `-o/--output` 时，默认生成 `source_consistency_check_<原文件名>`，不会覆盖输入文件。

## GUI

```bash
python3 tools/source_consistency_checker/check_source_consistency_gui.py
```

也可以运行 `qatools gui`，在“一键质量检查”中选择“同 Source 不同 Target”。

## 输出

结果工作簿保留原工作表，并新增 `同源译文不一致` 工作表。不一致组中的每个原始行都会单独列出，并包含：

- `行号 / source原文 / target原文 / 问题描述`
- target 版本数
- 同组全部行号
