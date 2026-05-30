# Target 中文检查

检查 Excel `target` 列是否包含中文字符，并把结果写入新的 Excel 文件。

## 功能

- 扫描指定工作表的 `target` 列
- 默认在 `target` 右侧一列写入检查结果
- 含中文字符的行标记为 `含中文`
- 不含中文字符的行留空
- 可通过 `-r/--result-column` 指定结果列
- 可通过 `--problem-sheet` 新增 `中文检查问题` 工作表
- 不覆盖原文件

## CLI

默认写入 target 右侧一列：

```bash
python3 tools/chinese_target_checker/check_chinese_target.py input.xlsx \
  -s Sheet1 \
  -t B \
  --start-row 2 \
  -o output_chinese_target_checked.xlsx
```

指定结果列并生成问题工作表：

```bash
python3 tools/chinese_target_checker/check_chinese_target.py input.xlsx \
  -s Sheet1 \
  -t B \
  -r C \
  --start-row 2 \
  --problem-sheet \
  -o output_chinese_target_checked.xlsx
```

## GUI

独立 GUI：

```bash
python3 tools/chinese_target_checker/check_chinese_target_gui.py
```

也可以使用统一入口：

```bash
python3 toolshub_gui.py
```

然后选择 `Target中文检查` 标签页。

## 输出

- 原数据工作表会新增或覆盖结果列，表头为 `中文检查`
- 命中行写入 `含中文`
- 启用 `--problem-sheet` 后会新增 `中文检查问题`，包含：
  - `行号`
  - `target文本`
  - `中文字符`
