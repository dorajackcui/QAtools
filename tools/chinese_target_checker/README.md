# Target 中文检查

检查 Excel `target` 列是否包含中文字符或中文/全角标点，并在原数据表写入检查结果。

## 功能

- 扫描指定工作表的 `target` 列
- 默认在 `target` 右侧新增一列写入检查结果，原右侧列会后移；如果右侧已经是 `中文检查` 列，则直接复用
- 含中文字符或中文/全角标点的行标记为 `含中文`
- 不含中文字符和中文/全角标点的行留空
- 可通过 `-r/--result-column` 指定结果列
- 不会额外生成问题工作表
- 默认直接修改原文件；指定 `-o/--output` 时才另存为新文件

## CLI

默认在原文件的 target 右侧新增一列：

```bash
python3 tools/chinese_target_checker/check_chinese_target.py input.xlsx \
  -s Sheet1 \
  -t B \
  --start-row 2
```

指定结果列并另存为新文件：

```bash
python3 tools/chinese_target_checker/check_chinese_target.py input.xlsx \
  -s Sheet1 \
  -t B \
  -r C \
  --start-row 2 \
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
- 检查范围包含汉字、CJK 标点、全角标点和常见中文排版符号，例如 `【】（）`、`，。！？`、`《》“”‘’—…·`
- 不把普通 ASCII 标点或全角英数单独算作中文
- 不会额外生成问题工作表；如果工作簿里已有旧的 `中文检查问题` 工作表，运行时会移除
