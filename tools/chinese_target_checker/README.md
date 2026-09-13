# Target 中文检查

统一 CLI：`qatools chinese-check --help`。

检查 Excel `target` 列是否包含中文字符或中文/全角标点，并输出独立问题表。

## 功能

- 扫描指定工作表的 `source` / `target` 列
- 含中文字符或中文/全角标点的 target 行写入 `Target中文问题`
- 原数据工作表不新增标记列，也不修改原文
- 默认生成 `target_chinese_check_<原文件名>`；可通过 `-o/--output` 指定输出文件

## CLI

调用示例、参数和兼容脚本集中在 [CLI 手册](../../docs/cli-usage.md)。

## GUI

旧 GUI 启动脚本（转接到 PySide6 QA 页，仅预选 Target 中文，输出统一 QA 报告）：

```bash
python3 tools/chinese_target_checker/check_chinese_target_gui.py
```

也可以使用统一入口：

```bash
qatools gui
```

然后在“一键质量检查”中选择 `Target 中文`。

## 输出

- 原数据工作表保持不变
- `Target中文问题` 前四列为 `行号 / source原文 / target原文 / 问题描述`，第五列为 `命中字符`
- 检查范围包含汉字、CJK 标点、全角标点和常见中文排版符号，例如 `【】（）`、`，。！？`、`《》“”‘’·`；单字符 `—` 和 `…` 放行
- 不把普通 ASCII 标点或全角英数单独算作中文
- 已有的同名问题表会重建，旧版 `中文检查问题` 工作表会移除

问题描述直接展示命中字符，例如 `命中：“中”“文”“，”`；字符顺序、去重方式
与“命中字符”列一致。
