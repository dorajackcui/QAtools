# 同 Source 不同 Target

统一 CLI：`qatools consistency-check --help`。

检查 Excel 中归一化后相同的 `source` 是否对应多个不同 `target`，用于定位同源文本译法不一致问题。

一键质量检查 GUI 和统一报告使用直观名称“同 Source 不同 Target”；现有
`qatools consistency-check` 命令保持兼容。

## 规则

- source 分组与 target 版本比较遵循[共享归一化规则](../../docs/consistency-normalization.md)，报告保留各行原文
- 归一化后为空的 source 跳过
- 空 target 会作为一种译文参与比较；同一 source 同时出现空 target 和非空 target 时会报错
- target 归一化后相同的重复 source 不报错

## CLI

调用示例、参数和兼容脚本集中在 [CLI 手册](../../docs/cli-usage.md)。

不指定 `-o/--output` 时，默认生成 `source_consistency_check_<原文件名>`，不会覆盖输入文件。

## GUI

```bash
python3 tools/source_consistency_checker/check_source_consistency_gui.py
```

也可以运行 `qatools gui`，在“一键质量检查”中选择“同 Source 不同 Target”。
旧脚本同样打开 PySide6 QA 页，仅预选此检查，并输出统一 QA 报告。

## 输出

结果工作簿保留原工作表，并新增 `同源译文不一致` 工作表。不一致组中的每个原始行都会单独列出，并包含：

- `行号 / source原文 / target原文 / 问题描述`
- target 版本数
- 同组全部行号

问题描述使用 `2 种译法：1: Astro Boy；2: Atom`，按归一化后的不同译法去重，
按首次出现顺序编号，每种展示首次出现的原文摘要；归一化后为空时显示 `[空译文]`。
长文本和过多版本按[报告描述长度规则](../workflow/README.md#描述长度)截断并标注省略。
独立问题表的版本数和完整同组行号列保持不变。
