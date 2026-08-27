# 同 Target 不同 Source

检查 Excel 中完全相同的非空 `target` 是否对应多个不同 `source`。该规则用于发现
可能被错误复用的译文，但多个 source 合理共用一个译文的情况也很常见，因此在
一键质量检查 GUI 中默认关闭。

## 规则

- 按 target 单元格文本精确分组，不忽略大小写、空格或换行差异。
- 空 target 和只包含空白字符的 target 跳过。
- source 按单元格文本精确比较；空 source 会作为一种原文参与比较。
- 相同 target 对应的 source 完全相同时不报告。
- 问题表会列出组内每一行、source 版本数和同组行号。

该检查当前通过 `qatools qa --check target-consistency` 或一键质量检查 GUI 使用。
