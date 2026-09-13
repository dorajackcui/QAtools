# 同 Target 不同 Source

检查 Excel 中归一化后相同的非空 `target` 是否对应多个不同 `source`。该规则用于发现
可能被错误复用的译文，但多个 source 合理共用一个译文的情况也很常见，因此在
一键质量检查 GUI 中默认关闭。

## 规则

- target 分组与 source 版本比较遵循[共享归一化规则](../../docs/consistency-normalization.md)，报告保留各行原文。
- 归一化后为空的 target 跳过。
- 归一化后为空的 source 会作为一种原文参与比较。
- 相同 target 对应的 source 归一化后相同时不报告。
- 问题表会列出组内每一行、source 版本数和同组行号。

该检查当前通过 `qatools qa --check target-consistency` 或一键质量检查 GUI 使用。

问题描述使用 `3 种原文（第 2、6、9 行）`，保留版本数和完整同组行号。
