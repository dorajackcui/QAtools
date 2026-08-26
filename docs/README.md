# QAtools 文档索引

## 当前文档

- [项目 README](../README.md)：快速开始、功能目录和仓库结构。
- [Agent 入口](../AGENTS.md)：代码地图、开发约束、验证与文档维护规则。
- [CLI 使用指南](cli-usage.md)：统一命令、自动化约定、示例和兼容入口。
- [PhraseLoom](../phraseloom/README.md)：Strings 导出、翻译与回填契约。
- [术语与 Tag 检查规则](term-and-tag-check-rules.md)：跨工具检查规则。
- [macOS Finder 快速操作](macos-finder-workflow.md)：右键执行 QA 或 NBSP 恢复。

单项工具的业务规则保存在对应的 `tools/<tool>/README.md`。从根目录
[功能目录](../README.md#功能目录) 可以直接进入各工具文档。

## 历史文档

[archive](archive/README.md) 保存已完成的设计、实施计划和迁移记录，仅用于
追溯背景，不代表当前行为。

## 维护规则

- 用户入口变化：更新根 README。
- CLI 参数变化：更新 `cli-usage.md`。
- 单项业务规则变化：更新对应工具 README。
- Agent 开发路径或约束变化：更新相应 `AGENTS.md`。
- 已实施且不再指导开发的方案：移入 `archive/`。
