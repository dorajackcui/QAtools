# QAtools 文档

## 使用

- [CLI 使用指南](cli-usage.md)：统一命令、自动化约定、示例和旧入口映射
- [PhraseLoom](../phraseloom/README.md)：Strings 导出、清洗、翻译与回填
- [术语与 Tag 检查规则](term-and-tag-check-rules.md)：检查边界和示例
- [macOS Finder 快速操作](macos-finder-workflow.md)：右键发送到一键质量检查

各单项工具的完整业务规则位于对应的 `tools/<tool>/README.md`。

## 架构与维护

- [PhraseLoom 合并记录](phraseloom-migration.md)
- [模块化设计](superpowers/specs/2026-05-31-project-modularization-design.md)
- [PhraseLoom Strings 设计](superpowers/specs/2026-07-30-strings-workflow-redesign.md)

历史计划和不再活跃的设计位于 [archive](archive/README.md)。

## 文档职责

- 根目录 `README.md`：快速开始、功能目录和文档导航。
- `docs/cli-usage.md`：唯一的统一 CLI 手册。
- 工具 README：该工具的业务规则、输入输出和专项说明。
- `docs/superpowers/specs/`：仍然有效的架构与格式设计。
- `docs/archive/`：保留历史背景，不作为当前行为依据。
