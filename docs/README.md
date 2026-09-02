# QAtools 文档路由

本索引面向用户、自动化和 AI agent。按任务读取最少的权威文档，不需要遍历整个
仓库。

| 任务 | 读取 |
|---|---|
| 首次使用或了解能力 | [项目 README](../README.md) |
| 按 GUI 运行一键质量检查 | [一键质量检查 GUI 使用指南](qa-workflow-guide/index.html) |
| 调用 CLI、编写脚本 | [CLI 使用指南](cli-usage.md) |
| 运行一键检查、理解统一报告 | [workflow README](../tools/workflow/README.md) |
| 导出或回填 Strings | [PhraseLoom README](../phraseloom/README.md) |
| 修改代码或文档 | [AGENTS.md](../AGENTS.md) |

## 规则文档

- 术语与 Tag：[术语检查](../tools/term_pair_checker/README.md)、
  [Tag / Placeholder](../tools/tag_placeholder_checker/README.md)、
  [共享规则](term-and-tag-check-rules.md)
- 双向文本一致性：[同 Source 不同 Target](../tools/source_consistency_checker/README.md)、
  [同 Target 不同 Source](../tools/target_consistency_checker/README.md)
- 内容保真：[数字与 URL](../tools/content_fidelity_checker/README.md)、
  [换行数量](../tools/line_break_checker/README.md)
- Target 文本质量：[Target 中文](../tools/chinese_target_checker/README.md)、
  [Target 文本规范](../tools/target_text_checker/README.md)
- 其他工具：[法语 NBSP](../tools/french_nbsp_restorer/README.md)、
  [Batch 拆分与复原](../tools/excel_batcher/README.md)、
  [合并表格](../tools/excel_merger/README.md)、
  [Xbench QA 转换](../tools/xbench_report_transformer/README.md)

平台说明：[Windows 安装版](../packaging/README-Windows.txt)、
[macOS Finder 工作流](macos-finder-workflow.md)。

## 文档契约

- 当前行为以代码和测试为准，其次是当前文档。
- 一个事实只维护在一个位置：调用参数在 CLI 手册，业务规则在工具 README，
  开发约束在 `AGENTS.md`。
- 根 README 只保留快速入口和能力地图，不复制完整参数或规则。
- `archive/` 只保存历史背景，不能作为当前实现依据。

更新入口、CLI、业务规则或开发约束时，分别修改根 README、`cli-usage.md`、
对应工具 README 或 `AGENTS.md`。
