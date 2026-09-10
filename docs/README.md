# QAtools 文档路由

本索引面向用户、自动化和 AI agent。按任务读取最少的权威文档，不需要遍历整个
仓库。

| 任务 | 读取 |
|---|---|
| 首次使用或了解能力 | [项目 README](../README.md) |
| 按 GUI 操作质量检查、内容同步和文件工具 | [GUI 使用指南](qa-workflow-guide/index.html)；[网页维护与验证](qa-workflow-guide/README.md) |
| 调用 CLI、编写脚本 | [CLI 使用指南](cli-usage.md) |
| 运行一键检查、理解统一报告 | [workflow README](../tools/workflow/README.md) |
| 导出或回填 Strings | [PhraseLoom README](../phraseloom/README.md) |
| 修改代码或文档 | [AGENTS.md](../AGENTS.md)、[仓库地图](repository-map.md) |
| 扩展工具、查看大文件评估 | [仓库地图](repository-map.md#扩展一个工具) |
| 查询未完成工作 | [待办](backlog.md) |
| GUI 表头别名、布局与状态 | [GUI 约定](gui-conventions.md) |
| 使用 Master ↔ 小表同步 | [内容同步](../tools/content_sync/README.md) |
| 清空、插入或删除 Excel 列 | [列操作](../tools/column_tools/README.md) |
| 用 Excel 重新保存工作簿 | [兼容性重存](../tools/excel_compatibility/README.md) |
| 按文件名替换整个工作簿 | [同名文件替换](../tools/deep_replace/README.md) |
| 汇总待翻译字数/词数和行数 | [未翻译统计](../tools/untranslated_stats/README.md) |
| 查看六个新工具的实时处理详情 | [运行日志](tool-logs.md) |
| 追溯导入来源 | [归档索引](archive/README.md) |

## 业务规则

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

## 文档维护

根 [AGENTS.md](../AGENTS.md#文档职责)规定各文档的权威范围。当前规则写入工具 README，
调用写入 CLI 手册，待办写入 backlog；已完成决策只保留来源和约束，不维护第二套业务说明。

提交前运行 `python scripts/check_docs.py` 检查文档链接和 CLI 命令目录；
GUI/CLI/业务变更仍需运行对应测试及全量回归。
