# 仓库地图与扩展路径

本页维护代码职责和依赖方向；业务规则以各工具 README 为准，开发约束见根 [AGENTS.md](../AGENTS.md)。

## 入口与调用链

```text
qatools / python -m qatools
  → qatools/cli.py 的 COMMANDS
  → 工具原生 CLI → 业务函数 → Excel / 文件系统

qatools gui / toolshub_gui.py
  → TOOL_GROUPS（导航名称和顺序）
  → tools/qt_pages.py 的 PAGE_FACTORIES（页面注册、兼容导入）
  → 各工具 qt_page.py（收集参数、显示结果）
  → AsyncPage 后台任务 → 同一业务函数
```

`tools/qt_pages.py` 不再实现页面。旧的页面类导入路径继续可用；新增页面实现放在工具包内，
不要把业务函数或设置表单重新堆回注册表。测试应 patch 实际使用依赖的模块。

## 代码位置

| 职责 | 入口 | 主要验证 |
|---|---|---|
| CLI 分发、别名、帮助 | [qatools/cli.py](../qatools/cli.py) | `test_qatools_cli.py` |
| 目录工具 CLI 公共参数、摘要与退出码 | [batch_cli.py](../tools/batch_cli.py)；工具包内 `cli.py` | `test_new_tools_cli.py` |
| 主窗口、导航、退出保护 | [toolshub_gui.py](../toolshub_gui.py) | `test_toolshub_gui.py` |
| 页面注册 | [tools/qt_pages.py](../tools/qt_pages.py) | 同上 |
| 已运行实例的页面跳转 | [qt_navigation.py](../tools/qt_navigation.py) | `test_gui_entries.py` |
| Qt 主题、输入控件、线程池 | [qt_gui_common.py](../tools/qt_gui_common.py) | GUI 回归 |
| Qt 页面布局与选择器 | [qt_page_helpers.py](../tools/qt_page_helpers.py) | GUI 回归 |
| 表头别名编辑页 | [qt_settings_page.py](../tools/qt_settings_page.py) | `test_header_aliases.py`、GUI 回归 |
| Excel 表头识别、别名 | [excel_metadata.py](../tools/excel_metadata.py)、[header_aliases.py](../tools/header_aliases.py) | 对应同名测试 |
| 旧 GUI 脚本转接与 Qt 表单行为 | 各工具 `*_gui.py` → `toolshub_gui.py` | `test_gui_entries.py`、`test_gui_excel_selection.py` |
| GUI 按钮到实际文件输出、双向流程和失败重试 | Qt 页面 → 真实后台任务 → 工作簿校验 | `test_gui_functional.py`；真实 Excel 用例沿用[COM 验证开关](../tools/excel_compatibility/README.md#原生操作的共同规则) |
| 工作簿编辑、保留 VBA | [excel_output.py](../tools/excel_output.py) | `test_excel_output_paths.py` |
| 新工具扫描、输出路径、原子复制 | [excel_file_ops.py](../tools/excel_file_ops.py) | `test_content_sync.py`、`test_momotools_utilities.py` |
| Excel COM 原生操作 | [excel_com.py](../tools/excel_com.py) | `test_momotools_utilities.py` |
| 日志文本与有界 GUI 队列 | [operation_logs.py](../tools/operation_logs.py)、[qt_operation_logs.py](../tools/qt_operation_logs.py) | `test_operation_logs.py` |
| 内容同步有界文件任务调度 | [content_sync/parallel.py](../tools/content_sync/parallel.py)；业务层负责汇总、日志与冲突顺序 | `test_content_sync_parallel.py`、`test_content_sync_gui.py` |
| 内容同步选择阶段检查 | [content_sync/preflight.py](../tools/content_sync/preflight.py)；占用提示、目录统计与只读抽检，由 GUI 后台调用 | `test_content_sync_preflight.py`、`test_content_sync_gui.py` |
| 历史 TB、项目配置与术语匹配 | [history_tb.py](../tools/history_tb.py)、[tb_projects.py](../tools/tb_projects.py)、[term_matching.py](../tools/term_matching.py) | 对应同名测试 |

测试路径未写前缀时均位于 `tests/`。

## 工作流与工具

| 功能 | 业务代码 | 统一 GUI | 规则 / 测试入口 |
|---|---|---|---|
| 一键质量检查 | `tools/workflow/workflow_runner.py` | `tools/workflow/qt_page.py` | [README](../tools/workflow/README.md)、`test_workflow_runner.py` |
| QA 详细设置与取消恢复 | `tools/workflow/gui_options.py` 保存选项 | `tools/workflow/qt_settings.py` | `test_toolshub_gui.py` |
| QA 汇总 / 修订回填 | `tools/workflow/review_sheet.py`、`revision_applier.py` | 同上 | `test_workflow_review_sheet.py` |
| PhraseLoom | `phraseloom/strings_workflow.py` | `phraseloom/qt_page.py` | [专项地图](../phraseloom/AGENTS.md)、`tests/phraseloom_tests/` |
| 内容同步 | `tools/content_sync/master_to_target.py`、`target_to_master.py` | 同包 `qt_page.py` | [README](../tools/content_sync/README.md)、`test_content_sync*.py`、`test_momotools_utilities.py` |
| 列操作 / 兼容性重存 | `tools/column_tools/processor.py`、`tools/excel_compatibility/processor.py` | `tools/excel_utilities_pages.py` | [列操作](../tools/column_tools/README.md)、[重存](../tools/excel_compatibility/README.md)、`test_momotools_utilities.py` |
| 同名文件替换 / 未翻译统计 | `tools/deep_replace/replacer.py`、`tools/untranslated_stats/stats.py` | 同上 | [替换](../tools/deep_replace/README.md)、[统计](../tools/untranslated_stats/README.md)、`test_momotools_utilities.py` |
| Batch 拆分与复原 | `tools/excel_batcher/excel_batcher.py` | 同包 `qt_page.py` | [README](../tools/excel_batcher/README.md)、`test_excel_batcher.py` |
| 合并表格 | `tools/excel_merger/merge_active_sheets.py` | 同包 `qt_page.py` | [README](../tools/excel_merger/README.md)、`test_excel_merger.py` |
| 法语 NBSP | `tools/french_nbsp_restorer/restore_french_nbsp.py` | 同包 `qt_page.py` | [README](../tools/french_nbsp_restorer/README.md)、`test_french_nbsp_restorer.py` |
| Xbench 转换 | `tools/xbench_report_transformer/transform_xbench_report.py` | 同包 `qt_page.py` | [README](../tools/xbench_report_transformer/README.md)、`test_xbench_report_transformer.py` |

三项一致性检查共用 [consistency_text.py](../tools/consistency_text.py) 生成比较文本和原文偏移，
规则见[文本归一化](consistency-normalization.md)，覆盖于 `test_consistency_text.py`。

单项检查包：`term_pair_checker`、`tag_placeholder_checker`、`line_break_checker`、
`source_consistency_checker`、`target_consistency_checker`、`substring_consistency_checker`、
`content_fidelity_checker`、
`chinese_target_checker`、`target_text_checker`。业务位于包内处理模块，测试位于 `tests/test_<工具>.py`；
规则从[文档索引](README.md#业务规则)进入。统一 QA 调用这些处理能力，不复制判定逻辑。

## 兼容边界

- 所有 GUI 仅使用 PySide6。`tools/*/*_gui.py`、`tools/workflow/workflow_gui.py` 和 `phraseloom/gui.py` 只保留启动转接，不再维护独立页面或 Tk 控件。
- `phraseloom gui` / `phraseloom-gui` 打开 PhraseLoom；Batch、合并、NBSP、Xbench 旧脚本打开对应 Qt 页面；术语、Tag、换行、同 Source 不同 Target、Target 中文旧脚本进入统一 QA 并仅预选对应检查，不自动运行。
- 单项检查旧 GUI 入口现在使用统一 QA 报告与修订流程；需要原单项报告格式时使用对应 CLI。业务 CLI 和工作簿协议保持兼容。
- 根目录 `extract_terms_from_excel.py`、`extract_terms_gui.py`、`qatools_cli.py` 保留兼容作用；不能因名称旧就删除。
- 内容同步、PhraseLoom 和 Batch 复原是三种不同的匹配/映射协议，不合并它们的写回逻辑。
- 新工具的原位输出约定不应用到既有 QA 报告工具；日志无附加 report 的约定也不适用于 QA 报告工作簿。

## 平台、文档与产物

- Windows：[构建脚本](../scripts/build_windows_release.ps1)、[安装定义](../packaging/QAtools.iss)、[使用说明](../packaging/README-Windows.txt)。
- macOS：[Finder 安装脚本](../scripts/install_macos_qa_workflow.py)、[操作手册](macos-finder-workflow.md)。
- `docs/qa-workflow-guide/`：静态说明网页；截图生成、离线资源和验证步骤见[维护入口](qa-workflow-guide/README.md)。
- `docs/archive/`：仅保存来源与已完成决策；不保留待执行旧方案作为当前需求。
- `build/`、`dist/`、`outputs/`、`testfiles/` 和生成的 Excel 均为本地产物，不纳入提交。

## 扩展一个工具

1. 先定义并测试业务函数的输入、输出、空值、冲突和保存失败行为；更新工具 README。
2. GUI 在工具包内增加 `qt_page.py`，继承 `AsyncPage`；GUI 线程采集参数，后台处理文件。
3. 在 `PAGE_FACTORIES` 注册稳定 key，在 `TOOL_GROUPS` 选择分组；后台任务必须被退出保护识别，包括嵌套 Tab。
4. CLI 用原生参数解析器调用同一业务函数，再登记 `COMMANDS` 和别名；更新[CLI 手册](cli-usage.md)，不要把仅 GUI 的功能写成已支持命令。
5. 更新本页入口与验证路径、文档索引和能力地图。运行相关测试，再按根 AGENTS 完成全量检查。

## 大文件评估

2026-09-10 维护盘点，行数是本次判断的快照，不作为硬性拆分阈值。

| 文件 | 行数 / 调整 | 判断 |
|---|---|---|
| 原 `tools/qt_pages.py` | 1,807 → 39 | 已按工具拆分；最大新页面 696 行，QA 设置独立 314 行；保持原类导入与方法行为 |
| `tools/excel_batcher/excel_batcher.py` | 1,341 | 后续可分为 OOXML 读写、manifest、拆分/复原编排、CLI；先保留外部函数和工作簿结构契约 |
| `tools/workflow/workflow_gui.py` | 1,158 → 20 | 已收敛为 Qt 启动转接；统一页面在同包 `qt_page.py` |
| `tools/term_pair_checker/extract_terms_from_excel.py` | 958 | 匹配、输出已有辅助模块；后续修改相关规则时再考虑抽 CLI，避免同时搬动复杂回扫逻辑 |
| `tools/tag_placeholder_checker/check_tags_and_placeholders.py` | 816 | 后续可分 token 提取/结构校验与 Excel 报告；必须保留 memoQ、引号属性、兄弟节点换序测试 |
| `phraseloom/strings_workflow.py` | 721 | 导出/回填编排已有包边界，暂不为行数改变 Strings 协议 |
| `tools/excel_merger/merge_active_sheets.py` | 687 | OOXML 合并需保持 shared strings、命名空间与行列上限，暂不拆 |
| 大型测试文件（约 800–1,200 行） | 场景多于职责混杂 | 后续可按场景拆，保留公共 fixture 和全部断言，不因文件长删除覆盖 |

上述后续建议不代表本轮已实施；业务重构需独立验证，不和新增 CLI 混在一次改动中。
