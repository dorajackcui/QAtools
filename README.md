# QAtools

QAtools 是面向本地化团队的本地 Excel 工具箱，提供统一的 Windows GUI、CLI
和可组合的质量检查工作流。它不会上传工作簿；检查结果默认另存为新的 Excel。

## 最快开始

Windows 安装版无需安装 Python：运行安装包后，从开始菜单打开 QAtools。

源码运行要求 Python 3.11 或更高版本：

```bash
python -m pip install -e .
qatools gui
```

Agent、脚本和批处理任务优先使用 CLI：

```bash
qatools qa input.xlsx -s Sheet1 -c A -t B
qatools --help
qatools help qa
```

GUI 表单不显示常驻说明段落，保留输入标签、占位提示、输出预览和运行状态；各工具的详细规则见下方工具文档。

### 自定义表头识别

一键质量检查和法语 NBSP 支持自定义表头别名；优先级、歧义处理和保存方式见
[GUI 约定](docs/gui-conventions.md#表头别名)。

## 一键质量检查

`qatools qa` 将选中的检查写入同一份报告。组合运行时，所有检查共用一次主工作簿
读取和一次保存，适合处理大型 Excel。

可选检查、默认状态及规则入口见 [CLI 手册中的检查目录](docs/cli-usage.md#一键质量检查)。

报告中的 `问题处理` 会把同一原始行的多个问题合并。填写 `修改后target` 后，
可以在 GUI 中应用修订并生成新的工作簿。界面操作见
[GUI 使用指南](docs/qa-workflow-guide/index.html)，完整选择参数见
[CLI 使用指南](docs/cli-usage.md#一键质量检查)，报告与回填规则见
[workflow README](tools/workflow/README.md)。

## 能力地图

下表按 GUI 导航顺序排列；CLI 列只列出已支持的入口。

| GUI 分组 | 工具与规则 | CLI | 用途 |
|---|---|---|---|
| 常用流程 | [一键质量检查](tools/workflow/README.md) | `qatools qa` | 多项检查生成统一报告，GUI 可应用修订 |
| 常用流程 | [PhraseLoom](phraseloom/README.md) | `qatools phraseloom` | 导出 Strings，完成翻译后回填并另存工作簿 |
| 常用流程 | [内容同步](tools/content_sync/README.md) | `qatools content-sync` | 两个 Tab：Master → 小表、小表 → Master |
| 文件与表格 | [兼容性重存](tools/excel_compatibility/README.md) | `qatools compatibility` | 通过 Microsoft Excel 重新保存工作簿 |
| 文件与表格 | [列操作](tools/column_tools/README.md) | `qatools columns` | 批量清空、插入、删除列 |
| 文件与表格 | [Batch 拆分](tools/excel_batcher/README.md) | `qatools batch` | 分批处理大型工作表并按原行位复原 |
| 文件与表格 | [合并表格](tools/excel_merger/README.md) | `qatools merge-sheets` | 合并目录内工作簿的活动工作表 |
| 文件与表格 | [同名文件替换](tools/deep_replace/README.md) | `qatools deep-replace` | 递归替换整个同名文件 |
| 翻译辅助 | [法语 NBSP 恢复](tools/french_nbsp_restorer/README.md) | `qatools french-nbsp` | 恢复法语标点所需的不换行空格 |
| 翻译辅助 | [未翻译统计](tools/untranslated_stats/README.md) | `qatools untranslated-stats` | 中文字符/英文词数及行数汇总 |
| 翻译辅助 | [Xbench QA 转换](tools/xbench_report_transformer/README.md) | `qatools xbench` | 把 Xbench 报告转换为行级问题表 |

单项检查仍保留兼容 CLI；完整命令目录和示例统一维护在
[CLI 使用指南](docs/cli-usage.md)。

内容同步双向及四项目录工具共用 GUI/CLI 业务层；调用与退出码见[目录批处理手册](docs/cli-usage.md#目录批处理)。列操作、兼容性重存需要桌面 Excel；
源码环境安装 COM 支持请使用 `python -m pip install -e ".[excel-com]"`。

## 文档路由

| 需要 | 入口 |
|---|---|
| 运行命令或编写自动化 | [CLI 使用指南](docs/cli-usage.md) |
| 确认检查、输出或回填规则 | [文档索引](docs/README.md) → 对应工具 README |
| 使用 PhraseLoom | [PhraseLoom README](phraseloom/README.md) |
| 修改代码 | [AGENTS.md](AGENTS.md) |

[仓库地图](docs/repository-map.md)说明实现位置和扩展方式；[待办](docs/backlog.md)记录未完成工作。

文档职责：根 README 负责入口，CLI 手册负责调用，工具 README
负责业务规则；`docs/archive/` 仅用于追溯，不代表当前行为。

## Windows 安装包构建

```powershell
powershell -ExecutionPolicy Bypass -File scripts\build_windows_release.ps1
```

脚本会运行回归测试，并在 `dist/` 生成单个 Windows 安装包。安装程序按当前用户
安装；后续版本直接运行新安装包即可覆盖升级。安装说明见
[README-Windows.txt](packaging/README-Windows.txt)。

## 开发验证

修改前先阅读 [AGENTS.md](AGENTS.md)。完整验证命令：

```bash
python -m unittest discover -s tests -v
python -m compileall -q qatools phraseloom tools tests
git diff --check
python scripts/check_docs.py
```
