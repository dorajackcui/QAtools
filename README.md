# QAtools

QAtools 是面向本地化团队的本地 Excel 工具箱，提供统一的 Windows GUI、CLI
和可组合的质量检查工作流。它不会上传工作簿；检查结果默认另存为新的 Excel。

## 最快开始

Windows 便携版无需安装 Python：解压发布包后双击 `QAtools.exe`。

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

### 自定义表头识别

GUI 默认从工作表首行自动识别 `source` 和 `target` 列。需要兼容其他常用表头时，
点击侧栏底部的“⚙ 设置”，分别为 Source 和 Target 添加别名；每行填写一个，匹配时
忽略大小写和首尾空格。内置的 `source` / `target` 始终保留，设置会持久化到当前用户。

## 一键质量检查

`qatools qa` 将选中的检查写入同一份报告。组合运行时，所有检查共用一次主工作簿
读取和一次保存，适合处理大型 Excel。

| 分组 | 检查 | 默认 | 用途 |
|---|---|---:|---|
| 术语与翻译一致性 | 术语检查 | 开 | 新术语 mark 与历史 TB |
| 术语与翻译一致性 | 同 Source 不同 Target | 开 | 定位同源文本的译法分歧 |
| 术语与翻译一致性 | 同 Target 不同 Source | 关 | 辅助发现可能的译文误复用 |
| 内容保真检查 | Tag / Placeholder | 开 | 比较 Tag、占位符和 memoQ Marker |
| 内容保真检查 | 换行数量 | 开 | 比较真实换行数量 |
| 内容保真检查 | 数字一致性 | 开 | 比较排除 URL、Tag 后的数字表达式 |
| 内容保真检查 | URL 一致性 | 开 | 比较 URL 内容和重复次数 |
| Target 文本质量 | Target 中文 | 开 | 定位中文字符和中文标点残留 |
| Target 文本质量 | Target 文本规范 | 开 | 检查异常标点、空格和全半角混用 |

报告中的 `问题处理` 会把同一原始行的多个问题合并。填写 `修改后target` 后，
可以在 GUI 中应用修订并生成新的工作簿。界面操作见
[GUI 使用指南](docs/qa-workflow-guide/index.html)，完整选择参数见
[CLI 使用指南](docs/cli-usage.md#一键质量检查)，报告与回填规则见
[workflow README](tools/workflow/README.md)。

## 其他工作流

| 工作流 | CLI | 说明 |
|---|---|---|
| PhraseLoom Strings | `qatools phraseloom` | 导出待翻译 Strings，完成后回填原工作簿 |
| 法语 NBSP 恢复 | `qatools french-nbsp` | 恢复法语标点所需的不换行空格 |
| Batch 拆分与复原 | `qatools batch` | 分批处理大型工作表并按原行位复原 |
| 合并表格 | `qatools merge-sheets` | 合并目录内工作簿的活动工作表 |
| Xbench QA 转换 | `qatools xbench` | 把 Xbench 报告转换为行级问题表 |

单项检查仍保留兼容 CLI；完整命令目录和示例统一维护在
[CLI 使用指南](docs/cli-usage.md)。

## 文档路由

| 需要 | 入口 |
|---|---|
| 运行命令或编写自动化 | [CLI 使用指南](docs/cli-usage.md) |
| 确认检查、输出或回填规则 | [文档索引](docs/README.md) → 对应工具 README |
| 使用 PhraseLoom | [PhraseLoom README](phraseloom/README.md) |
| 修改代码 | [AGENTS.md](AGENTS.md) |

文档按 AI-first 原则维护：根 README 负责入口，CLI 手册负责调用，工具 README
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
```
