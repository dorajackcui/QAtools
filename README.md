# QAtools

QAtools 是面向本地化工作的 Excel 工具箱，统一提供质量检查、文本修复、
Xbench 报告转换和 PhraseLoom Strings 工作流。

所有功能共用一个仓库、一个安装入口、一个 GUI 和一个 CLI；原有脚本路径继续
保留，避免已有自动化失效。

## 快速开始

要求 Python 3.11 或更高版本。

```bash
python -m pip install -e .
```

打开统一 GUI：

```bash
qatools gui
```

查看统一 CLI：

```bash
qatools --help
qatools list
```

未安装命令行脚本时，可等价使用：

```bash
python -m qatools --help
python -m qatools gui
```

## 功能目录

| 功能 | 统一 CLI | GUI | 详细说明 |
|---|---|---:|---|
| 一键质量检查 | `qatools qa` | 是 | [CLI 指南](docs/cli-usage.md#一键质量检查) |
| PhraseLoom Strings | `qatools phraseloom` | 是 | [PhraseLoom](phraseloom/README.md) |
| 术语检查 | `qatools term-check` | 是 | [术语检查](tools/term_pair_checker/README.md) |
| Tag / Placeholder 检查 | `qatools tag-check` | 是 | [Tag 检查](tools/tag_placeholder_checker/README.md) |
| 换行数量检查 | `qatools line-break-check` | 是 | [换行检查](tools/line_break_checker/README.md) |
| 同源译文一致性 | `qatools consistency-check` | 是 | [一致性检查](tools/source_consistency_checker/README.md) |
| Target 中文检查 | `qatools chinese-check` | 是 | [中文检查](tools/chinese_target_checker/README.md) |
| 法语 NBSP 恢复 | `qatools french-nbsp` | 是 | [法语 NBSP](tools/french_nbsp_restorer/README.md) |
| Xbench QA 转换 | `qatools xbench` | 是 | [Xbench 转换](tools/xbench_report_transformer/README.md) |

完整参数、批处理建议和旧入口映射统一维护在
[CLI 使用指南](docs/cli-usage.md)。

## 常用流程

一次执行全部质量检查：

```bash
qatools qa input.xlsx -s Sheet1 -c A -t B -o workflow_check_input.xlsx
```

只执行部分检查：

```bash
qatools qa input.xlsx -c A -t B \
  --check tag \
  --check line-break \
  --check consistency
```

PhraseLoom 导出与回填：

```bash
qatools phraseloom export source.xlsx
qatools phraseloom restore source_strings.xlsx
```

## Windows 便携版

在 64 位 Windows 上生成无需安装 Python 的开箱即用版本：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\build_windows_release.ps1
```

脚本会先运行回归测试，再生成 `QAtools.exe`（统一 GUI）、
`QAtools-CLI.exe`（统一 CLI）和完整 ZIP，输出到 `dist/`。如果已经单独完成测试，
可以追加 `-SkipTests`。分发包内的具体用法见 `packaging/README-Windows.txt`。

## 仓库结构

```text
qatools/                 统一 CLI 和命令注册表
phraseloom/              Strings 导出、清洗与回填
tools/                   各 QA 工具及统一 workflow
tests/                   全仓库回归测试
docs/                    CLI、规则、设计和迁移文档
toolshub_gui.py          统一 GUI 入口
pyproject.toml           安装、依赖和命令行入口
```

新增工具时，应保持三层边界：

1. 业务逻辑留在独立包中。
2. GUI 页面注册到 `toolshub_gui.py`。
3. CLI 命令注册到 `qatools/cli.py`，并在 CLI 指南中补一个示例。

测试会检查命令注册表中的每个正式命令是否已出现在 CLI 指南中。

## 文档

文档总入口：[docs/README.md](docs/README.md)

- 自动化和完整命令参数：[docs/cli-usage.md](docs/cli-usage.md)
- 术语与 Tag 规则：[docs/term-and-tag-check-rules.md](docs/term-and-tag-check-rules.md)
- PhraseLoom 合并记录：[docs/phraseloom-migration.md](docs/phraseloom-migration.md)
- macOS Finder 快速操作：[docs/macos-finder-workflow.md](docs/macos-finder-workflow.md)

## 开发验证

```bash
python -m unittest discover -s tests -v
python -m compileall -q qatools phraseloom tools tests
git diff --check
```

Finder 文件转发测试依赖 macOS 的 `fcntl`，在 Windows 上会自动跳过。
