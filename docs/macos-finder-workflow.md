# macOS Finder 快速操作

QAtools 可以安装两个 Finder 快速操作，支持 `.xlsx` 和 `.xlsm` 文件：

- `ABC · QA workflow`：把文件发送到统一 GUI 的“一键质量检查”页面。
- `ABC · NBSP restore`：自动识别工作表和 Target 列，直接执行法语 NBSP 恢复。

## 安装

在仓库根目录运行：

```bash
python scripts/install_macos_qa_workflow.py
```

安装 `NBSP restore`：

```bash
python scripts/install_macos_qa_workflow.py --action nbsp
```

同时安装或更新两项：

```bash
python scripts/install_macos_qa_workflow.py --action all
```

安装后，在 Finder 中右键 Excel 文件，选择所需操作：

```text
快速操作 → ABC · QA workflow
快速操作 → ABC · NBSP restore
```

## 行为

- Toolshub 已打开：复用当前窗口，切换到“一键质量检查”并载入文件。
- Toolshub 未打开：启动 Toolshub，然后载入文件。
- 页面会尝试自动识别工作表以及 Source / Target 列。
- `ABC · NBSP restore` 会切换到“法语 NBSP 恢复”，使用自动识别的 Target 列和
  默认开始行 2 直接处理。
- NBSP 恢复不会覆盖输入文件；结果写到同一目录下的
  `french_nbsp_restore_<原文件名>`。

如果 Finder 没有显示入口，请前往：

```text
系统设置 → 隐私与安全性 → 扩展 → Finder
```

确认名称以“ABC ·”开头的快捷操作已启用。
