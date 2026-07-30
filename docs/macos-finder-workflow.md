# macOS Finder 一键质量检查

QAtools 可以安装 Finder 快速操作，把 `.xlsx` 或 `.xlsm` 文件直接发送到
统一 GUI 的“一键质量检查”页面。

## 安装

在仓库根目录运行：

```bash
python scripts/install_macos_qa_workflow.py
```

安装后，在 Finder 中右键 Excel 文件，选择：

```text
快速操作 → QA workflow
```

## 行为

- Toolshub 已打开：复用当前窗口，切换到“一键质量检查”并载入文件。
- Toolshub 未打开：启动 Toolshub，然后载入文件。
- 页面会尝试自动识别工作表以及 Source / Target 列。

如果 Finder 没有显示入口，请前往：

```text
系统设置 → 隐私与安全性 → 扩展 → Finder
```

确认“QA workflow”已启用。
