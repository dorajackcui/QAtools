# 法语 NBSP 恢复

统一 CLI：`qatools french-nbsp --help`。

这个工具用于恢复 Excel target 列中的法语 non-breaking space（NBSP），输出新的 Excel 文件，不覆盖原文件。

## 规则

- `;`、`:`、`?`、`!`、`%` 前恢复 NBSP。
- `«` 后和 `»` 前恢复 NBSP。
- 已有普通空格、NBSP 或 narrow NBSP 会统一为 NBSP。
- URL 内标点和 `12:30` 这类时间冒号不会被改写。

## CLI

调用示例、参数和兼容脚本集中在 [CLI 手册](../../docs/cli-usage.md)。

如果指定结果列，不需要修复的 target 也会复制到结果列。

## GUI

以下旧脚本直接打开工具箱中的 PySide6 页面：

```bash
python3 tools/french_nbsp_restorer/restore_french_nbsp_gui.py
```

也可以通过统一入口打开：

```bash
qatools gui
```

统一 GUI 会自动识别工作表首行的 Target 列；匹配优先级、冲突处理及自定义配置见
[GUI 表头别名设置](../../README.md#自定义表头识别)。

## 实现位置

统一 PySide6 页面在同目录 `qt_page.py`；页面注册、共享控件和对应测试见[仓库地图](../../docs/repository-map.md)。
