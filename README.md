![QAtools：Excel 本地化工具箱，橙色句点字标与表格网格](docs/assets/readme-hero.svg)

# Excel 本地化，从待译到交付

整理待译文本、检查译文、回填修订。工作簿全程在本机处理。

**[开始使用](#开始使用)** · [浏览工具](#能力地图) · [使用文档](docs/README.md)

## 能力地图

### 一份报告，处理整批译文问题

[一键质量检查](tools/workflow/README.md)将漏译、术语、译文一致性、Tag、数字和格式等问题汇总到一份 Excel 报告。同一原文行的问题集中显示；填好修订后，可以在应用里一键回填并另存工作簿。

### 把待译文本整理成一张清单

[PhraseLoom](phraseloom/README.md)将待译内容去重后导出为 Strings 工作簿。译者在清单里填写译文，完成后回填到原有表格结构。

### 在 Master 和小表之间同步

[内容同步](tools/content_sync/README.md)支持把 Master 的内容分发到小表，也支持把小表更新汇回 Master。

更多常用工具：

- **批量处理表格**：[Batch 拆分与复原](tools/excel_batcher/README.md) · [合并表格](tools/excel_merger/README.md) · [列操作](tools/column_tools/README.md) · [兼容性重存](tools/excel_compatibility/README.md)
- **整理项目文件**：[文件分组](tools/file_grouping/README.md) · [文件提取](tools/file_collector/README.md) · [同名文件替换](tools/deep_replace/README.md)
- **翻译辅助**：[法语 NBSP 恢复](tools/french_nbsp_restorer/README.md) · [未翻译统计](tools/untranslated_stats/README.md) · [Xbench QA 转换](tools/xbench_report_transformer/README.md)

## 开始使用

- **Windows**：运行安装包，从开始菜单打开 QAtools。也可通过资源管理器右键菜单带入文件或目录。见 [Windows 安装说明](packaging/README-Windows.txt)。
- **macOS**：打开与芯片架构匹配的 DMG，将 QAtools 拖到「应用程序」。见 [macOS 安装说明与功能限制](packaging/README-macOS.txt)。

安装包无需 Python。如果从源码运行，在仓库根目录使用 Python 3.11 或更高版本执行：

```bash
python -m pip install -e .
qatools gui
```

需要批量运行或接入自动化？见 [CLI 使用指南](docs/cli-usage.md)。
