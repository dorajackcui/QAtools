# GUI 约定

统一 GUI 使用 PySide6，入口为 `qatools gui`。工具排列见[产品能力地图](../README.md#能力地图)，实现位置见[仓库地图](repository-map.md)。

## 页面与状态

- 页面在主窗口中持久保留，切换导航不会重新构造页面。
- 内容同步的两个 Tab 分别保存输入、结果和日志；后台任务继续运行。
- 表单不显示常驻解释段落；保留输入标签、占位提示、输出预览和运行状态。
- 底部操作栏固定；导航可滚动，“设置”固定在侧栏底部。
- 目录工具的实时日志行为见[运行日志](tool-logs.md)；输入/输出规则仍分别见工具 README。

## 表头别名

一键质量检查和法语 NBSP 页面从工作表首行自动识别 Source / Target 列。
在侧栏“设置”中为 Source 和 Target 分别填写别名，每行一个。

- 忽略大小写及首尾空白，精确匹配整个表头；`Source` 不匹配 `SourceFile`。
- 自定义别名优先于内置 `source` / `target`；未命中自定义别名才使用内置表头。
- 同级别匹配多列或没有匹配时清空对应列、显示“请指定”，由用户填写列字母。
- 设置持久化到当前用户；保存后重新识别已经加载的 QA / NBSP 工作簿。
- 这不是所有工具的列规则：内容同步与未翻译统计手动指定列；PhraseLoom 使用自己的表头/索引参数。

实现：[header_aliases.py](../tools/header_aliases.py)、[excel_metadata.py](../tools/excel_metadata.py)。
QA 的“记住选项”是独立配置，保存范围与恢复行为见 [workflow README](../tools/workflow/README.md#使用)。
