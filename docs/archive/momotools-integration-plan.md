# momoTools 合并记录

2026-09-10 完成第一期合并。早期实施方案、旧导航和另存目录建议已移除，避免与当前规则冲突；历史版本可通过 Git 查看。

## 来源与范围

- 来源仓库：`dorajackcui/momoTools`，调研基线 `bb0a7a846edd8884fec637389fc2f609ac31b77a`。
- QAtools 合并前基线：`dbe6e30642439a706dc8b14a8da994aa2802fc09`。
- 集成提交：`1b090f6de48bc62314215f6b13b1cf754fc063c0`。
- 六个工具：Master → 小表、小表 → Master、Column Clear、Compatibility、Deep Replace、Untranslated Stats。

## 保留的决策

- 双向同步使用 openpyxl，不引入 pandas 的缺失值转换；空白和正文语义以[内容同步 README](../../tools/content_sync/README.md)为准。
- 列操作和兼容性重存依赖独立的 Excel COM 会话；原位、失败保护和宏行为以[重存 README](../../tools/excel_compatibility/README.md)为准。
- GUI 复用 QAtools 风格；第一期先迁入 GUI，后续 CLI 的当前入口见[调用手册](../cli-usage.md#目录批处理)。
- 不迁入 momoTools 的 Term Extractor、Master Update、Batch 配置执行器，不改变 PhraseLoom 和现有 Batch 的协议。

代码落点见[仓库地图](../repository-map.md)，当前业务规则均从[文档索引](../README.md)进入。
