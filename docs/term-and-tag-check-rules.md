# 术语检查与 Tag 检查规则汇总

本文档汇总当前 runtime 中“术语检查”和“Tag 检查”的识别类型、默认值和过滤规则。

## 术语检查

术语检查支持两种模式：选择 mark 时，从 source / target 中提取术语并建立 source -> target 术语对；不选择 mark 时，必须提供历史 TB，并仅按历史术语回扫整表。回扫时会忽略支持的 mark 外壳，用去 mark 后的术语文本做命中和校验。

### 支持的术语 mark

| mark 类型 | 可识别形式 |
| --- | --- |
| `【】` | `【术语】` |
| `[]` | `[术语]`、`［术语］` |

`<...>` 和 `{...}` 不作为术语 mark；它们交给常规 Tag 或 memoQ Marker 检查。

### 默认选择

| 入口 | 默认 mark |
| --- | --- |
| CLI / `process_excel()` 未显式传参 | `【】`、`[]` |
| 术语检查 GUI | `【】`、`[]` |
| Workflow GUI | `【】`、`[]` |

GUI 可以取消全部 mark，但此时必须提供历史 TB。CLI 使用 `--no-term-mark --history-tb <文件>` 进入同一模式。

### 术语候选过滤

术语提取会先找出被所选 mark 包裹的候选，再过滤明显不是术语的片段。

内置过滤规则：

- 单个 ASCII 字母会被过滤，例如 `[b]`、`【Z】`。
- 不包含任何文字或字母的候选会被过滤，例如 `[123]`、`【+10%】`。
- 方括号格式 tag 会被过滤，例如 `[color=red]`、`[/color]`。

会保留的例子：

- `[火]`
- `[HP]`
- `【苹果】`

### 术语配对和回扫

- 同一行 source / target 都提取到多个术语时，按出现顺序一一配对。
- 术语表以去 mark 后的 source 术语作为唯一键。
- 术语检查回扫时会去掉支持的 mark 外壳，按纯文本匹配。
- 默认忽略大小写。
- 默认使用混合边界匹配：中文按包含匹配，英文和数字按边界匹配，避免 `ACC` 命中 `account`。
- 如果传入历史 TB，会先加载历史 TB，再合并本批次新增术语；命中历史 source 时优先使用历史 target。
- 两侧 mark 数量不一致时，以 source 术语为校验义务：只要所有 source 术语都命中已知预期译法，target 侧额外的 mark 不会单独产生数量问题，因为它可能对应 source 中未加 mark 的表达。

## Tag 检查

Tag 检查用于逐行比较 source / target 中 tag、placeholder 和 mark token 是否一致。它检查 token 计数和文本，并校验成对尖括号 tag 的父子结构，不建立术语对。

### 支持的检查类型

| 内部类型 | UI 名称 | 可识别形式 |
| --- | --- | --- |
| `angle` | `<...> tag` | `<...>`，包括空 tag `<>` 和仅含空白的 tag `< >` |
| `square_color` | `[color=...] tag` | `[color=...]`、`[/color]` |
| `brace` | `{...} placeholder` | `{name}`、`{{name}}`、`{count}`、`{1}`、空 placeholder `{}` 等花括号 placeholder |
| `newline` | `\n mark` | 字面量 `\n` |
| `memoq` | memoQ marker | `{1}`、`{2>`、`<3}` |

### 默认选择

| 入口 | 默认检查类型 |
| --- | --- |
| CLI / `process_excel()` 未显式传参 | 常规 Tag：`angle`、`square_color`、`brace`、`newline` |
| Tag 检查 GUI | 常规 Tag：`angle`、`square_color`、`brace`、`newline` |
| Workflow GUI | 常规 Tag：`angle`、`square_color`、`brace`、`newline` |

GUI 提供互斥的“常规 Tag”和“memoQ Marker”两种模式。常规模式检查所选的四类常规 token；memoQ 模式只检查 memoQ marker。Tag 独立 GUI 和 Workflow GUI 都可选择尖括号过滤 JSON；CLI / `process_excel()` 仍可显式组合类型。

### Tag token 过滤

- `angle` 默认检查 `<...>`，包括空 tag `<>` 和仅含空白的 tag `< >`，能跳过引号属性中的 `>`；尖括号内含非空白内容且首尾带空白的普通比较表达式不作为 tag。
- `square_color` 只检查 `[color=...]` 和 `[/color]`，不会把普通 `[stage1]` 或术语 mark `[]` 都当成 tag。
- 常规模式会把 `{1}` 和完整的 `{2>...<3}` 包络分别当作一个花括号 placeholder，但不会单独捕获 `{2>`、`<3}` 两侧半 marker。
- memoQ 模式会提取 `{1}`、`{2>`、`<3}`。显式同时选择 `brace` 和 `memoq` 时，纯数字 `{1}` 只归入 `memoq`，避免重复报告。
- 普通文本里的 `<apple>` 也会作为 `<...>` tag 检查。

### 比对方式

- 每一行分别提取 source 和 target 的 token。
- 按检查类型分别计数，要求 source / target 的 token 文本和数量完全一致。
- 成对尖括号 tag 还会比较闭合和父子层级；允许彼此独立的兄弟 tag 随翻译语序换序。
- 不一致时写入 `标签占位问题` sheet，并说明 target 缺少或多出了哪些 token。
- 统计信息写入 `检查汇总` sheet。
- `标签占位问题` 和 `检查汇总` 是保留结果表名，结果文件中的同名工作表会被重建。

## Workflow 中的关系

Workflow 同时运行术语检查和 Tag 检查时，顺序是：

1. 先运行术语检查，写入 `术语表` 和 `问题列`。
2. 再运行 Tag 检查，写入 `标签占位问题` 和 `检查汇总`。

两类检查不会共享术语结果，也不会互相修改 source / target 原文。术语检查不再使用 `<...>` 作为术语 mark；常规 Tag 与 memoQ Marker 在 GUI 中作为互斥模式选择。
