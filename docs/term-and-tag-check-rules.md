# 术语检查与 Tag 检查规则汇总

本文档汇总当前 runtime 中“术语对检查”和“Tag 检查”的识别类型、默认值和过滤规则。

## 术语对检查

术语对检查用于从 source / target 中提取被 mark 包裹的术语，并按出现顺序建立 source -> target 术语对。后续回扫整表时，会忽略支持的 mark 外壳，用去 mark 后的术语文本做命中和校验。

### 支持的术语 mark

| mark 类型 | 可识别形式 |
| --- | --- |
| `【】` | `【术语】` |
| `[]` | `[术语]`、`［术语］` |

`<...>` 和 `{...}` 不作为术语 mark；它们分别交给普通 tag / placeholder 检查和 memoQ tag 检查。

### 默认选择

| 入口 | 默认 mark |
| --- | --- |
| CLI / `process_excel()` 未显式传参 | `【】`、`[]` |
| 术语对检查 GUI | `【】`、`[]` |
| Workflow GUI | `【】`、`[]` |

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

## Tag 检查

Tag 检查用于逐行比较 source / target 中 tag、placeholder 和 mark token 是否一致。它只检查 token 计数和文本是否一致，不建立术语对。

### 支持的检查类型

| 内部类型 | UI 名称 | 可识别形式 |
| --- | --- | --- |
| `angle` | `<...> tag` | `<...>` |
| `square_color` | `[color=...] tag` | `[color=...]`、`[/color]` |
| `brace` | `{...} placeholder` | `{name}`、`{count}` 等普通花括号 placeholder |
| `newline` | `\n mark` | 字面量 `\n` |
| `memoq` | memoQ tag | `{1}`、`{2>`、`<3}` |

### 默认选择

| 入口 | 默认检查类型 |
| --- | --- |
| CLI / `process_excel()` 未显式传参 | `angle`、`square_color`、`brace`、`newline`、`memoq` |
| Tag 检查 GUI | `angle`、`square_color`、`brace`、`newline` |
| Workflow GUI | `angle`、`square_color`、`brace`、`newline` |

GUI 中 memoQ tag 与其他检查类型互斥：勾选 memoQ 会取消普通 tag 组，勾选普通 tag 组会取消 memoQ。CLI / `process_excel()` 保持可显式组合，方便脚本按需调用。

### Tag token 过滤

- `angle` 默认检查所有 `<...>`。
- `square_color` 只检查 `[color=...]` 和 `[/color]`，不会把普通 `[stage1]` 或术语 mark `[]` 都当成 tag。
- `{1}`、`{2>...<3}` 这类 memoQ protected marker 不会作为普通 `{...}` placeholder 检查，会交给 `memoq` 类型处理。
- 普通文本里的 `<apple>` 也会作为 `<...>` tag 检查。

### 比对方式

- 每一行分别提取 source 和 target 的 token。
- 按检查类型分别计数，要求 source / target 的 token 文本和数量完全一致。
- 不一致时写入 `标签占位问题` sheet，并说明 target 缺少或多出了哪些 token。
- 统计信息写入 `检查汇总` sheet。

## Workflow 中的关系

Workflow 同时运行术语对检查和 Tag 检查时，顺序是：

1. 先运行术语对检查，写入 `术语表` 和 `问题列`。
2. 再运行 Tag 检查，写入 `标签占位问题` 和 `检查汇总`。

两类检查不会共享术语结果，也不会互相修改 source / target 原文。术语检查不再使用 `<...>` 作为术语 mark；普通 tag / placeholder 与 memoQ tag 在 Tag 检查中作为独立检查类型选择。
