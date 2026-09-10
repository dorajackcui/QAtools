# 术语与 Tag 的职责边界

本文只维护两类工具之间的分工，不复制各工具的识别表和参数默认值。

| 对象 | 负责工具 | 权威规则 |
|---|---|---|
| `【术语】`、`[术语]` / `［术语］` 的术语对与历史 TB 回扫 | 术语检查 | [术语 README](../tools/term_pair_checker/README.md#规则) |
| 尖括号 Tag、color Tag、花括号 Placeholder、字面换行标记、memoQ Marker | Tag 检查 | [Tag README](../tools/tag_placeholder_checker/README.md#规则) |
| 多检查器组合、统一问题表与修订回填 | Workflow | [workflow README](../tools/workflow/README.md) |

- `<...>` 和 `{...}` 不作为术语 mark；`[color=...]`、`[/color]` 不进入术语表。
- 两类检查不共享术语映射，也不修改被检查的 Source / Target 正文。
- 统一 GUI 的“常规 Tag”和“memoQ Marker”互斥；CLI 可显式组合 token 类型，纯数字 `{n}` 在组合模式下归 memoQ，避免重复报告。
- 统一 QA 的最终报告不是独立检查器报告表的简单拼接，合并和删除临时表规则只在 workflow README 维护。

调用参数与兼容别名见 [CLI 手册](cli-usage.md)，不要从旧截图推断默认值。
