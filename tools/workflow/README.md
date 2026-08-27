# 一键质量检查与修订回填

统一 CLI：`qatools qa --help`。GUI 仍用于交互式检查和修订回填。

workflow 会按顺序执行质量检查板块中选中的项目，并将结果写入同一个 Excel。

## CLI 快速开始

默认运行八项常用质量检查：术语、同 Source 不同 Target、Tag、换行数量、
数字、URL、Target 中文和 Target 文本规范。“同 Target 不同 Source”默认关闭，
可在 GUI 中勾选或通过 `--check target-consistency` 运行。

```bash
qatools qa input.xlsx -s Sheet1 -c A -t B
```

只运行指定检查：

```bash
qatools qa input.xlsx -c A -t B \
  --check tag \
  --check number \
  --check url
```

## GUI 检查区块

GUI 按风险类型划分为三个区块：

- `术语与翻译一致性`：术语检查，以及并排显示的“同 Source 不同 Target”和
  “同 Target 不同 Source”。
- `内容保真检查`：Tag / Placeholder、换行数量、数字一致性和 URL 一致性。
- `Target 文本质量`：Target 中文和 Target 文本规范。

统一 GUI 在术语检查、Tag / Placeholder 和 Target 文本规范右侧提供倒三角入口；
点击后在模态窗口中调整对应设置，不再占用主页面的纵向空间。点击“确定”应用本次修改，
点击“取消”、关闭按钮或按 `Esc` 会恢复打开窗口前的表单值。TB 项目的“保存当前”
和“删除”属于明确的项目管理操作，执行后立即生效。

“同 Target 不同 Source”可能命中合理的译文复用，因此默认关闭；其余常用检查
默认开启。点击“全选”仍会选中包括该辅助检查在内的全部项目。

Target 文本规范检查默认启用全部四条规则，也可以单独选择：

```bash
qatools qa input.xlsx -c A -t B \
  --check text \
  --text-rule abnormal-punctuation \
  --text-rule consecutive-spaces \
  --text-rule leading-trailing-spaces
```

可选文本规则为：

- `abnormal-punctuation`：检查异常省略号及重复逗号、句号、冒号、分号等；`...`、`…`、连续感叹号和问号放行。
- `consecutive-spaces`：Target 任意位置出现 2 个及以上普通空格，包括开头和结尾。
- `leading-trailing-spaces`：Target 开头或结尾出现一个或多个普通空格。
- `mixed-width`：同类字符或标点在一个 Target 中混用全角和半角形式。

完整参数见 [`../../docs/cli-usage.md`](../../docs/cli-usage.md)。

## 快速修订流程

1. 在 GUI 中点击“开始检查”。
2. 打开生成的 `workflow_check_<原文件名>`。
3. 在 `问题处理` 工作表筛选、查看问题，并在黄色的 `修改后target` 列填写最终译文。
4. 不需要修改的行保持空白。
5. 保存检查报告，回到 GUI 点击“应用修订”。
6. 选择检查报告和输出位置，生成 `revised_<原文件名>`。

## 问题处理工作表

列顺序为：

- `行号`
- `source`
- `target`
- `修改后target`
- `问题描述`
- `检查项`

同一原始行命中多个检查时只出现一行，检查项和问题描述会合并。各检查问题表中的专项字段也会合并进问题描述，包括术语来源、Tag 问题类型、换行数量、同组行号和命中字符。点击行号会跳转到原数据工作表对应的 target 单元格。

workflow 报告不保留各类检查自己的问题工作表。最终可见工作表为原业务工作表、可选的 `术语表`、`问题处理` 和 `质量检查汇总`。

`修改后target`填写内容表示需要修改，留空表示忽略。该工作表不再需要额外的处理状态。

## 安全规则

- 修订稿默认另存，不覆盖检查报告。
- 回填前会比较 `问题处理` 中保存的原 target 与当前数据表；不一致的行会跳过并在完成提示中列出。
- workflow 报告生成时即会移除 `术语QA问题` 辅助列，并恢复其右侧原有列的位置。
- 修订稿保留原工作簿中的业务工作表，只移除本次 workflow 生成的 `术语表`、`问题处理` 和 `质量检查汇总`。
