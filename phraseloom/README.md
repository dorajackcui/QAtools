# PhraseLoom

统一 CLI：`qatools phraseloom --help`。原有 `phraseloom` 命令继续兼容。

PhraseLoom 把原始 Excel 整理成干净、去重、按相似结构分组的 Strings
工作簿，并在翻译完成后准确写回原文件。

它不做 TM 预填、术语抽取或自动生成译文。

## 安装

```powershell
python -m venv .venv
.venv\Scripts\python -m pip install -e .
```

运行测试：

```powershell
python -m unittest discover -s tests/phraseloom_tests -v
```

只运行完整 Strings 流程回归：

```powershell
python -m unittest discover -s tests/phraseloom_tests `
  -p test_strings_workflow_e2e.py -v
```

该用例会临时生成测试工作簿，自动完成“清洗 -> 自动完成 -> 填写测试译文
-> 回填”，并逐行验证重复文本、数字模板、关卡号、版本序列、颜色值、
Angle Tag、BBCode、花括号 Placeholder、已有 Target、自动完成、
首条 Sample/Context，以及相似句分组选项和最终工作簿结构。

## 日常流程

```text
原始 Excel
  -> 导出 Strings
  -> 译者填写 Target
  -> 回填译文
  -> 保留原格式的翻译结果
```

桌面界面：

```powershell
qatools phraseloom gui
```

交互式终端：

```powershell
qatools phraseloom
```

桌面 GUI 只有“导出 Strings”主页面。翻译完成后，点击页面右下角的
“回填译文…”并选择一个 Strings 工作簿，程序会直接开始回填，不再打开
单独的回填页面。交互式终端仍按“导出 / 回填”两个动作运行。

## 1. 导出 Strings

```powershell
qatools phraseloom export source.xlsx
```

默认生成：

```text
source_strings.xlsx
```

导出规则：

- Source 为空：忽略。
- Target 已有内容：视为已完成，不进入翻译清单且不修改。
- 纯数字、纯符号或 Tag-only：自动将 Source 作为 Target，不进入翻译清单。
- 完全相同的待翻译 Source：合并为一条。
- 沿用原有 Translation Unit 清洗：数字等可变部分压缩为模板，例如
  `通用补偿器LV1` 到 `LV4` 只导出 `通用补偿器LV{num1}`。
- 默认按各 Unit 首次出现的原始行排列，不启用相似句分组。
- 用户主动开启后，才对清洗结果进行相似结构分组。
- 只有同一模板的变体会合并；普通相似句不会合并，也不会自动生成 Target。

默认识别 `source`、`target` 和可选的 `context` 列。列名不同
时可以指定：

```powershell
qatools phraseloom export source.xlsx `
  --source-col en `
  --target-col fr `
  --context-col screen
```

需要相似句分组时显式启用：

```powershell
qatools phraseloom export source.xlsx --group-similar
```

输出工作簿包含两个可见部分：

1. `strings`：真正需要用户翻译的内容。
2. `completed`：已有 Target，以及自动完成的纯数字、纯符号和 Tag-only 行。

`strings` 的可见列：

```text
group | source | target | sample_sources | context | occurrences
```

`string_id`、原始行位置和原工作簿副本保存在隐藏内容中，不需要人工编辑。
`source` 中的 `{num1}` 和 Protected Token 是需要原样保留的清洗占位符。
`sample_sources` 显示模板首个原始 Source；`context` 显示同一原始行中的
参考语境，保证两列相互对应。这两列仅供翻译参考，不参与回填。

`completed` 的可见列：

```text
status | source | target | context
```

`existing_target` 表示原文件已有译文；`auto_passthrough` 表示 Source 会原样
写入 Target。该工作表只供复核，不需要用户处理。

## 相似句分组

相似句分组是默认关闭的可选功能。它只负责发现相似结构并调整显示顺序。

例如：

```text
Pikachu launched an attack
Squirtle launched an attack
Bulbasaur launched an attack
```

开启后会进入同一组，但三条 Source 仍然分别填写 Target。

分组的输入是清洗后的 Translation Unit，而不是原始行。模板变体只需翻译
一次，回填时 PhraseLoom 会按隐藏映射恢复每一行的原始变量值和 Tag。

不开启时，`group` 为空，所有 Unit 按原始首次出现顺序排列。开启后：

1. 未聚类 Unit 排在前面，并保持原始首次出现顺序。
2. 聚类 Unit 统一排在后面。
3. 同组内容连续排列；组与组内均按最早原始行排序。

相似句分组不抽取术语或结构变量，也不会用分组结果拼接译文。

## 2. 回填译文

译者填写 `strings` 工作表中的 Target 后运行：

```powershell
qatools phraseloom restore source_strings.xlsx
```

默认生成：

```text
source_translated.xlsx
```

回填时会：

- 将一条去重译文写回所有对应位置。
- 将 `{num1}` 等模板变量展开为每个原始 Source 对应的值。
- 将 Protected Token 恢复为原始 Tag 或占位符。
- 保留导出前已经存在的 Target。
- 将 `completed` 中的自动完成行写回空 Target。
- 恢复原工作表的顺序、可见状态和格式。
- 检查 Tag、变量和占位符是否遗漏或增加。
- 仅在存在空 Target 或保护内容问题时生成 `*_restore_issues.xlsx`。
