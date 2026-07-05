# Excel Tag / Placeholder 检查工具

用于读取一份双语 Excel，逐行检查 `source` / `target` 中的 `<...>`、`[color=...]` / `[/color]`、`{...}`、`\n` 和 memoQ tag 是否一致，并输出新的检查结果 Excel。

处理完成后会生成新的检查结果 Excel，并新增两个工作表：

- `标签占位问题`：列出 tag / placeholder 不一致的问题
- `检查汇总`：列出统计信息

## 规则

- 输入为一份双语 Excel：`source` / `target`
- 默认同时检查五类片段：
  - `<...>`：按 tag 处理
  - `[color=...]`、`[/color]`：按方括号 color tag 处理
  - `{...}`：按 placeholder 处理
  - `\n`：按独立 mark 处理
  - `{n}`、`{n>`、`<n}`：按 memoQ tag 处理，其中 `n` 为数字
- 逐行比较 source / target 中的片段集合
- `<...>` 默认全部按普通 tag 处理
- `{1}{2>Glace du Néant<3}` 会识别为 memoQ tag `{1}`、`{2>`、`<3}`，不会把整段 `{2>Glace du Néant<3}` 当成普通 `{...}` placeholder
- `memoq` 是独立检查类型；旧的 `numeric` 参数仍兼容，但不推荐继续使用
- 会检查：
  - target 缺少 source 中已有的片段
  - target 多出 source 中没有的片段
  - 同一片段数量不一致
- 当前按片段原样精确比较，例如 `<b>` 与 `</b>`、`{name}` 与 `{Name}` 会被视为不同内容

## 运行方式

```bash
python3 tools/tag_placeholder_checker/check_tags_and_placeholders.py input.xlsx \
  -s Sheet1 \
  -c A \
  -t B \
  --start-row 2 \
  --token-type angle \
  --token-type square_color \
  --token-type brace \
  --token-type newline \
  --token-type memoq
```

如需自定义尖括号 tag 过滤规则，可显式传入配置文件；不传时检查所有 `<...>`：

```bash
python3 tools/tag_placeholder_checker/check_tags_and_placeholders.py input.xlsx \
  -s Sheet1 \
  -c A \
  -t B \
  --start-row 2 \
  --token-type angle \
  --angle-config ./custom_angle_tags.json
```

图形界面：

```bash
python3 tools/tag_placeholder_checker/check_tags_and_placeholders_gui.py
```

GUI 支持：

- 自动读取工作表列表，用下拉框选择工作表
- 自动识别第 1 行表头中的 `source` / `target` 列并回填
- 默认输出为新的 Excel 文件，也可手动指定输出路径
- 可单独勾选 `<...>`、`[color=...]` / `[/color]`、`{...}`、`\n` 或 memoQ tag 类型进行检查
- GUI 中 memoQ tag 与其他检查类型互斥：勾选 memoQ 会取消普通 tag 组，勾选普通 tag 组会取消 memoQ

## 常用参数

- `--sheet`：工作表名称
- `--source-column`：source 列
- `--target-column`：target 列
- `--start-row`：开始处理的行号，默认 `2`
- `--token-type`：检查类型，可选 `angle`、`square_color`、`brace`、`newline`、`memoq`，可重复传入；`numeric` 作为旧别名兼容
- `--angle-config`：可选的尖括号 tag 过滤配置文件路径；不传时检查所有 `<...>`
- `-o, --output`：输出文件路径，可选，默认生成 `tag_check_<原文件名>`

## 输出说明

`标签占位问题` 工作表包含以下列：

- `行号`
- `问题类型`
- `描述`
- `source文本`
- `target文本`

`检查汇总` 工作表包含以下统计项：

- 检查工作表
- source / target 列
- 开始行
- 检查类型
- 总行数
- 命中检查类型行数
- 含尖括号 tag 行数
- 含方括号 color tag 行数
- 含花括号 placeholder 行数
- 含 `\n` mark 行数
- 含 memoQ tag 行数
- 问题行数
- 问题条数
