# Excel 术语对检查工具

用于从 Excel 的 `source` 列和 `target` 列同时提取术语，检查术语对是否对齐，并输出为新的 Excel 文件。

处理完成后会新增两个工作表：

- `术语表`：A 列为 `source术语`，B 列为 `target术语`
- `术语表（无mark）`：A 列为去 mark 后的 `source术语`，B 列为去 mark 后的 `target术语`
- `问题列`：A 列为问题行号，B 列为问题类型，C 列为问题简述，D/E 列为有问题的 source/target 原文

## 规则

- 新术语只会由带所选 tag 的显式术语对触发发现
- 术语特征：被所选 tag 包裹的完整片段，输出到 `术语表` 时会保留外层 tag 本身
- 当前支持三种 tag：`【】`、`[]`、`<>`
- 可同时选择多种 tag 类型，系统会按文本出现顺序提取并配对
- 选择 `[]` 时，也会兼容全角方括号 `［］`
- 选择 `<>` 时，也会兼容全角尖括号 `＜＞`
- 同一行如果两侧都提取出多个术语，按顺序一一配对
- 术语检查时会忽略支持的 tag 外壳，按去 mark 后的纯术语进行匹配
- 一旦后面某行通过 tag 学到术语对，会回溯检查整张表中更早和更晚的未标注出现
- 回扫检查默认使用混合边界：中文按包含匹配，英文/数字按边界匹配，避免 `ACC` 命中 `account`
- 默认会通过 `false_positive_exclusions.json` 排除 `</>`、`<color=...>`、`<outline color=...>` 这类伪标签误判
- 如需新增或调整排除规则，直接编辑 `tools/term_pair_checker/false_positive_exclusions.json`
- 术语表以去 mark 后的 `source` 为唯一键，首次出现的映射作为基准；不同 mark 但相同纯术语不会重复建条目
- 后续同一个纯 `source` 如果对应到不同的纯 `target`，该行记入 `问题列`
- 如果某一行两侧术语数量不一致，或只有一侧提取到术语，该行记入 `问题列`
- 如果某一行两侧都没有提取到术语，则忽略

## 运行方式

```bash
python3 tools/term_pair_checker/extract_terms_from_excel.py input.xlsx -c A -t B
```

兼容旧入口：

```bash
python3 extract_terms_from_excel.py input.xlsx -c A -t B
```

图形界面：

```bash
python3 tools/term_pair_checker/extract_terms_gui.py
```

GUI 现在支持：

- 选择 Excel 后自动读取工作表列表，用下拉框选择工作表
- 默认输出为新的 Excel 文件，也可手动指定输出路径
- 自动识别第 1 行表头中的 `source` / `target` 列并回填
- 如果未识别到列，可继续手动填写列字母作为回退

兼容旧入口：

```bash
python3 extract_terms_gui.py
```

## 常用参数

- `-c, --source-column`：source 列，例如 `A`
- `-t, --target-column`：target 列，例如 `B`
- `-s, --sheet`：工作表名称，可选
- `--start-row`：从第几行开始处理，默认 `2`
- `--mark-style`：提取 tag 类型，可重复传入，例如 `--mark-style [] --mark-style <>`
- `--exclusion-config`：误判排除 JSON 配置文件路径；默认读取工具目录下的 `false_positive_exclusions.json`
- `-o, --output`：输出文件路径，可选，默认生成 `<原文件名>_term_pairs.xlsx`
