# Excel 术语检查工具

统一处理两种术语检查模式：

- 选择术语 mark：从 Excel 的 `source` / `target` 中提取新术语对，并结合历史 TB 回扫整表。
- 不选择术语 mark：必须提供历史 TB，仅检查文本中的历史术语是否使用了预期 target；该模式替代原“术语表命中检查”。

处理完成后会新增两个工作表：

- `术语表`：A/B 列为保留 mark 的 `source术语` / `target术语`，C/D 列为去 mark 后的 `source术语（无mark）` / `target术语（无mark）`，E 列为 `术语来源`
- `问题列`：A 列为问题行号，B 列为问题source术语，C 列为预期target术语，D 列为术语来源，E 列为问题简述，F/G 列为有问题的 source/target 原文；默认按 `问题source术语` 排序

## 规则

- 新术语只会由带所选 mark 的显式术语对触发发现；无 mark 模式不会新增术语
- 术语特征：被所选 mark 包裹的完整片段，输出到 `术语表` 时会同时保留带 mark 和去 mark 两套结果
- 当前支持两种术语 mark：`【】`、普通 `[]`
- `<...>` 和 `{...}` 不作为术语 mark，统一交给 tag / placeholder 检查
- 可同时选择多种 mark 类型，系统会按文本出现顺序提取并配对
- 选择 `[]` 时，也会兼容全角方括号 `［］`
- `[color=...]`、`[/color]` 这类方括号格式 tag 不进入术语表
- 同一行如果两侧都提取出多个术语，按顺序一一配对
- 术语检查时会忽略支持的 tag 外壳，按去 mark 后的纯术语进行匹配
- 可选传入历史 TB；历史 TB 会自动识别第 1 行的 `source` / `target` 列（也兼容 `source术语` / `target术语`），读取值时会去掉支持的 mark
- 传入历史 TB 后，会先把历史 TB 全量加入检查词库，再合并本批次新增术语；最终用“历史 TB 全量 + 本批次新增 TB”回扫整表
- 不选择任何 mark 时，历史 TB 为必填；程序跳过新术语提取，直接用历史 TB 回扫整表
- 本批次 source 如果命中历史 TB，会使用历史 target 作为配对和校验目标；未命中历史 TB 的 source 才按本批次第一次出现建立新增配对
- 输出的 `术语表` 不会把历史 TB 全量写出，只写本次检查文本中实际涉及的历史术语和本批次新增术语
- 一旦后面某行通过 tag 学到术语对，会回溯检查整张表中更早和更晚的未标注出现
- 复数处理：回扫时如果 `source` 命中整条 `术语+s`，且 `target` 也命中整条 `译法+s`，视为通过；其他复数形态疑似变体不进入问题报告
- 回扫检查默认使用混合边界：中文按包含匹配，英文/数字按边界匹配，避免 `ACC` 命中 `account`
- 术语表以去 mark 后的 `source` 为唯一键，首次出现的映射作为基准；不同 mark 但相同纯术语不会重复建条目
- 如果某一行 `source` 提取到了术语、但 `target` 没有对应术语，该 `source` 术语仍会写入术语表，`target` 留空
- 这类仅 `source` 有术语的空 target 条目不会参与后续“预期 target 命中”回扫校验
- 后续同一个纯 `source` 如果对应到不同的纯 `target`，该行记入 `问题列`
- 如果某一行两侧术语 mark 数量不一致，会先进入待复核；回扫时若所有 source 术语都能按已知映射在 target 中找到预期译法，则视为对齐
- target 侧额外的 mark 不单独算数量问题；它可能对应 source 中未加 mark 的表达。source 术语缺少预期 target 译法时仍会记入 `问题列`
- 如果同一个 Excel 行命中多个术语问题类型，会在 `问题列` 中重复写出多条问题
- 如果某一行两侧都没有提取到术语，则忽略

## 运行方式

```bash
python3 tools/term_pair_checker/extract_terms_from_excel.py input.xlsx -c A -t B
```

仅使用历史 TB 检查：

```bash
python3 tools/term_pair_checker/extract_terms_from_excel.py input.xlsx \
  -c A -t B \
  --no-term-mark \
  --history-tb glossary.xlsx \
  --history-sheet Glossary
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
- 可选选择历史 TB，并自动识别历史 TB 的工作表和 `source` / `target` 列
- 可以取消所有术语 mark；此时 GUI 会要求已选择历史 TB，并进入仅历史 TB 检查模式
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
- `--mark-style`：提取术语 mark 类型，可重复传入，例如 `--mark-style [] --mark-style '【】'`
- `--no-term-mark`：不从文本提取新术语；必须同时使用 `--history-tb`
- `--exclusion-config`：可选的自定义术语候选排除 JSON 配置文件路径
- `--history-tb`：历史 TB Excel 文件路径，可选
- `--history-sheet`：历史 TB 工作表名称，可选；默认优先使用 `术语表`
- `--history-source-column` / `--history-target-column`：历史 TB source / target 列，可选；不填则自动识别表头
- `--history-start-row`：历史 TB 开始读取行号，默认 `2`
- `-o, --output`：输出文件路径，可选，默认生成 `term_pair_check_<原文件名>`
