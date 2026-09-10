# momoTools 功能合并方案

状态：历史设计。六个工具的业务层和统一 GUI 已接入；CLI 按用户要求延后。
下文保留调研时的建议，不作为当前实现依据。当前行为见
[内容同步](../../tools/content_sync/README.md)、[列操作](../../tools/column_tools/README.md)、
[兼容性重存](../../tools/excel_compatibility/README.md)、[文件替换](../../tools/deep_replace/README.md)、
[统计](../../tools/untranslated_stats/README.md)。

调研日期：2026-09-10。

- momoTools 基线：[`bb0a7a8`](https://github.com/dorajackcui/momoTools/tree/bb0a7a846edd8884fec637389fc2f609ac31b77a)。
- QAtools 基线：`dbe6e30642439a706dc8b14a8da994aa2802fc09`。
- 依据：上述版本的实现、相关测试和当前文档；本轮未运行 momoTools 或真实 Excel COM。

## 1. 范围与接入方式

合并六个用户入口：Master->Target、Target->Master、Column Clear、Compatibility、
Deep Replace、Untranslated Stats。

建议在 QAtools 中增加“内容同步”分组，放两个同步入口；四个 Utilities 放入
“Excel 实用工具”分组。新增页面使用现有 PySide6 风格、后台任务和底部操作栏。
同步共用一个业务包，界面分别呈现两个方向。

这次不纳入 momoTools 的 Batch 配置运行、Term Extractor 和整个 Master Update
分组。尤其不能把 Target->Master 误接到 Source Text / Translation 的处理器。
功能归属依据：[momoTools 页面注册](https://github.com/dorajackcui/momoTools/blob/bb0a7a846edd8884fec637389fc2f609ac31b77a/app_shell/registry.py)。

与现有工具的关系：

| 现有能力 | 与新功能的边界 |
| --- | --- |
| PhraseLoom | 按 Strings 内隐藏映射、模板变量和 Tag 回填；内容同步按 Key + 原文匹配任意总表/小表，独立保留 |
| Batch 拆分与复原 | 按 manifest 和原行位置复原；不能替代按 Key + 原文同步，也不迁入 momoTools 的另一个 Batch |
| 一键质量检查 | 负责检查和报告；同步、文件替换、列操作使用独立入口 |
| 现有 Excel/GUI 公共能力 | 复用列识别、工作表枚举、保留 VBA 的加载函数、输出命名和后台任务 |

现有契约分别见 [PhraseLoom](../../phraseloom/README.md)、
[Batch](../../tools/excel_batcher/README.md) 和 [workflow](../../tools/workflow/README.md)。

## 2. 六项功能的真实行为

### Master->Target

- 一个 Master 向递归目录内的小表写入内容；双方分别配置 Key、match（原文）、内容列。
- 当前界面已把单列和连续多列合为一个入口：更新列数为 1 时走单列处理器，大于 1
  时走多列处理器。迁移必须覆盖两种情况。
- Key 和 match 都转字符串并去首尾空白，任一为空则跳过；匹配大小写敏感。
- 内容保留首尾空白；默认覆盖已有内容，默认不写入空白来源值；可选择仅填空、允许空白写入。
- 多列模式逐个目标单元格判断“仅填空”和“空白来源”，不是整行决定。
- Master 重复身份由后行覆盖前行；目标所有匹配行都可能更新。
- 默认启用 Excel COM 兼容性重存，仅处理本轮发生写入的文件。
- 当前 Master 通过 pandas 默认读第一张表并把首行当表头；目标使用活动表且从第 1 行扫描。
  迁移时需显式统一工作表和表头设置，不能把这个不对称行为当成新默认。

依据：[单列实现](https://github.com/dorajackcui/momoTools/blob/bb0a7a846edd8884fec637389fc2f609ac31b77a/core/excel_processor.py)、
[多列实现](https://github.com/dorajackcui/momoTools/blob/bb0a7a846edd8884fec637389fc2f609ac31b77a/core/multi_column_processor.py)、
[界面](https://github.com/dorajackcui/momoTools/blob/bb0a7a846edd8884fec637389fc2f609ac31b77a/ui/views/updater.py)。

### Target->Master

- 递归读取小表的活动表，从第 2 行开始；按相同的 Key + match 身份更新 Master 活动表。
- 当前支持单列回填，不追加未匹配的新行。
- 默认覆盖、跳过空白来源；支持仅填空和允许空白写入。
- 小表路径按小写形式排序，后文件覆盖前文件；文件内后面的有效候选覆盖前面的候选。
  默认跳过空白值发生在候选收集阶段，所以后来的空白不会盖掉先前的有效译文。
- 某个小表读取失败时，仍应用其他成功读取的小表；Master 当前扫描包含第 1 行。
- 当前直接保存 Master 原文件。

依据：[反向同步实现](https://github.com/dorajackcui/momoTools/blob/bb0a7a846edd8884fec637389fc2f609ac31b77a/core/reverse_excel_processor.py)。

### Column Clear

实际包含三个动作：清空指定列第 2 行起的内容、在指定位置插入列并写表头
`Translation`、删除指定整列。三者都递归处理目录中的工作簿活动表，使用 Excel COM
并原地保存。清空使用 `ClearContents`，不等于删除列或清除格式。

建议 GUI 显示“列操作”，并保留英文名称辅助原用户识别。首版保留 COM 实现路径，
尤其插入和删除涉及公式引用、表结构变化，不能未经真实样例验证就替换为 openpyxl 操作。
无 Excel 的清空列实现可以后续单独评估，不列为本次必须交付。

依据：[列操作实现](https://github.com/dorajackcui/momoTools/blob/bb0a7a846edd8884fec637389fc2f609ac31b77a/core/excel_cleaner.py)、
[三个界面动作](https://github.com/dorajackcui/momoTools/blob/bb0a7a846edd8884fec637389fc2f609ac31b77a/ui/views/clearer.py)。

### Compatibility

递归枚举文件，用本机 Microsoft Excel 打开、保存、关闭。它提供 Excel 原生重存，
不是格式转换器，也没有检查或修复任意损坏工作簿的保证。建议显示“Excel 兼容性重存”。

独立入口与 Master->Target 的可选后处理共用一个 COM 服务，避免维护两份生命周期逻辑。

依据：[兼容性实现](https://github.com/dorajackcui/momoTools/blob/bb0a7a846edd8884fec637389fc2f609ac31b77a/core/excel_compatibility_processor.py)。

### Deep Replace

递归枚举源目录的 Excel 文件，在目标目录树中按完整文件名查找并替换整份文件。
当前只替换遍历时找到的第一个同名目标；成功后删除临时 `.bak`，复制失败时尝试恢复。
它不是单元格文字替换。建议显示“按文件名替换”。

迁移建议保留“同名整文件替换”用途；当源目录或目标目录有多个同名文件时，报告歧义并
跳过该文件，避免依赖不稳定的遍历顺序。单个文件失败后继续处理其他文件，并保留恢复失败
时的备份位置。此项是明确的行为调整，需单独验收。

依据：[文件替换实现](https://github.com/dorajackcui/momoTools/blob/bb0a7a846edd8884fec637389fc2f609ac31b77a/core/deep_replace_processor.py)。

### Untranslated Stats

- 递归统计小表活动表，跳过首行；配置 Source 和 Translation 列。
- 中文计数使用 `U+4E00–U+9FA5` 字符范围；英文使用字母词正则，包含内部英文撇号或连字符。
  不做 PhraseLoom 去重、模板化或可翻译字符清洗。
- 按文件输出未翻译字/词数、未翻译行数、总字/词数、总行数，并追加总计。
  报告工作表叫 `未翻译统计`，英文模式的表头使用“词数”。
- 旧逻辑把空值、纯空白及字面量 `nan`（不区分大小写）判为未翻译；这与同步模块
  把字面量 `nan` 视为有效内容的规则不同，不能无意合并为空值公共函数。
- 使用公式缓存值读取。没有缓存的公式可能表现为空；结果不等同于重新计算后的 Excel。
- 旧逻辑还会跳过数值 0 的 Source；文件读取失败返回全零结果。这两点建议修正：
  0 作为有效 Source 行计入行数，失败文件明确记录失败，不伪装成零工作量。

建议首版保留旧计数范围、英文正则、`nan` 判断及报告主体格式；如需改变统计口径，作为
独立选项或后续需求处理。报告成功数据和失败信息分开呈现。

依据：[统计实现](https://github.com/dorajackcui/momoTools/blob/bb0a7a846edd8884fec637389fc2f609ac31b77a/core/untranslated_stats_processor.py)。

## 3. 建议统一的运行契约

下表是迁移建议，并非已获确认的业务变更。

| 项目 | 建议 |
| --- | --- |
| 输出 | 默认另存：小表保留相对目录结构，反向同步生成新 Master；列操作、重存及文件替换先操作目标副本。原地更新仅在显式选择该模式时执行 |
| 文件格式 | 同步和统计首版使用 `.xlsx/.xlsm`，保留 VBA；COM 列操作/重存以及整文件复制另覆盖 `.xls`。界面按工具列出实际支持格式 |
| `.xls` 差异 | momoTools 共享枚举包含 `.xls`，但多个读取路径固定使用 openpyxl；不能因此宣称已有可靠的旧格式同步支持。同步旧格式列为待确认需求 |
| 文件集合 | 运行前固定递归候选清单，稳定排序；排除 `~$` 锁文件、输出、备份及同步任务的 Master；拒绝导致自覆盖或循环读取的路径组合 |
| 工作表与表头 | 默认活动表、1 行表头；Master 与小表可分别指定工作表和表头行数。指定表不存在时报该文件失败，不悄悄换表 |
| 列映射 | GUI/CLI 用 Excel 列字母，边界统一转换；内部用结构化映射。单列与连续多列共用正向同步引擎；不新增反向多列需求 |
| 身份 | 使用规范化后的 `(key, match)` 二元组，保持精确匹配，避免旧 `key + '|' + match` 在内容含竖线时碰撞 |
| 内容 | 同步保持字符串化、保留内容首尾空白；表头不写入。列重排后必须仍按配置映射，避免 pandas 选列顺序影响含义 |
| 重复身份 | 保留正向 Master 后行优先和反向后文件优先，并记录冲突来源与最终采用项；Deep Replace 的同名歧义单独按上一节处理 |
| 写入策略 | 仅填空与允许空白写入为两个独立开关，保持旧默认；已匹配但值相同计为“无变化”，与实际写入数区分 |
| 保存与失败 | 单文件先保存临时结果再替换输出；成功、失败、跳过、无变化分别计数。部分失败可保留成功结果，但 GUI 和 CLI 均明确报告部分失败 |
| COM | Windows 可选能力；在执行线程初始化/释放 COM，使用独立 Excel 实例，关闭本次打开的工作簿并退出该实例；只在触发相关操作时加载依赖 |
| 同步后处理 | 建议默认关闭；启用前验证 Excel 能力，并只重存本次写入输出。没有 Excel 的机器仍能启动整个 GUI 并运行非 COM 工具 |

公式和特殊对象需作为独立验收点：确认同步列里的公式是传递公式文本还是缓存值，验证
前导零、日期、外链、命名区域、图表、宏和隐藏表的保留范围。不能以“保存成功”代替验证。

首版不增加 pandas、xlrd 为通用运行依赖；优先用现有 openpyxl 完成读写，但须通过
新旧结果对照验证字符串化与列映射差异。`pywin32` 作为 Windows 可选依赖，在安装版中
单独验证打包；其存在不代表用户机器已安装 Excel。

### Content Sync：去掉 pandas，保留值搬运契约

此处的空白处理是必须继承的业务约束，不属于第三节待调整的默认值。
权威依据是 momoTools 的
[Excel IO Contract](https://github.com/dorajackcui/momoTools/blob/bb0a7a846edd8884fec637389fc2f609ac31b77a/docs/io-contract.md)
及其 [值转换实现](https://github.com/dorajackcui/momoTools/blob/bb0a7a846edd8884fec637389fc2f609ac31b77a/core/kernel/excel_io.py)。

**依赖判断。** 正向单列/多列仅用 pandas 读取 Master，随后遍历 `master_df.values`
建立 Python 字典；匹配和写回没有使用 DataFrame 的 merge、groupby 或向量化能力。
反向同步的处理器已直接使用 openpyxl。对当前功能，建议统一使用
“openpyxl 读取 → 明确的值转换与身份索引 → 按单元格更新”，去掉正向读取的 DataFrame 中间层。

pandas 的读取入口提供选列、表头解析和类型转换的便利，在复杂表格分析中也有价值；
但这些便利并不是当前同步业务对它的必要依赖。减少中间对象和依赖有明确的结构收益，
实际速度和峰值内存收益仍需测量，不能直接宣称 openpyxl 一定更快。

原正向代码已经设置 `na_filter=False`，避免把空串及 `NA`、`nan` 等文本自动识别为缺失值。
不能把移除 pandas 的理由写成“原代码必然把空白变成 NaN”；这是现有实现主动防范的事情。
该参数关闭时，`keep_default_na` 和 `na_values` 不再参与缺失值识别，见
[pandas 官方参数说明](https://pandas.pydata.org/docs/reference/api/pandas.read_excel.html)。

**必须保留的语义。** Python 的 `None` 与 Excel 中用户输入的文字 `"None"` 是两回事。
读取器内部用 `None` 表示空单元格可以接受；它必须转为空内容，不能经无条件 `str(value)`
写成文字 `"None"`，也不能为真实空白引入 `"nan"`、`"null"` 或自定义转义占位符。

| 读入值 | 同步内容转换结果 | 是否为空白来源 |
| --- | --- | --- |
| 空单元格对应的 Python `None` | 空字符串 | 是 |
| 空字符串 | 空字符串 | 是 |
| 纯空格或纯空白字符 | 保留原字符，不擅自裁剪 | 是；默认跳过，允许空白写入时原样写入 |
| 用户原本输入的 `None`、`nan`、`NA`、`NULL` 等文字 | 保留原文字及大小写 | 否 |
| 数值 0 / False | 按原契约字符串化为 `"0"` / `"False"` | 否 |
| 文本 `00123` | 保留 `00123` | 否 |
| 带首尾空格、换行、Tag 的正文 | 保留完整正文 | 否（含非空白字符时） |

这里继承的是原有字符串化契约，不承诺复制 Excel 的所有底层类型或显示格式。
例如数值 123 使用格式 `00000` 显示为 00123，与原本存储的文本 `00123` 不同；
不能凭“所见即所得”的简称擅自新增显示格式渲染逻辑。原契约还将传入的浮点 NaN
字符串化为 `"nan"`，但读取真实空单元格时不得自行制造 NaN 再写回。

空白判断、是否写入、正文转换三者分开处理：

- Key/match 保留原有去首尾空白规则；该规则不能应用到搬运正文。
- 默认 `allow_blank_write=False` 时，空白来源不触碰目标已有内容。
- 开启空白写入才允许空内容清空目标；`fill_blank_only` 继续独立限制目标是否可写。
- 禁止用 `value or ""` 归一化，以免误清除 0/False；禁止把文字 `nan`/`None` 当空值清洗。
- 不需要人为区分空字符串和空单元格在 Excel 文件里的序列化形式；重新读取后仍为空，
  且未产生占位文字，即满足空内容约束。

**验证边界。** 本轮用内存工作簿做了小型对照：本机 pandas 2.2.3 的原正向读取参数与
openpyxl 3.1.5 加显式转换，在空单元格、空串、纯空格、文字 `None`/`nan`/`NA`/`NULL`、
数值 0、文本前导零、首尾空白及换行样例上得到一致的内容字符串。
这不是完整兼容性或性能证明；正式迁移需把这些断言覆盖到正向单列、正向多列和反向同步，
并检查最终保存后重新打开的内容，而不仅比较内存对象。公式缓存、Excel 错误值、日期和
数值转换仍需单独对照，防止移除读取层时无意改变解释方式。

## 4. 代码落点与命令草案

迁入业务规则和测试样例，按 QAtools 包结构实现；GUI 控制器和 momoTools 应用壳无需搬入。

| 建议业务包 | 职责 | 建议命令入口（尚不可用） |
| --- | --- | --- |
| `tools/content_sync/` | 身份、列映射、正向单/多列、反向汇总、冲突报告 | `qatools content-sync master-to-target` / `target-to-master` |
| `tools/column_tools/` | 清空、插入、删除列 | `qatools column-tools clear` / `insert` / `delete` |
| `tools/excel_compatibility/` | Excel 原生重存 | `qatools excel-compatibility` |
| `tools/deep_replace/` | 同名文件索引、复制替换、恢复 | `qatools deep-replace` |
| `tools/untranslated_stats/` | 计数规则、文件汇总、报告 | `qatools untranslated-stats` |

只提取被多个工具实际需要的公共模块，例如递归文件清单/输出映射和 COM 会话管理；
不为本次迁移搭建通用插件系统或新的工作流框架。

接入点：

- `qatools/cli.py`：按现有延迟分发方式注册，参数完整说明进入 `docs/cli-usage.md`。
- `toolshub_gui.py`：增加分组及入口；`tools/qt_pages.py` 保持页面注册。
  新页面可放在相应工具包的 `qt_page.py`，避免把大量界面继续集中到一个文件。
- `tools/qt_gui_common.py`：复用 `AsyncPage` 和后台任务；不在 GUI 工作线程里直接操作控件。
- `tools/excel_metadata.py`、`tools/header_aliases.py`：复用 Source/Target 识别；Key 单独指定。
- `tools/excel_output.py`：复用保留 VBA 的加载、已有输出命名与报告辅助能力；新增写入规则
  先局限于新工具，避免顺带改变旧检查器。
- 每个新业务包配 README；项目概览和文档路由只增加链接，不复制完整规则。

## 5. 分阶段交付

第一期交付业务层、GUI、工具 README 和验收结果；CLI 按用户要求留到后续。
下表保留分期范围，已实现行为以对应工具 README 为准。

| 阶段 | 交付内容 | 完成门槛 |
| --- | --- | --- |
| P0：契约与样例 | 确定第三节的显式行为调整；建立匿名小型样例和新旧结果对照表；完成文件清单、输出映射的最小实现 | 能解释所有预期差异；确认 `.xls`、原地更新、COM 后处理的实际需求 |
| P1：正向同步（已接入 GUI） | Master->Target 单列及连续多列、写入选项、新目录输出、匹配/冲突/失败汇总；CLI 延后 | 精确身份、覆盖/仅填空/空白写入、乱序列映射均通过；不改动输入；兼容性后处理待 P5 接入 |
| P2：反向同步 | Target->Master 单列、稳定来源优先级、新 Master 输出 | 路径与线程完成顺序不影响结果；坏文件与冲突有明确结果；未匹配身份不新增行 |
| P3：统计 | Untranslated Stats、旧报告主体、失败清单 | 中文/英文口径和总计正确；公式缓存限制、`nan`、0 和失败文件行为有测试 |
| P4：文件替换 | Deep Replace、同名歧义检测、目标副本输出及显式原地模式 | 验证真正候选与替换清单一致；同名冲突不误替换；失败恢复可追溯 |
| P5：Excel 原生操作 | Column Clear 三动作、Compatibility、正向同步可选后处理 | COM mock 回归与真实 Excel 冒烟通过；无 Excel 环境不影响其他入口；安装版可启动 |

优先完成双向同步，使“分发总表内容—回收小表译文”的主流程先闭环。Utilities 再逐项加入；
COM 能力集中收尾，减少对前几阶段开发与验证环境的依赖。

## 6. 验收矩阵

原仓库可移植的行为测试位于
[核心回归](https://github.com/dorajackcui/momoTools/blob/bb0a7a846edd8884fec637389fc2f609ac31b77a/tests/test_core_processors_regression.py)
和 [COM 回归](https://github.com/dorajackcui/momoTools/blob/bb0a7a846edd8884fec637389fc2f609ac31b77a/tests/test_com_processors.py)。
保留业务断言，适配新输出位置；旧错误计数等缺陷断言应按已确认的新契约改写。

| 范围 | 必测场景 |
| --- | --- |
| 身份与值 | 同 Key 不同原文不匹配；缺一部分跳过；两部分含竖线不碰撞；首尾空白、0、`nan`、前导零、日期、换行 |
| 正向 | 单列/多列、非连续身份列、内容列在身份列之前、逐单元格仅填空、空白写入、重复身份、目标重复行 |
| 反向 | a/b 文件后者优先、后文件空白、同文件重复行、并发读取次序、坏文件继续、未知身份、Master 保存失败 |
| 工作簿 | 活动表非第一张、指定表缺失、首行与多行表头、公式及其缓存、隐藏表、格式、`.xlsm` VBA、被锁定文件 |
| 文件集合 | 深层目录、大写扩展名、锁文件、空目录、输出位于输入树、Master 被枚举、同路径/重叠路径、重复 basename |
| 统计 | 中文计数边界、英文撇号和连字符、混合语言、空白/`nan`/0、已翻译全文件、报告表名/表头/总计、失败信息 |
| 替换 | 替换后文件内容一致、目标树结构保留、源/目标同名歧义、复制失败回滚、恢复失败保留备份、没有匹配文件 |
| COM | 清空保留表头与格式、空表/UsedRange 起点偏移、插入表头、删除及引用变化、打开/保存失败、会话释放 |
| GUI/CLI | 命令唯一及帮助完整、页面构造不读文件或启动 Excel、任务后台运行、部分失败可见、未安装 pywin32/Excel 时其他功能可用 |

列操作旧实现把 `UsedRange.Rows.Count` 当作最后行号，也把尝试文件数当成处理数；
新实现需用实际范围末行并区分成功/失败，测试不能照搬这两个问题。

性能验证用可重复生成的工作簿，记录行数/文件数/列宽、读取与保存耗时和峰值内存。
避免按格式撑大的空白范围全表遍历；Master 索引只构建一次，反向 Master 只保存一次。
先与相同样例的原版比较，再依据实际规模决定是否需要并行，不预先承诺未经测量的加速比。

每次实现先运行相关测试，再按仓库要求运行：

```powershell
python -m unittest discover -s tests -v
python -m compileall -q qatools phraseloom tools tests
git diff --check
```

COM mock 不能替代真实 Microsoft Excel 验收；最终发布需覆盖有 Excel 和无 Excel 两种环境。

## 7. 实施前的少量业务决策

1. **输出习惯**：建议默认新文件/新目录，并提供显式原地模式；确认旧流程是否必须默认原地更新。
2. **旧格式需求**：建议同步/统计首版覆盖 `.xlsx/.xlsm`；若实际大量使用 `.xls`，需把可靠的
   转换或专用读写路径放到 P0/P1，并重新评估依赖。
3. **同步后处理默认值**：建议默认关闭、按需调用 Excel 重存；如果小表必须经 Excel
   重存才能进入下游，应提前交付 COM 服务并保留默认启用。

其余建议规则可在 P0 的对照表逐项落定。完成迁移后，将最终业务规则归入各工具 README，
本计划移入 `docs/archive/`，更新文档路由，避免方案与当前说明长期并存为两套事实。
