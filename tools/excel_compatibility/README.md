# 兼容性重存（Compatibility）

GUI【文件与表格 → 兼容性重存】用 Microsoft Excel 打开工作簿，再按原格式保存。
用于需要经过 Excel 原生保存的下游流程；不改变文件扩展名，也不自动修复所有损坏工作簿。
CLI 为 `qatools compatibility`，参数见[CLI 手册](../../docs/cli-usage.md#列操作与兼容性重存)。

选择文件目录，点击【兼容性重存】。新输出目录默认留空，直接重存原文件；
填写尚不存在的新输出目录时按相对路径另存副本。输出可手动清空，更换输入不会自动填充输出。

## 环境

- 需要 Windows、已安装且可以正常启动的桌面 Microsoft Excel，以及 pywin32。
- 源码安装：`python -m pip install -e ".[excel-com]"`。Windows 构建脚本检查并打包 COM 依赖，
  但安装包不附带 Microsoft Excel。
- 没有 Excel 或 pywin32 时仍可打开所有 GUI 页面，使用双向同步、统计、文件替换及原有非 COM 工具。
  执行原生操作时才检查环境；启动失败不会创建结果目录。

## 原生操作的共同规则

列操作、独立兼容性重存和正向同步可选重存共用 `tools/excel_com.py`：

- 在执行线程初始化和释放 COM，创建独立的隐藏 Excel 实例；不连接或退出用户已有 Excel 会话。
- 关闭交互提示、事件、链接更新及程序化打开文件时的 VBA 宏执行；不依靠宏完成任务。
- 先复制到临时目录，Excel 保存并关闭该副本成功后才替换原文件或移到另存路径。打开、编辑或保存失败时，
  原文件和已有结果不会被半成品替换；关闭本次工作簿，最后退出独立实例。
- 批处理支持 `.xlsx/.xlsm/.xls`，`.xlsb` 记录为跳过。递归清单稳定排序，忽略 `~$` 锁文件、
  `.qatools-backups` 备份树、另存输出及本次应用会话已生成的结果子目录，其他格式不参与。
- 每个失败文件保留具体错误，继续其余文件。【查看日志】按文件显示单行状态，
  保留跳过和失败原因；输出位置每次任务显示一次，Excel 退出异常单独提示警告。输出目录不生成附加 report。
  日志与历史输出识别规则见[运行日志](../../docs/tool-logs.md)。
- Excel 可能重算公式并规范化工作簿内部结构，输出不保证逐字节相同；密码、只读、保护等限制会报错。

真实 Excel 冒烟验证（仅使用生成的临时工作簿）：

```powershell
$env:QATOOLS_TEST_EXCEL_COM='1'
python -m unittest discover -s tests -p 'test_momotools_utilities.py' -v
```

普通回归使用模拟 COM；设置以上环境变量才执行真实 Excel 测试。
