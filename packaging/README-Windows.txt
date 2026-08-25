QAtools Windows 便携版
=====================

系统要求
--------
- 64 位 Windows 10 或 Windows 11
- 不需要安装 Python，也不需要联网

图形界面
--------
双击 QAtools.exe 即可打开统一工具箱。

工具箱包含一键质量检查、PhraseLoom、术语检查、Tag 检查、换行检查、
同源译文一致性检查、Target 中文检查、Excel 分行、Excel batch 拆分与复原、
活动工作表合并、法语 NBSP 恢复和 Xbench QA 转换。

命令行
------
在当前文件夹的地址栏输入 cmd 后回车，然后运行：

  QAtools-CLI.exe --help
  QAtools-CLI.exe list
  QAtools-CLI.exe qa input.xlsx -c A -t B
  QAtools-CLI.exe phraseloom export source.xlsx
  QAtools-CLI.exe batch split input.xlsx --batch-size 1000
  QAtools-CLI.exe batch restore input_batches -o input_restored.xlsx
  QAtools-CLI.exe merge-sheets excel_files -o merged.xlsx

也可以双击 QAtools-CLI.cmd 查看帮助。传递参数时，QAtools-CLI.cmd 与
QAtools-CLI.exe 的用法相同。

文件说明
--------
- QAtools.exe：统一图形界面
- QAtools-CLI.exe：统一命令行
- QAtools-CLI.cmd：便于在 cmd 中调用的包装脚本
- VERSION.txt：版本与构建信息
- SHA256SUMS.txt：可执行文件校验值

注意事项
--------
- 第一次启动单文件程序时，Windows 可能需要几秒钟解压运行文件。
- 程序没有代码签名；Windows SmartScreen 若显示未知发布者，请确认文件来源和
  SHA256 校验值后再选择运行。
- 历史 TB 项目配置保存在当前 Windows 用户的
  %APPDATA%\Toolshub\tb_projects.json，不会写入程序目录。
