QAtools Windows 安装版
=====================

系统要求
--------
- 64 位 Windows 10 或 Windows 11
- 不需要安装 Python，也不需要联网

安装与升级
----------
双击 QAtools 安装包，按向导完成安装。程序默认安装到当前用户目录，不要求管理员
权限。

以后收到新版本时，直接运行新安装包即可覆盖升级。安装程序会沿用原安装目录和
快捷方式；历史 TB 等用户配置保存在 %APPDATA%\Toolshub，不会因升级被删除。

图形界面
--------
从开始菜单或桌面快捷方式启动 QAtools。程序只会保留一个图形界面实例；再次启动
不会创建第二个窗口。

工具箱包含一键质量检查、PhraseLoom、法语 NBSP 恢复、Batch 拆分与复原、
活动工作表合并和 Xbench QA 转换。一键质量检查内含术语、双向文本一致性、
Tag / Placeholder、换行、数字、URL 和 Target 文本质量检查。

命令行
------
安装目录中同时提供 QAtools-CLI.exe。示例：

  QAtools-CLI.exe --help
  QAtools-CLI.exe list
  QAtools-CLI.exe qa input.xlsx -c A -t B
  QAtools-CLI.exe phraseloom export source.xlsx
  QAtools-CLI.exe batch split input.xlsx --batch-size 1000
  QAtools-CLI.exe batch restore input_batches -o input_restored.xlsx
  QAtools-CLI.exe merge-sheets excel_files -o merged.xlsx

也可以双击 QAtools-CLI.cmd 查看帮助。

卸载
----
在 Windows“设置 > 应用”中找到 QAtools 并卸载。卸载程序不会删除
%APPDATA%\Toolshub 中的用户配置。

注意事项
--------
- 程序没有代码签名；Windows SmartScreen 若显示未知发布者，请确认安装包来源后
  再选择运行。
- 不要手动移动安装目录中的 QAtools.exe 或 _internal；请使用安装程序升级，使用
  Windows“设置 > 应用”卸载。
