QAtools Windows 安装版
=====================

系统要求
--------
- 64 位 Windows 10 或 Windows 11
- 不需要安装 Python，也不需要联网
- 列操作、兼容性重存和同步后的可选重存需要已安装的桌面 Microsoft Excel；其他工具无需 Excel

安装与升级
----------
双击 QAtools 安装包，按向导完成安装。程序默认安装到当前用户目录，不要求管理员
权限。

以后收到新版本时，直接运行新安装包即可覆盖升级。安装程序会沿用原安装目录和
快捷方式；历史 TB 等用户配置保存在 %APPDATA%\Toolshub，不会因升级被删除。
从 0.1.2 起安装包只提供 GUI。覆盖安装会删除安装目录中的旧 QAtools-CLI.exe、
QAtools-CLI.cmd，以及 _internal 下不再需要的 NumPy/OpenBLAS、PythonWin 和 YAML 组件。
其他用户文件和配置不会被这项清理删除。
从 0.1.6 起使用圆角 Q 图标；桌面和开始菜单快捷方式引用带版本号的图标资源，
降低覆盖升级后仍显示旧图标的缓存问题。

图形界面
--------
从开始菜单或桌面快捷方式启动 QAtools。程序只会保留一个图形界面实例；再次启动
不会创建第二个窗口。

工具箱包含一键质量检查、PhraseLoom、法语 NBSP 恢复、Batch 拆分与复原、
活动工作表合并、Xbench QA 转换、Master 与小表双向同步、列操作、兼容性重存、
同名文件替换、文件提取和未翻译统计。一键质量检查内含术语、双向及子串译文一致性、
Tag / Placeholder、换行、数字、URL 和 Target 文本质量检查。

资源管理器右键菜单（0.1.7 起）
----------------------------
安装时默认勾选“添加 Excel 文件和文件夹右键菜单”，只注册当前用户，无需管理员权限。
Windows 11 通常在“显示更多选项”中显示这些菜单；每次选择一个文件或文件夹。

右键 .xlsx / .xlsm 文件：
- 一键质量检查：打开页面并载入所选文件，确认配置后点击运行。
- 法语 NBSP 修复：载入文件后直接执行，结果规则见仓库 tools/french_nbsp_restorer/README.md。

右键文件夹：
- 兼容性重存
- 合并表格
- 统计未翻译（打开“未翻译统计”页面）
这三项均打开对应页面并填入所选目录，由用户确认配置后点击运行。

程序已打开时复用当前窗口；对应工具正在处理文件时提示等待，不替换其输入。
文件关联保持不变，双击 Excel 仍使用原来的默认应用。覆盖安装更新菜单命令；
重跑安装器并取消右键菜单选项可以移除这些入口，卸载也会清理。

命令行
------
此安装包不附带 CLI。需要命令行的用户可安装 Python 3.11+，在源码目录执行
python -m pip install -e .，再使用 qatools 命令。完整参数见仓库 docs/cli-usage.md。

打包依赖
--------
保留 PySide6、openpyxl、pyahocorasick、Excel COM 及 Pillow（用于保留工作簿图片）。
排除规则维护在 packaging/gui-excludes.txt：Tk、NumPy（含 OpenBLAS）、PyYAML、
PythonWin 界面模块均非本工具 GUI 工作流所需；源码 CLI 和业务处理器仍保留。

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
