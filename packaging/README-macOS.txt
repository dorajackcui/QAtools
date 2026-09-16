QAtools macOS GUI 版
===================

安装
----
打开 DMG，把 QAtools.app 拖到 Applications，然后从应用程序中打开。
不需要安装 Python，不附带 CLI。更新时先退出旧版，再替换应用程序。
配置保存在当前用户目录，替换应用不会删除配置。

架构与系统
----------
arm64 包用于 Apple Silicon Mac；x86_64 包用于 Intel Mac，请匹配下载。
当前构建配置要求 macOS 13 或更新版本；具体系统兼容性需要在目标系统验证。

功能
----
一键质量检查、PhraseLoom 和直接读写 Excel 文件的工具可使用。
列操作、Excel 兼容性重存及同步后的可选 Excel 重存依赖 Windows COM，Mac 不支持，
即使安装了 Mac 版 Excel 也不能使用。相关页面仍保留，操作时会说明平台限制。

本地测试包
----------
此构建仅使用 ad-hoc 签名，没有 Apple Developer ID 签名和公证。
从网络下载后可能被 macOS 拦截；确认来源后，可在“系统设置 > 隐私与安全性”中
允许打开。不要关闭系统的安全检查。正式对外分发前应完成 Developer ID 签名和公证。

源码构建（在 Mac 上执行）
------------------------
python3 -m venv .venv
.venv/bin/python -m pip install -e . "pyinstaller>=6,<7" Pillow
.venv/bin/python scripts/build_macos_release.py

默认先运行完整回归测试，始终执行冻结应用的 GUI 启动检查和签名校验。
输出 dist/QAtools-v<版本>-macos-<架构>.dmg；.app 在 build/macos-release/bundle/。
使用构建 Python 的架构；Intel 包需要在 Intel Python 环境中另行构建验证。
已有同版本 DMG 仅在新包构建和验证成功后替换。--skip-tests 可跳过源码回归，
但不跳过冻结应用的启动检查。
