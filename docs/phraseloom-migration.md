# PhraseLoom 合并记录

PhraseLoom 已作为一级 Strings 工作流并入 QAtools。

## 来源

- 原仓库：<https://github.com/dorajackcui/PhraseLoom>
- 导入分支：`main`
- 导入提交：`14af9603748575a2aee726608ba502d437462a0d`
- 导入日期：2026-07-30

导入使用保留历史的一次性 Git subtree 合并。QAtools 不保留需要持续同步的
PhraseLoom remote；合并后的 `phraseloom/` 是后续开发的唯一代码来源。

## 目录映射

- 核心包：`phraseloom/`
- PhraseLoom 测试：`tests/phraseloom_tests/`
- Strings 工作流设计：`docs/superpowers/specs/2026-07-30-strings-workflow-redesign.md`
- Toolshub 入口：`toolshub_gui.py`

PhraseLoom 的独立 GUI 仍可通过 `phraseloom gui` 或
`python -m phraseloom.gui` 启动；统一 GUI 中则复用同一个
`PhraseLoomApp` 页面。

## 后续维护

所有 PhraseLoom 修复和功能更新都应直接提交到 QAtools。旧仓库在 QAtools
版本验收后可设为只读归档，并在其 README 指向本仓库；不要在两个仓库继续
并行开发。
