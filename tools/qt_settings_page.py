"""Unified Qt page for SettingsPage."""

from __future__ import annotations

from PySide6.QtCore import Signal
from PySide6.QtWidgets import QGridLayout, QLabel, QPlainTextEdit, QPushButton, QVBoxLayout, QWidget
from tools.header_aliases import HeaderAliases, HeaderAliasStore
from tools.qt_gui_common import muted_label, primary_button, section, show_error
from tools.qt_page_helpers import _scroll_page, _add_action_bar


class SettingsPage(QWidget):
    settings_saved = Signal()

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        header_alias_store: HeaderAliasStore | None = None,
    ) -> None:
        super().__init__(parent)
        self.header_alias_store = header_alias_store or HeaderAliasStore()

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        content = QWidget()
        layout = QVBoxLayout(content)
        layout.setContentsMargins(2, 2, 6, 2)
        layout.setSpacing(8)

        alias_box, alias_layout = section("Excel 表头自动识别")
        editors = QGridLayout()
        editors.setHorizontalSpacing(14)
        editors.setVerticalSpacing(5)
        source_label = QLabel("Source 表头别名")
        target_label = QLabel("Target 表头别名")
        self.source_aliases = QPlainTextEdit()
        self.source_aliases.setObjectName("sourceHeaderAliases")
        self.source_aliases.setPlaceholderText("例如：原文\nEnglish\n源语言")
        self.source_aliases.setFixedHeight(120)
        self.target_aliases = QPlainTextEdit()
        self.target_aliases.setObjectName("targetHeaderAliases")
        self.target_aliases.setPlaceholderText("例如：译文\nChinese\n目标语言")
        self.target_aliases.setFixedHeight(120)
        editors.addWidget(source_label, 0, 0)
        editors.addWidget(target_label, 0, 1)
        editors.addWidget(self.source_aliases, 1, 0)
        editors.addWidget(self.target_aliases, 1, 1)
        editors.setColumnStretch(0, 1)
        editors.setColumnStretch(1, 1)
        alias_layout.addLayout(editors)
        layout.addWidget(alias_box)

        self.status = muted_label(word_wrap=True)
        layout.addWidget(self.status)
        layout.addStretch(1)
        outer.addWidget(_scroll_page(content), 1)

        self.reset_button = QPushButton("恢复默认")
        self.save_button = primary_button("保存设置")
        self.reset_button.clicked.connect(self.reset_aliases)
        self.save_button.clicked.connect(self.save_aliases)
        self.action_bar = _add_action_bar(
            outer,
            self.reset_button,
            self.save_button,
        )
        self.reload_aliases()

    @staticmethod
    def _editor_aliases(editor: QPlainTextEdit) -> tuple[str, ...]:
        return tuple(editor.toPlainText().splitlines())

    def _show_aliases(self, aliases: HeaderAliases) -> None:
        self.source_aliases.setPlainText("\n".join(aliases.source))
        self.target_aliases.setPlainText("\n".join(aliases.target))

    def reload_aliases(self) -> None:
        try:
            aliases = self.header_alias_store.load()
        except ValueError as exc:
            self.status.setText(str(exc))
            return
        self._show_aliases(aliases)
        self.status.clear()

    def reset_aliases(self) -> None:
        self._show_aliases(HeaderAliases())
        self.status.setText("已清空自定义别名；点击“保存设置”后生效。")

    def save_aliases(self) -> None:
        try:
            aliases = HeaderAliases.create(
                source=self._editor_aliases(self.source_aliases),
                target=self._editor_aliases(self.target_aliases),
            )
            self.header_alias_store.save(aliases)
        except (OSError, ValueError) as exc:
            show_error(self, "无法保存设置", str(exc))
            return
        self._show_aliases(aliases)
        self.status.setText("设置已保存，已重新识别当前加载工作簿的 Source / Target 列。")
        self.settings_saved.emit()
