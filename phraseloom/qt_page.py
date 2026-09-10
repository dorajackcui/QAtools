"""Unified Qt page for PhraseLoomPage."""

from __future__ import annotations

from PySide6.QtWidgets import (
    QCheckBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QProgressBar,
    QPushButton,
    QVBoxLayout,
    QWidget,
)
from phraseloom.strings_workflow import (
    default_restored_output_path,
    default_strings_output_path,
    export_strings_workbook,
    restore_strings_workbook,
)
from tools.qt_gui_common import (
    AsyncPage,
    PathPicker,
    muted_label,
    primary_button,
    section,
    show_error,
    show_info,
)
from tools.qt_page_helpers import _scroll_page, _add_action_bar, _choose_excel


def _result_text(stats: dict[str, int | str], *, restore: bool) -> str:
    if restore:
        lines = [
            f"输出文件: {stats['output_path']}",
            f"已回填 Source 行: {stats['restored_row_count']}",
            f"问题数: {stats['issue_count']}",
        ]
        if "audit_output_path" in stats:
            lines.append(f"复核文件: {stats['audit_output_path']}")
        return "\n".join(lines)
    return "\n".join(
        [
            f"输出文件: {stats['output_path']}",
            f"待翻译 Strings: {stats['string_count']}",
            f"待处理 Source 行: {stats['pending_row_count']}",
            f"已有译文跳过: {stats['completed_row_count']}",
            f"相似句分组: {stats['group_count'] if stats['grouping_enabled'] else '关闭'}",
        ]
    )



class PhraseLoomPage(AsyncPage):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        content = QWidget()
        layout = QVBoxLayout(content)
        layout.setContentsMargins(2, 2, 6, 2)
        layout.setSpacing(8)

        input_box, input_layout = section("输入与范围")
        self.input_picker = PathPicker("原始 Excel")
        self.input_picker.choose_button.clicked.connect(self.choose_input)
        self.input_picker.path_changed.connect(self.update_preview)
        input_layout.addWidget(self.input_picker)
        columns = QHBoxLayout()
        self.source_column = QLineEdit("source")
        self.target_column = QLineEdit("target")
        self.context_column = QLineEdit()
        self.context_column.setPlaceholderText("可选")
        for label, field in (
            ("Source 列", self.source_column),
            ("Target 列", self.target_column),
            ("Context 列", self.context_column),
        ):
            columns.addWidget(QLabel(label))
            columns.addWidget(field)
        columns.addStretch(1)
        input_layout.addLayout(columns)
        layout.addWidget(input_box)

        options_box, options_layout = section("导出选项")
        self.split_lines = QCheckBox("按换行拆分多行 Source（回填时自动合并）")
        self.split_lines.setChecked(True)
        self.group_similar = QCheckBox("启用相似句分组（未聚类在前，聚类内容在后）")
        self.tag_picker = PathPicker("Tag 配置（可选）", allow_clear=True)
        self.tag_picker.choose_button.clicked.connect(self.choose_tag_config)
        options_layout.addWidget(self.split_lines)
        options_layout.addWidget(self.group_similar)
        options_layout.addWidget(self.tag_picker)
        layout.addWidget(options_box)

        self.export_button = primary_button("导出 Strings")
        self.restore_button = QPushButton("回填译文…")
        self.export_button.clicked.connect(self.run_export)
        self.restore_button.clicked.connect(self.choose_and_restore)
        self.preview = muted_label()
        layout.addWidget(self.preview)
        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.hide()
        layout.addWidget(self.progress)
        layout.addStretch(1)
        outer.addWidget(_scroll_page(content), 1)
        self.action_bar = _add_action_bar(
            outer,
            self.restore_button,
            self.export_button,
        )

    def choose_input(self) -> None:
        if path := _choose_excel(self, "选择原始 Excel"):
            self.input_picker.set_path(path)

    def choose_tag_config(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "选择 Tag 配置",
            "",
            "TOML 配置 (*.toml);;所有文件 (*)",
        )
        if path:
            self.tag_picker.set_path(path)

    def update_preview(self, _path: str = "") -> None:
        path = self.input_picker.path()
        self.preview.setText(
            f"输出文件：{default_strings_output_path(path).name}" if path else ""
        )

    def _set_running(self, running: bool) -> None:
        self.export_button.setEnabled(not running)
        self.restore_button.setEnabled(not running)
        self.progress.setVisible(running)

    def run_export(self) -> None:
        input_path = self.input_picker.path()
        if not input_path:
            show_error(self, "无法开始", "请选择原始 Excel。")
            return
        self._set_running(True)
        self.run_in_background(
            export_strings_workbook,
            args=(input_path,),
            kwargs={
                "source_col": self.source_column.text().strip() or "source",
                "target_col": self.target_column.text().strip() or "target",
                "context_col": self.context_column.text().strip() or None,
                "group_similar": self.group_similar.isChecked(),
                "tag_config": self.tag_picker.path() or None,
                "split_lines": self.split_lines.isChecked(),
            },
            on_success=self._finish_export,
            on_error=lambda message: self._finish_error("导出失败", message),
        )

    def choose_and_restore(self) -> None:
        path = _choose_excel(self, "选择翻译完成的 Strings 工作簿")
        if not path:
            return
        try:
            self.preview.setText(f"回填输出：{default_restored_output_path(path).name}")
        except Exception:
            self.preview.setText("")
        self._set_running(True)
        self.run_in_background(
            restore_strings_workbook,
            args=(path,),
            on_success=self._finish_restore,
            on_error=lambda message: self._finish_error("回填失败", message),
        )

    def _finish_export(self, stats: object) -> None:
        self._set_running(False)
        self.update_preview()
        show_info(self, "导出完成", _result_text(dict(stats), restore=False))

    def _finish_restore(self, stats: object) -> None:
        self._set_running(False)
        show_info(self, "回填完成", _result_text(dict(stats), restore=True))

    def _finish_error(self, title: str, message: str) -> None:
        self._set_running(False)
        show_error(self, title, message)
