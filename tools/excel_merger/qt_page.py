"""Unified Qt page for ExcelMergerPage."""

from __future__ import annotations

from pathlib import Path
from PySide6.QtWidgets import QCheckBox, QFileDialog, QVBoxLayout, QWidget
from tools.excel_merger.merge_active_sheets import (
    build_default_output_path as build_merge_output_path,
    merge_active_sheets,
)
from tools.qt_gui_common import (
    AsyncPage,
    PathPicker,
    muted_label,
    primary_button,
    section,
    show_error,
    show_info,
    show_warning,
)
from tools.qt_page_helpers import _add_action_bar


class ExcelMergerPage(AsyncPage):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(2, 2, 6, 2)
        layout.setSpacing(8)
        input_box, input_layout = section("输入目录与表头")
        self.input_dir = PathPicker("Excel 所在目录", choose_text="选择目录")
        self.input_dir.choose_button.clicked.connect(self.choose_input_dir)
        self.input_dir.path_changed.connect(self.update_preview)
        self.keep_headers = QCheckBox("保留每个文件的表头")
        input_layout.addWidget(self.input_dir)
        input_layout.addWidget(self.keep_headers)
        layout.addWidget(input_box)
        self.run_button = primary_button("开始合并")
        self.run_button.clicked.connect(self.run_merge)
        self.status = muted_label()
        self.preview = muted_label()
        layout.addWidget(self.status)
        layout.addWidget(self.preview)
        layout.addStretch(1)
        self.action_bar = _add_action_bar(layout, self.run_button)

    def choose_input_dir(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "选择包含待合并 Excel 的目录", self.input_dir.path())
        if path:
            self.input_dir.set_path(path)

    def update_preview(self, _path: str = "") -> None:
        self.preview.setText(
            f"输出文件：{build_merge_output_path(Path(self.input_dir.path()))}"
            if self.input_dir.path()
            else ""
        )

    def run_merge(self) -> None:
        if not self.input_dir.path():
            show_error(self, "缺少目录", "请先选择包含 Excel 的目录。")
            return
        output = build_merge_output_path(Path(self.input_dir.path()))
        self.run_button.setEnabled(False)
        self.status.setText("正在读取 Excel…")
        self.run_in_background(
            merge_active_sheets,
            kwargs={
                "folder_path": self.input_dir.path(),
                "output_path": output,
                "keep_all_headers": self.keep_headers.isChecked(),
            },
            on_success=self._finish,
            on_error=self._fail,
        )

    def _finish(self, summary: object) -> None:
        self.run_button.setEnabled(True)
        self.status.clear()
        self.preview.setText(f"输出文件：{summary.output_path}")
        lines = [
            "活动工作表已合并。",
            f"输入文件数: {summary.supported_file_count}",
            f"输出行数: {summary.merged_row_count}",
            f"跳过 .xls/.xlsb: {summary.skipped_file_count}",
            f"读取失败: {summary.failed_file_count}",
            f"输出文件: {summary.output_path}",
        ]
        if summary.error_log_path is not None:
            lines.append(f"错误日志: {summary.error_log_path}")
        if summary.failed_file_count:
            show_warning(self, "合并完成（有文件失败）", "\n".join(lines))
        else:
            show_info(self, "合并完成", "\n".join(lines))

    def _fail(self, message: str) -> None:
        self.run_button.setEnabled(True)
        self.status.clear()
        show_error(self, "合并失败", message)
