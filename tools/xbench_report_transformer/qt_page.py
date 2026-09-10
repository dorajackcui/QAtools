"""Unified Qt page for XbenchPage."""

from __future__ import annotations

from pathlib import Path
from PySide6.QtWidgets import QComboBox, QHBoxLayout, QLabel, QVBoxLayout, QWidget
from tools.excel_metadata import list_workbook_sheets
from tools.qt_gui_common import (
    AsyncPage,
    PathPicker,
    muted_label,
    primary_button,
    section,
    show_error,
    show_info,
)
from tools.xbench_report_transformer.transform_xbench_report import (
    build_default_output_path as build_xbench_output_path,
    process_excel as transform_xbench_report,
)
from tools.qt_page_helpers import _add_action_bar, _choose_excel, _set_combo_values


class XbenchPage(AsyncPage):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(2, 2, 6, 2)
        layout.setSpacing(8)
        input_box, input_layout = section("输入与范围")
        self.input_picker = PathPicker("Xbench QA Report")
        self.input_picker.choose_button.clicked.connect(self.choose_input)
        input_layout.addWidget(self.input_picker)
        row = QHBoxLayout()
        self.sheet = QComboBox()
        self.sheet.setMinimumWidth(220)
        row.addWidget(QLabel("报告工作表"))
        row.addWidget(self.sheet)
        row.addStretch(1)
        input_layout.addLayout(row)
        layout.addWidget(input_box)
        self.run_button = primary_button("开始转换")
        self.run_button.clicked.connect(self.run_transform)
        self.preview = muted_label()
        layout.addWidget(self.preview)
        layout.addStretch(1)
        self.action_bar = _add_action_bar(layout, self.run_button)

    def choose_input(self) -> None:
        path = _choose_excel(self, "选择 Xbench QA Report Excel 文件")
        if not path:
            return
        self.input_picker.set_path(path)
        try:
            choices = list_workbook_sheets(path)
        except Exception as exc:  # noqa: BLE001
            show_error(self, "读取失败", str(exc))
            return
        _set_combo_values(self.sheet, choices.sheet_names, choices.default_sheet or "")
        self.preview.setText(f"输出文件：{build_xbench_output_path(Path(path)).name}")

    def run_transform(self) -> None:
        if not self.input_picker.path():
            show_error(self, "缺少文件", "请先选择 Xbench QA Report Excel 文件。")
            return
        self.run_button.setEnabled(False)
        self.run_in_background(
            transform_xbench_report,
            kwargs={"input_file": self.input_picker.path(), "sheet": self.sheet.currentText() or None, "output_file": None},
            on_success=self._finish,
            on_error=self._fail,
        )

    def _finish(self, summary: object) -> None:
        self.run_button.setEnabled(True)
        show_info(self, "处理完成", "\n".join([
            "Xbench QA Report 转换已完成。",
            f"工作表: {summary.worksheet_title}",
            f"读取明细数: {summary.detail_count}",
            f"输出行数: {summary.grouped_count}",
            f"输出文件: {summary.output_path}",
        ]))

    def _fail(self, message: str) -> None:
        self.run_button.setEnabled(True)
        show_error(self, "处理失败", message)
