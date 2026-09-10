"""Unified Qt page for ExcelBatcherPage."""

from __future__ import annotations

from pathlib import Path
from PySide6.QtWidgets import (
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QSpinBox,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)
from tools.excel_batcher.excel_batcher import (
    build_default_output_dir as build_batch_output_dir,
    build_default_restore_path,
    restore_batches,
    split_workbook,
)
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
from tools.qt_page_helpers import _add_action_bar, _choose_excel, _set_combo_values, EXCEL_FILTER


class ExcelBatcherPage(AsyncPage):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(2, 2, 6, 2)
        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_split_tab(), "拆分 batch")
        self.tabs.addTab(self._build_restore_tab(), "复原文件")
        layout.addWidget(self.tabs)

    def _build_split_tab(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(8)
        input_box, input_layout = section("输入与范围")
        self.split_input = PathPicker("输入 Excel")
        self.split_input.choose_button.clicked.connect(self.choose_split_input)
        input_layout.addWidget(self.split_input)
        row = QHBoxLayout()
        self.split_sheet = QComboBox()
        self.split_sheet.setMinimumWidth(160)
        self.batch_size = QSpinBox()
        self.batch_size.setRange(1, 1_000_000)
        self.batch_size.setValue(1000)
        self.batch_size.setMaximumWidth(86)
        self.header_rows = QSpinBox()
        self.header_rows.setRange(0, 1_000_000)
        self.header_rows.setValue(1)
        self.header_rows.setMaximumWidth(86)
        for label, widget in (
            ("工作表", self.split_sheet),
            ("每批行数", self.batch_size),
            ("表头行数", self.header_rows),
        ):
            row.addWidget(QLabel(label))
            row.addWidget(widget)
            row.addSpacing(6)
        row.addStretch(1)
        input_layout.addLayout(row)
        layout.addWidget(input_box)

        output_box, output_layout = section("输出")
        self.split_output = PathPicker("batch 目录（可选）", choose_text="选择目录", allow_clear=True)
        self.split_output.choose_button.clicked.connect(self.choose_split_output)
        self.split_output.path_changed.connect(self.update_split_preview)
        output_layout.addWidget(self.split_output)
        layout.addWidget(output_box)
        self.split_button = primary_button("开始拆分")
        self.split_button.clicked.connect(self.run_split)
        self.split_status = muted_label()
        layout.addWidget(self.split_status)
        self.split_preview = muted_label()
        layout.addWidget(self.split_preview)
        layout.addStretch(1)
        self.split_action_bar = _add_action_bar(layout, self.split_button)
        return page

    def _build_restore_tab(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(8)
        input_box, input_layout = section("batch 文件")
        self.restore_dir = PathPicker("batch 目录", choose_text="选择目录")
        self.restore_dir.choose_button.clicked.connect(self.choose_restore_dir)
        input_layout.addWidget(self.restore_dir)
        layout.addWidget(input_box)
        output_box, output_layout = section("输出")
        self.restore_output = PathPicker("复原 Excel（可选）", allow_clear=True)
        self.restore_output.choose_button.clicked.connect(self.choose_restore_output)
        output_layout.addWidget(self.restore_output)
        layout.addWidget(output_box)
        self.restore_button = primary_button("开始复原")
        self.restore_button.clicked.connect(self.run_restore)
        self.restore_status = muted_label()
        layout.addWidget(self.restore_status)
        layout.addStretch(1)
        self.restore_action_bar = _add_action_bar(layout, self.restore_button)
        return page

    def choose_split_input(self) -> None:
        path = _choose_excel(self, "选择要拆分的 Excel 文件")
        if not path:
            return
        self.split_input.set_path(path)
        try:
            choices = list_workbook_sheets(path)
        except Exception as exc:  # noqa: BLE001
            _set_combo_values(self.split_sheet, ())
            show_error(self, "读取失败", str(exc))
            return
        _set_combo_values(self.split_sheet, choices.sheet_names, choices.default_sheet or "")
        self.update_split_preview()

    def choose_split_output(self) -> None:
        initial = str(Path(self.split_input.path()).parent) if self.split_input.path() else ""
        path = QFileDialog.getExistingDirectory(self, "选择空的 batch 输出目录", initial)
        if path:
            self.split_output.set_path(path)

    def update_split_preview(self, _path: str = "") -> None:
        if self.split_output.path():
            output = self.split_output.path()
        elif self.split_input.path():
            output = str(build_batch_output_dir(self.split_input.path()))
        else:
            output = ""
        self.split_preview.setText(f"输出目录：{output}" if output else "")

    def run_split(self) -> None:
        if not self.split_input.path():
            show_error(self, "缺少文件", "请先选择要拆分的 Excel 文件。")
            return
        options = {
            "input_file": self.split_input.path(),
            "sheet": self.split_sheet.currentText() or None,
            "batch_size": self.batch_size.value(),
            "header_rows": self.header_rows.value(),
            "output_dir": self.split_output.path() or None,
        }
        self.split_button.setEnabled(False)
        self.split_status.setText("正在读取 Excel…")

        def task() -> object:
            return split_workbook(**options)

        self.run_in_background(task, on_success=self._finish_split, on_error=self._fail_split)

    def _finish_split(self, summary: object) -> None:
        self.split_button.setEnabled(True)
        self.split_status.clear()
        self.split_output.set_path(str(summary.output_dir))
        show_info(self, "拆分完成", "\n".join([
            "Excel batch 已生成。",
            f"工作表: {summary.worksheet_title}",
            f"数据行数: {summary.data_row_count}",
            f"batch 数: {summary.batch_count}",
            f"输出目录: {summary.output_dir}",
        ]))

    def _fail_split(self, message: str) -> None:
        self.split_button.setEnabled(True)
        self.split_status.clear()
        show_error(self, "拆分失败", message)

    def choose_restore_dir(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "选择包含 batch manifest 的目录")
        if path:
            self.restore_dir.set_path(path)

    def choose_restore_output(self) -> None:
        suggested: Path | None = None
        if self.restore_dir.path():
            try:
                suggested = build_default_restore_path(self.restore_dir.path())
            except Exception:
                suggested = None
        path, _ = QFileDialog.getSaveFileName(
            self,
            "选择复原 Excel 文件",
            str(suggested) if suggested else "",
            EXCEL_FILTER,
        )
        if path:
            self.restore_output.set_path(path)

    def run_restore(self) -> None:
        if not self.restore_dir.path():
            show_error(self, "缺少目录", "请先选择 batch 目录。")
            return
        self.restore_button.setEnabled(False)
        self.restore_status.setText("正在读取 batch…")
        self.run_in_background(
            restore_batches,
            kwargs={
                "manifest_or_directory": self.restore_dir.path(),
                "output_file": self.restore_output.path() or None,
            },
            on_success=self._finish_restore,
            on_error=self._fail_restore,
        )

    def _finish_restore(self, summary: object) -> None:
        self.restore_button.setEnabled(True)
        self.restore_status.clear()
        self.restore_output.set_path(str(summary.output_path))
        show_info(self, "复原完成", "\n".join([
            "batch 已复原为完整 Excel。",
            f"工作表: {summary.worksheet_title}",
            f"batch 数: {summary.batch_count}",
            f"复原数据行数: {summary.restored_row_count}",
            f"输出文件: {summary.output_path}",
        ]))

    def _fail_restore(self, message: str) -> None:
        self.restore_button.setEnabled(True)
        self.restore_status.clear()
        show_error(self, "复原失败", message)
