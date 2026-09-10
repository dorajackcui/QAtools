"""Content Sync page using the shared Toolshub widgets and background worker."""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Signal, Qt
from PySide6.QtWidgets import (
    QCheckBox, QComboBox, QFileDialog, QFrame, QGridLayout, QHBoxLayout, QLabel,
    QLineEdit, QPlainTextEdit, QProgressBar, QPushButton, QScrollArea, QSpinBox, QTabWidget, QVBoxLayout, QWidget,
)

from tools.qt_gui_common import (
    AsyncPage, PathPicker, muted_label, primary_button, section, show_error, show_warning,
)
from tools.qt_operation_logs import OperationLogDialog


def _master_sheets(path: str):
    from tools.excel_metadata import list_workbook_sheets
    return list_workbook_sheets(path)


class MasterToTargetPage(AsyncPage):
    progress_changed = Signal(int, int)
    reverse = False

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(2, 2, 6, 2)
        outer.setSpacing(8)
        self.content = QWidget()
        layout = QVBoxLayout(self.content)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        inputs, input_layout = section("输入与输出")
        self.master_picker = PathPicker("Master 总表")
        self.target_picker = PathPicker("小表目录", choose_text="选择目录")
        self.output_picker = PathPicker("新输出目录（可选）", choose_text="选择位置")
        self.output_picker.line_edit.setReadOnly(False)
        self.output_picker.line_edit.setClearButtonEnabled(True)
        self.output_picker.line_edit.textChanged.connect(self.output_picker.line_edit.setToolTip)
        self.output_picker.line_edit.setPlaceholderText("留空：直接更新原 Master" if self.reverse else "留空：直接更新小表目录中的文件")
        self.master_picker.choose_button.clicked.connect(self.choose_master)
        self.target_picker.choose_button.clicked.connect(self.choose_targets)
        self.output_picker.choose_button.clicked.connect(self.choose_output)
        for picker in (self.master_picker, self.target_picker, self.output_picker):
            input_layout.addWidget(picker)
        layout.addWidget(inputs)

        mapping_box, mapping_layout = section("匹配与更新列")
        grid = QGridLayout()
        grid.setHorizontalSpacing(12)
        grid.setVerticalSpacing(6)
        for column, title in enumerate(("", "Key 列", "原文列", "译文列" if self.reverse else "内容开始列", "工作表", "表头行数")):
            grid.addWidget(QLabel(title), 0, column)
        self.master_columns = self._column_fields(("B", "C", "D"))
        self.target_columns = self._column_fields(("A", "B", "C"))
        self.master_sheet = QComboBox()
        self.master_sheet.setEditable(True)
        self.master_sheet.lineEdit().setPlaceholderText("活动工作表")
        self.target_sheet = QLineEdit()
        self.target_sheet.setPlaceholderText("留空：各文件活动表")
        self.master_headers, self.target_headers = QSpinBox(), QSpinBox()
        for widget in (self.master_headers, self.target_headers):
            widget.setRange(0, 1048575)
            widget.setValue(1)
            widget.setMaximumWidth(95)
        for row, title, columns, sheet, headers in (
            (1, "Master", self.master_columns, self.master_sheet, self.master_headers),
            (2, "小表", self.target_columns, self.target_sheet, self.target_headers),
        ):
            grid.addWidget(QLabel(title), row, 0)
            for col, widget in enumerate(columns, 1):
                grid.addWidget(widget, row, col)
            grid.addWidget(sheet, row, 4)
            grid.addWidget(headers, row, 5)
        grid.setColumnStretch(4, 1)
        mapping_layout.addLayout(grid)
        layout.addWidget(mapping_box)

        options, options_layout = section("写入选项")
        row = QHBoxLayout()
        self.column_count = QSpinBox()
        self.column_count.setRange(1, 16384)
        self.column_count.setMaximumWidth(95)
        self.fill_blank_only = QCheckBox("仅填空")
        self.fill_blank_only.setToolTip("开启后保留已有目标内容；多列时逐单元格判断。")
        self.allow_blank_write = QCheckBox("允许空白写入")
        self.allow_blank_write.setToolTip("开启后，来源空内容可清空目标内容。")
        if not self.reverse:
            row.addWidget(QLabel("连续更新列数"))
            row.addWidget(self.column_count)
        else:
            self.column_count.hide()
        row.addSpacing(15)
        row.addWidget(self.fill_blank_only)
        row.addSpacing(15)
        row.addWidget(self.allow_blank_write)
        row.addStretch(1)
        options_layout.addLayout(row)
        self.compatibility_resave = QCheckBox("同步后用 Excel 兼容性重存（需要 Windows / Excel）")
        if not self.reverse:
            options_layout.addWidget(self.compatibility_resave)
        else:
            self.compatibility_resave.hide()
        layout.addWidget(options)
        layout.addStretch(1)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setWidget(self.content)
        outer.addWidget(scroll, 1)
        self.result = QPlainTextEdit()
        self.result.setReadOnly(True)
        self.result.setPlaceholderText("完成后显示同步结果，处理详情可查看日志。")
        self.result.setFixedHeight(92)
        self.result.hide()
        outer.addWidget(self.result)
        self.status = muted_label()
        self.progress = QProgressBar()
        self.progress.setMaximumHeight(8)
        self.progress.setTextVisible(False)
        self.progress.hide()
        outer.addWidget(self.status)
        outer.addWidget(self.progress)
        self.progress_changed.connect(self._progress)

        self.run_button = primary_button("回填 Master" if self.reverse else "同步到小表")
        self.run_button.clicked.connect(self.run_sync)
        self.action_bar = QFrame()
        self.action_bar.setObjectName("pageActionBar")
        self.action_bar.setFixedHeight(46)
        actions = QHBoxLayout(self.action_bar)
        actions.setContentsMargins(2, 9, 6, 2)
        self.logs = OperationLogDialog(self)
        self.logs_button = QPushButton("查看日志")
        self.logs_button.clicked.connect(self.logs.open_logs)
        actions.addWidget(self.logs_button)
        actions.addStretch(1)
        self.run_button.setFixedWidth(136)
        actions.addWidget(self.run_button)
        outer.addWidget(self.action_bar)

    @staticmethod
    def _column_fields(defaults: tuple[str, ...]) -> tuple[QLineEdit, ...]:
        fields = tuple(QLineEdit(value) for value in defaults)
        for widget in fields:
            widget.setMaximumWidth(86)
            widget.setPlaceholderText("列字母")
        return fields

    def _set_busy(self, busy: bool, message: str = "") -> None:
        self.content.setEnabled(not busy)
        self.run_button.setEnabled(not busy)
        self.status.setText(message)
        self.progress.setVisible(busy)
        self.progress.setRange(0, 0 if busy else 1)

    def choose_master(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "选择 Master 总表", "", "Excel 文件 (*.xlsx *.xlsm)")
        if path:
            self.master_picker.set_path(path)
            self.master_sheet.clear()
            self._set_busy(True, "正在读取 Master 工作表…")
            self.run_in_background(_master_sheets, args=(path,), on_success=self._sheets_loaded,
                                   on_error=self._fail)

    def _sheets_loaded(self, choices) -> None:
        self.master_sheet.addItems(choices.sheet_names)
        self.master_sheet.setCurrentText(choices.default_sheet or "")
        self._set_busy(False)

    def choose_targets(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "选择小表目录", self.target_picker.path())
        if path:
            self.target_picker.set_path(path)

    def choose_output(self) -> None:
        parent = QFileDialog.getExistingDirectory(self, "选择输出位置（将在其中新建结果目录）")
        if parent:
            name = Path(self.target_picker.path()).name or "targets"
            suffix = "master_updated" if self.reverse else "synced"
            self.output_picker.set_path(str(Path(parent) / f"{name}_{suffix}"))

    def run_sync(self) -> None:
        if self.has_running_tasks():
            return
        if not all(picker.path() for picker in (self.master_picker, self.target_picker)):
            show_error(self, "缺少路径", "请选择 Master 和小表目录。")
            return
        from .master_to_target import ColumnMapping, sync_master_to_targets

        try:
            master_columns = ColumnMapping(*(widget.text() for widget in self.master_columns))
            target_columns = ColumnMapping(*(widget.text() for widget in self.target_columns))
            master_columns.indexes(self.column_count.value())
            target_columns.indexes(self.column_count.value())
        except ValueError as exc:
            show_error(self, "列配置错误", str(exc))
            return
        # Snapshot every widget value before dispatching to the worker.
        kwargs = {
            "master_file": self.master_picker.path(), "target_dir": self.target_picker.path(),
            "output_dir": self.output_picker.path() or None, "master_columns": master_columns,
            "target_columns": target_columns, "column_count": self.column_count.value(),
            "master_sheet": self.master_sheet.currentText().strip() or None,
            "target_sheet": self.target_sheet.text().strip() or None,
            "master_header_rows": self.master_headers.value(), "target_header_rows": self.target_headers.value(),
            "fill_blank_only": self.fill_blank_only.isChecked(),
            "allow_blank_write": self.allow_blank_write.isChecked(),
            "progress_callback": self.progress_changed.emit,
            "log_callback": self.logs.buffer.append,
        }
        operation = sync_master_to_targets
        if self.reverse:
            from .target_to_master import sync_targets_to_master
            operation = sync_targets_to_master
            kwargs.pop("column_count")
        else:
            kwargs["compatibility_resave"] = self.compatibility_resave.isChecked()
        self.result.clear()
        self.result.hide()
        self.logs.start_run()
        self._set_busy(True, "正在读取小表并回填 Master…" if self.reverse else "正在读取 Master 并同步…")
        self.run_in_background(operation, kwargs=kwargs,
                               on_success=self._finish, on_error=self._fail)

    def _progress(self, completed: int, total: int) -> None:
        self.progress.setRange(0, max(total, 1))
        self.progress.setValue(completed)
        self.status.setText(f"已处理小表 {completed} / {total}")

    def _finish(self, summary) -> None:
        from .master_to_target import format_summary

        has_issues = summary.failed_files or summary.warnings
        self._set_busy(False, "同步完成（请查看日志）" if has_issues else "同步完成")
        text = (summary.describe() + f"\n实际更新单元格: {summary.details['updated_cells']}"
                if self.reverse else format_summary(summary))
        self.result.setPlainText(text)
        self.result.show()
        self.logs.finish_run()
        if has_issues:
            show_warning(self, "同步完成（请查看日志）", text)

    def _fail(self, message: str) -> None:
        self.logs.finish_run(f"[失败] {message}")
        self._set_busy(False, "处理失败")
        self.result.setPlainText(message)
        self.result.show()
        show_error(self, "处理失败", message)


class TargetToMasterPage(MasterToTargetPage):
    reverse = True


class ContentSyncPage(AsyncPage):
    """Keep both directions alive when switching tabs, including their workers."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(2, 2, 6, 2)
        self.tabs = QTabWidget()
        self.master_to_target_page = MasterToTargetPage()
        self.target_to_master_page = TargetToMasterPage()
        self.tabs.addTab(self.master_to_target_page, "Master → 小表")
        self.tabs.addTab(self.target_to_master_page, "小表 → Master")
        layout.addWidget(self.tabs)

    def has_running_tasks(self) -> bool:
        return (super().has_running_tasks()
                or self.master_to_target_page.has_running_tasks()
                or self.target_to_master_page.has_running_tasks())
