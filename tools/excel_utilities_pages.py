"""Lightweight Qt entry points for the four Excel utilities."""

from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QComboBox, QFileDialog, QFormLayout, QFrame, QHBoxLayout, QLineEdit,
    QPlainTextEdit, QProgressBar, QPushButton, QScrollArea, QSpinBox, QVBoxLayout, QWidget,
)

from tools.qt_gui_common import AsyncPage, PathPicker, muted_label, primary_button, section, show_error, show_warning
from tools.qt_operation_logs import OperationLogDialog


class UtilityPage(AsyncPage):
    progress_changed = Signal(int, int)
    suffix = "processed"
    button_text = "开始处理"
    input_title = "文件目录"
    output_hint = "留空：直接更新文件目录中的原文件"

    def __init__(self, parent=None):
        super().__init__(parent)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(2, 2, 6, 2)
        outer.setSpacing(8)
        self.content = QWidget()
        layout = QVBoxLayout(self.content)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        inputs, input_layout = section("输入与输出")
        self.input_layout = input_layout
        self.add_source_picker()
        self.input_picker = self.make_picker(self.input_title, "选择目录")
        self.output_picker = self.make_picker("新输出目录（可选）", "选择位置")
        self.output_picker.line_edit.setReadOnly(False)
        self.output_picker.line_edit.setClearButtonEnabled(True)
        self.output_picker.line_edit.textChanged.connect(self.output_picker.line_edit.setToolTip)
        self.output_picker.line_edit.setPlaceholderText(self.output_hint)
        self.input_picker.choose_button.clicked.connect(lambda: self.choose_folder(self.input_picker))
        self.output_picker.choose_button.clicked.connect(self.choose_output)
        layout.addWidget(inputs)
        options, options_layout = section("处理选项")
        self.form = QFormLayout()
        self.form.setHorizontalSpacing(16)
        self.form.setVerticalSpacing(10)
        options_layout.addLayout(self.form)
        self.add_options()
        if self.form.rowCount():
            layout.addWidget(options)
        else:
            options.deleteLater()
        layout.addStretch(1)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setWidget(self.content)
        outer.addWidget(scroll, 1)
        self.result = QPlainTextEdit()
        self.result.setReadOnly(True)
        self.result.setFixedHeight(110)
        self.result.hide()
        outer.addWidget(self.result)
        self.status = muted_label()
        self.progress = QProgressBar()
        self.progress.setMaximumHeight(8)
        self.progress.setTextVisible(False)
        self.progress.hide()
        self.progress_changed.connect(self._progress)
        outer.addWidget(self.status)
        outer.addWidget(self.progress)
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
        self.run_button = primary_button(self.button_text)
        self.run_button.setFixedWidth(136)
        self.run_button.clicked.connect(self.run_operation)
        actions.addWidget(self.run_button)
        outer.addWidget(self.action_bar)

    def make_picker(self, title, button):
        picker = PathPicker(title, choose_text=button)
        self.input_layout.addWidget(picker)
        return picker

    def add_source_picker(self):
        pass

    def add_options(self):
        pass

    def sheet_options(self):
        self.sheet = QLineEdit()
        self.sheet.setPlaceholderText("留空：各文件活动工作表")
        self.form.addRow("工作表", self.sheet)
        self.header_rows = QSpinBox()
        self.header_rows.setRange(0, 1048575)
        self.header_rows.setValue(1)
        self.header_rows.setMaximumWidth(110)
        self.form.addRow("表头行数", self.header_rows)

    def choose_folder(self, picker):
        path = QFileDialog.getExistingDirectory(self, "选择目录", picker.path())
        if path:
            picker.set_path(path)

    def choose_output(self):
        parent = QFileDialog.getExistingDirectory(self, "选择输出位置（将在其中新建结果目录）")
        if parent:
            name = Path(self.input_picker.path()).name or "excel"
            self.output_picker.set_path(str(Path(parent) / f"{name}_{self.suffix}"))

    def _busy(self, busy, message=""):
        self.content.setEnabled(not busy)
        self.run_button.setEnabled(not busy)
        self.status.setText(message)
        self.progress.setVisible(busy)
        self.progress.setRange(0, 0 if busy else 1)

    def run_operation(self):
        if self.has_running_tasks():
            return
        if not self.input_picker.path():
            show_error(self, "缺少路径", "请选择文件目录。")
            return
        try:
            operation, kwargs = self.operation()
        except ValueError as exc:
            show_error(self, "配置错误", str(exc))
            return
        kwargs.update(output_dir=self.output_picker.path() or None, progress_callback=self.progress_changed.emit,
                      log_callback=self.logs.buffer.append)
        self.logs.start_run()
        self.result.hide()
        self._busy(True, "正在处理…")
        self.run_in_background(operation, kwargs=kwargs, on_success=self._finish, on_error=self._fail)

    def _progress(self, done, total):
        self.progress.setRange(0, max(1, total))
        self.progress.setValue(done)
        self.status.setText(f"已处理 {done} / {total}")

    def _finish(self, summary):
        text = summary.describe()
        if "excel_report" in summary.details:
            text += f"\n统计表: {summary.details['excel_report']}"
        if "replaced_files" in summary.details:
            text += f"\n替换文件: {summary.details['replaced_files']}"
        self._busy(False, "处理完成（请查看日志）" if summary.failed_files or summary.warnings else "处理完成")
        self.result.setPlainText(text)
        self.result.show()
        self.logs.finish_run()
        if summary.failed_files or summary.warnings:
            show_warning(self, "处理完成（请查看日志）", text)

    def _fail(self, message):
        self.logs.finish_run(f"[失败] {message}")
        self._busy(False, "处理失败")
        self.result.setPlainText(message)
        self.result.show()
        show_error(self, "处理失败", message)


class CompatibilityPage(UtilityPage):
    suffix = "compatible"
    button_text = "兼容性重存"

    def operation(self):
        from tools.excel_compatibility.processor import resave_workbooks
        return resave_workbooks, {"folder_path": self.input_picker.path()}


class ColumnToolsPage(UtilityPage):
    suffix = "columns"
    button_text = "执行列操作"

    def add_options(self):
        self.action = QComboBox()
        for label, value in (("清空列内容（保留格式）", "clear"), ("插入一列", "insert"), ("删除整列", "delete")):
            self.action.addItem(label, value)
        self.form.addRow("动作", self.action)
        self.column = QLineEdit("C")
        self.column.setMaximumWidth(110)
        self.form.addRow("列字母", self.column)
        self.sheet_options()
        self.inserted_header = QLineEdit("Translation")
        self.form.addRow("插入列的首行标题", self.inserted_header)
        self.action.currentIndexChanged.connect(self._action_changed)
        self._action_changed()

    def _action_changed(self):
        self.header_rows.setEnabled(self.action.currentData() == "clear")
        self.inserted_header.setEnabled(self.action.currentData() == "insert")

    def operation(self):
        from openpyxl.utils.cell import column_index_from_string
        from tools.column_tools.processor import operate_columns
        column = self.column.text().strip().upper()
        if not column or column_index_from_string(column) > 16384:
            raise ValueError("列必须是 A–XFD 范围内的字母。")
        return operate_columns, dict(folder_path=self.input_picker.path(), action=self.action.currentData(),
                                     column=column, sheet=self.sheet.text().strip() or None,
                                     header_rows=self.header_rows.value(), inserted_header=self.inserted_header.text())


class DeepReplacePage(UtilityPage):
    input_title = "目标文件目录"
    suffix = "replaced"
    button_text = "替换同名文件"

    def add_source_picker(self):
        self.source_picker = self.make_picker("替换来源目录", "选择目录")
        self.source_picker.choose_button.clicked.connect(lambda: self.choose_folder(self.source_picker))


    def operation(self):
        from tools.deep_replace.replacer import replace_files
        if not self.source_picker.path():
            raise ValueError("请选择替换来源目录。")
        return replace_files, dict(source_dir=self.source_picker.path(), target_dir=self.input_picker.path())


class UntranslatedStatsPage(UtilityPage):
    suffix = "stats"
    button_text = "生成统计表"
    output_hint = "留空：统计表保存到所选文件目录"

    def add_options(self):
        self.source_column = QLineEdit("B")
        self.target_column = QLineEdit("C")
        for widget in (self.source_column, self.target_column):
            widget.setMaximumWidth(110)
        self.form.addRow("原文列", self.source_column)
        self.form.addRow("译文列", self.target_column)
        self.mode = QComboBox()
        self.mode.addItem("中文字符数", "chinese_chars")
        self.mode.addItem("英文词数", "english_words")
        self.form.addRow("计数口径", self.mode)
        self.sheet_options()

    def operation(self):
        from tools.untranslated_stats.stats import untranslated_stats
        return untranslated_stats, dict(target_dir=self.input_picker.path(), source_column=self.source_column.text(),
                                        target_column=self.target_column.text(), mode=self.mode.currentData(),
                                        sheet=self.sheet.text().strip() or None, header_rows=self.header_rows.value())
