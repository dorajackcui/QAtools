"""Filename-list preview and extraction in the shared Qt shell."""
from pathlib import Path

from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt, QUrl, Signal
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QCheckBox, QFileDialog, QFrame, QHBoxLayout, QHeaderView, QLabel, QPlainTextEdit,
    QProgressBar, QPushButton, QTableView, QVBoxLayout, QWidget,
)

from tools.qt_gui_common import AsyncPage, PathPicker, muted_label, primary_button, section, show_error
from tools.qt_operation_logs import OperationLogDialog


class PreviewModel(QAbstractTableModel):
    headers = ("清单项", "来源相对路径", "输出相对路径", "状态")

    def __init__(self, parent=None):
        super().__init__(parent)
        self.plan = None
        self.results = {}

    def set_plan(self, plan, summary=None):
        self.beginResetModel()
        self.plan = plan
        self.results = {item.source: item for item in summary.results} if summary else {}
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() or self.plan is None else len(self.plan.entries) + len(self.plan.not_found)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.headers)

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole and orientation == Qt.Orientation.Horizontal:
            return self.headers[section]
        return None

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid() or self.plan is None:
            return None
        if role not in (Qt.ItemDataRole.DisplayRole, Qt.ItemDataRole.ToolTipRole):
            return None
        if index.row() >= len(self.plan.entries):
            name = self.plan.not_found[index.row() - len(self.plan.entries)]
            return (name, "—", "—", "未找到")[index.column()]
        entry = self.plan.entries[index.row()]
        result = self.results.get(str(entry.source))
        status = result.status if result else entry.status
        labels = {"ready": "可复制", "conflict": "冲突", "unavailable": "不可读",
                  "copied": "已复制", "failed": "失败", "skipped": "跳过"}
        message = result.message if result else entry.message
        if role == Qt.ItemDataRole.ToolTipRole and index.column() == 3 and message:
            return message
        return (" / ".join(entry.names), str(entry.source), str(entry.destination), labels[status])[index.column()]


class FileCollectorPage(AsyncPage):
    progress_changed = Signal(int, int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.plan = None
        self.summary = None
        self._generation = 0
        self._running = False
        self._suggested_output = ""
        outer = QVBoxLayout(self)
        outer.setContentsMargins(2, 2, 6, 2)
        outer.setSpacing(8)
        self.inputs, inputs = section("输入与输出")
        self.source_picker = PathPicker("来源目录", choose_text="选择目录")
        self.output_picker = PathPicker("新输出目录", choose_text="选择位置")
        for picker in (self.source_picker, self.output_picker):
            picker.line_edit.setReadOnly(False)
            picker.line_edit.setClearButtonEnabled(True)
            picker.line_edit.textChanged.connect(self._invalidate)
            picker.line_edit.textChanged.connect(picker.line_edit.setToolTip)
        self.output_picker.line_edit.setPlaceholderText("必填：尚不存在的新目录")
        inputs.addWidget(self.source_picker)
        inputs.addWidget(self.output_picker)
        names_row = QHBoxLayout()
        names_row.addWidget(QLabel("文件名列表"))
        names_row.addStretch(1)
        self.import_button = QPushButton("导入 TXT")
        names_row.addWidget(self.import_button)
        inputs.addLayout(names_row)
        self.names = QPlainTextEdit()
        self.names.setPlaceholderText("每行一个文件名，可从 Excel 单列粘贴，例如：\nDialogue.xlsx\nUI 文本")
        self.names.setFixedHeight(104)
        self.names.textChanged.connect(self._invalidate)
        inputs.addWidget(self.names)
        options = QHBoxLayout()
        self.preserve_tree = QCheckBox("保留来源目录结构")
        self.comma_separated = QCheckBox("同时按中英文逗号分隔")
        options.addWidget(self.preserve_tree)
        options.addWidget(self.comma_separated)
        options.addStretch(1)
        inputs.addLayout(options)
        outer.addWidget(self.inputs)
        for option in (self.preserve_tree, self.comma_separated):
            option.toggled.connect(self._invalidate)

        self.counts = muted_label("尚未预览", word_wrap=True)
        outer.addWidget(self.counts)
        self.table = QTableView()
        self.model = PreviewModel(self.table)
        self.table.setModel(self.model)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        self.table.setWordWrap(False)
        self.table.verticalHeader().hide()
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Fixed)
        self.table.setColumnWidth(3, 88)
        self.table.setMinimumHeight(90)
        outer.addWidget(self.table, 1)
        self.status = muted_label("", word_wrap=True)
        outer.addWidget(self.status)
        self.progress = QProgressBar()
        self.progress.setMaximumHeight(8)
        self.progress.setTextVisible(False)
        self.progress.hide()
        self.progress_changed.connect(self._progress)
        outer.addWidget(self.progress)
        self.action_bar = QFrame()
        self.action_bar.setObjectName("pageActionBar")
        self.action_bar.setFixedHeight(46)
        actions = QHBoxLayout(self.action_bar)
        actions.setContentsMargins(2, 9, 6, 2)
        self.logs = OperationLogDialog(self)
        self.logs_button = QPushButton("查看日志")
        self.logs_button.clicked.connect(self.logs.open_logs)
        self.open_button = QPushButton("打开输出目录")
        self.open_button.setEnabled(False)
        self.open_button.clicked.connect(self._open_output)
        self.preview_button = QPushButton("预览匹配")
        self.run_button = primary_button("复制可用文件")
        self.run_button.setFixedWidth(136)
        self.run_button.setEnabled(False)
        actions.addWidget(self.logs_button)
        actions.addWidget(self.open_button)
        actions.addStretch(1)
        actions.addWidget(self.preview_button)
        actions.addWidget(self.run_button)
        outer.addWidget(self.action_bar)
        self.source_picker.choose_button.clicked.connect(self.choose_source)
        self.output_picker.choose_button.clicked.connect(self.choose_output)
        self.import_button.clicked.connect(self.import_names)
        self.preview_button.clicked.connect(self.preview)
        self.run_button.clicked.connect(self.copy_files)

    def _invalidate(self, *_):
        self._generation += 1
        self.plan = None
        self.summary = None
        self.model.set_plan(None)
        self.counts.setText("输入已变化，请预览匹配")
        self.status.clear()
        self.run_button.setEnabled(False)
        self.open_button.setEnabled(False)

    def choose_source(self):
        path = QFileDialog.getExistingDirectory(self, "选择来源目录", self.source_picker.path())
        if path:
            suggest = not self.output_picker.path() or self.output_picker.path() == self._suggested_output
            self.source_picker.set_path(path)
            if suggest:
                source = Path(path)
                self._suggested_output = str(source.parent / f"{source.name}_selected")
                self.output_picker.set_path(self._suggested_output)

    def choose_output(self):
        path = QFileDialog.getExistingDirectory(self, "选择输出位置（将在其中新建结果目录）")
        if path:
            name = Path(self.source_picker.path()).name or "files"
            self.output_picker.set_path(str(Path(path) / f"{name}_selected"))

    def _busy(self, busy):
        self._running = busy
        self.inputs.setEnabled(not busy)
        self.preview_button.setEnabled(not busy)
        self.run_button.setEnabled(not busy and self.plan is not None and self.plan.ready_count > 0)
        self.open_button.setEnabled(not busy and self.summary is not None and self.summary.output_dir.is_dir())
        self.progress.setVisible(busy)
        if busy:
            self.progress.setRange(0, 0)

    def import_names(self):
        from .collector import read_names_file
        path, _ = QFileDialog.getOpenFileName(self, "导入 UTF-8 文件名清单", "", "文本文件 (*.txt);;所有文件 (*)")
        if not path:
            return
        self._busy(True)
        self.status.setText("正在读取文件名清单…")

        def loaded(text):
            self.names.setPlainText(text)
            self._busy(False)
            self.status.setText("清单已导入，请预览匹配。")

        self.run_in_background(read_names_file, args=(path,), on_success=loaded, on_error=self._failed)

    def preview(self):
        if self._running:
            return
        from .collector import build_copy_plan
        self._invalidate()
        generation = self._generation
        self.logs.start_run()
        self._busy(True)
        self.status.setText("正在扫描文件并检查输出路径…")

        def ready(plan):
            if generation == self._generation:
                self.plan = plan
                self.model.set_plan(plan)
                self.counts.setText(plan.describe())
                self.status.setText("预览完成，核对后复制可用文件。" if plan.ready_count else "没有可复制文件，请检查清单或冲突。")
            self.logs.finish_run()
            self._busy(False)

        self.run_in_background(build_copy_plan, args=(self.source_picker.path(), self.names.toPlainText()),
            kwargs={"output_dir": self.output_picker.path(), "preserve_tree": self.preserve_tree.isChecked(),
                    "comma_separated": self.comma_separated.isChecked(), "log_callback": self.logs.buffer.append},
            on_success=ready, on_error=self._failed)

    def copy_files(self):
        if self._running or self.plan is None or not self.plan.ready_count:
            return
        from .collector import execute_copy_plan
        plan, generation = self.plan, self._generation
        self.plan = None  # A copy plan is single-use, including after a failure.
        self.logs.start_run()
        self._busy(True)
        self.status.setText("正在复核预览并复制文件…")

        def copied(summary):
            if generation == self._generation:
                self.summary = summary
                self.model.set_plan(plan, summary)
                self.status.setText(f"复制 {summary.succeeded_files}；失败 {summary.failed_files}；"
                                    f"跳过 {summary.skipped_files}；未找到 {len(plan.not_found)} 项")
            self.logs.finish_run()
            self._busy(False)

        self.run_in_background(execute_copy_plan, args=(plan,), kwargs={
            "log_callback": self.logs.buffer.append, "progress_callback": self.progress_changed.emit},
            on_success=copied, on_error=self._failed)

    def _failed(self, message):
        self.plan = None
        self.logs.finish_run(f"失败：{message}")
        self.status.setText(message)
        self._busy(False)
        show_error(self, "文件提取失败", message)

    def _progress(self, current, total):
        self.progress.setRange(0, max(1, total))
        self.progress.setValue(current)

    def _open_output(self):
        if self.summary:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.summary.output_dir)))
