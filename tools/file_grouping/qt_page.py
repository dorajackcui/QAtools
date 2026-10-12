"""Minimal file-group scan and export page in the shared Qt shell."""
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt, QUrl, Signal
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QComboBox, QDialog, QDialogButtonBox, QFileDialog, QFormLayout, QHBoxLayout, QHeaderView, QLabel, QLineEdit,
    QPlainTextEdit, QProgressBar, QPushButton, QTableView, QVBoxLayout,
)

from tools.qt_gui_common import AsyncPage, PathPicker, muted_label, primary_button, section, show_error
from tools.qt_page_helpers import EXCEL_FILTER, _add_action_bar


class GroupModel(QAbstractTableModel):
    headers = ("分组", "文件数", "去重中文字数")

    def __init__(self, parent=None):
        super().__init__(parent)
        self.result = None

    def set_result(self, result):
        self.beginResetModel()
        self.result = result
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() or self.result is None else len(self.result.groups)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.headers)

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if orientation == Qt.Orientation.Horizontal:
            if role == Qt.ItemDataRole.DisplayRole:
                return self.headers[section]
            if role == Qt.ItemDataRole.TextAlignmentRole:
                return (Qt.AlignmentFlag.AlignLeft if section == 0 else Qt.AlignmentFlag.AlignRight) | Qt.AlignmentFlag.AlignVCenter
        return None

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid() or self.result is None:
            return None
        group = self.result.groups[index.row()]
        if role == Qt.ItemDataRole.DisplayRole:
            return (group.name, f"{len(group.files):,}", f"{group.unique_chars:,}")[index.column()]
        if role == Qt.ItemDataRole.TextAlignmentRole and index.column() > 0:
            return Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        if role == Qt.ItemDataRole.ToolTipRole:
            return "点击查看文件；各文件可分别分配。" if group.name == "其他" else "点击查看文件；同群文件建议一起分配。"
        return None


class FileGroupingPage(AsyncPage):
    progress_changed = Signal(int, int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.result = None
        self.exported = None
        self._running = False
        self._generation = 0
        outer = QVBoxLayout(self)
        outer.setContentsMargins(2, 2, 6, 2)
        outer.setSpacing(8)
        self.inputs, inputs = section()
        kind_row = QHBoxLayout()
        kind_label = QLabel("输入方式")
        kind_label.setMinimumWidth(112)
        self.input_kind = QComboBox()
        self.input_kind.addItem("文件夹", "folder")
        self.input_kind.addItem("合并表", "merged")
        kind_row.addWidget(kind_label)
        kind_row.addWidget(self.input_kind, 1)
        inputs.addLayout(kind_row)
        self.source_picker = PathPicker("来源", choose_text="选择目录")
        self.source_picker.line_edit.setReadOnly(False)
        self.source_picker.line_edit.setClearButtonEnabled(True)
        self.source_picker.line_edit.setPlaceholderText("递归扫描 .xlsx / .xlsm")
        self.source_picker.choose_button.clicked.connect(self.choose_source)
        self.source_picker.line_edit.textChanged.connect(self.source_picker.line_edit.setToolTip)
        inputs.addWidget(self.source_picker)
        form = QFormLayout()
        form.setHorizontalSpacing(8)
        self.input_form = form
        self.filename_column = QLineEdit("A")
        self.filename_column.setMaximumWidth(120)
        form.addRow("文件名列", self.filename_column)
        form.setRowVisible(self.filename_column, False)
        self.source_header = QLineEdit("MsgStr")
        self.target_header = QLineEdit("fr")
        form.addRow("原文表头", self.source_header)
        form.addRow("译文表头", self.target_header)
        for field in (self.filename_column, self.source_header, self.target_header):
            form.labelForField(field).setMinimumWidth(112)
        inputs.addLayout(form)
        outer.addWidget(self.inputs)
        for field in (self.source_picker.line_edit, self.filename_column, self.source_header, self.target_header):
            field.textChanged.connect(self._invalidate)
        self.input_kind.currentIndexChanged.connect(self._kind_changed)

        self.counts = muted_label("尚未扫描")
        outer.addWidget(self.counts)
        self.table = QTableView()
        self.model = GroupModel(self.table)
        self.table.setModel(self.model)
        self.table.setAccessibleName("文件分组；点击或按回车查看组内文件")
        self.table.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QTableView.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QTableView.EditTrigger.NoEditTriggers)
        self.table.setShowGrid(False)
        self.table.setWordWrap(False)
        self.table.verticalHeader().hide()
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.clicked.connect(self.show_group)
        self.table.activated.connect(self.show_group)
        outer.addWidget(self.table, 1)
        self.status = muted_label("", word_wrap=True)
        self.status.setTextFormat(Qt.TextFormat.RichText)
        self.status.linkActivated.connect(self._open_status)
        outer.addWidget(self.status)
        self.progress = QProgressBar()
        self.progress.setMaximumHeight(8)
        self.progress.setTextVisible(False)
        self.progress.hide()
        self.progress_changed.connect(self._progress)
        outer.addWidget(self.progress)
        self.scan_button = primary_button("扫描分组")
        self.export_button = QPushButton("导出分组…")
        self.export_button.setEnabled(False)
        self.scan_button.clicked.connect(self.scan)
        self.export_button.clicked.connect(self.export_groups)
        self.action_bar = _add_action_bar(outer, self.scan_button, self.export_button)
        self.details_dialog = QDialog(self)
        self.details_dialog.setModal(True)
        self.details_dialog.resize(620, 420)
        details = QVBoxLayout(self.details_dialog)
        self.details_note = QLabel()
        self.details_note.setWordWrap(True)
        details.addWidget(self.details_note)
        self.details_text = QPlainTextEdit()
        self.details_text.setReadOnly(True)
        details.addWidget(self.details_text, 1)
        close = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        close_button = close.button(QDialogButtonBox.StandardButton.Close)
        close_button.setText("关闭")
        close_button.setAutoDefault(False)
        close.rejected.connect(self.details_dialog.reject)
        details.addWidget(close)

    def _invalidate(self, *_):
        self._generation += 1
        self.result = self.exported = None
        self.model.set_result(None)
        self.counts.setText("输入已变化，请扫描分组")
        self.status.clear()
        self.status.setToolTip("")
        self.export_button.setEnabled(False)
        self.details_dialog.reject()

    def choose_source(self):
        if self.input_kind.currentData() == "merged":
            path, _ = QFileDialog.getOpenFileName(self, "选择合并表", self.source_picker.path(), EXCEL_FILTER)
        else:
            path = QFileDialog.getExistingDirectory(self, "选择来源目录", self.source_picker.path())
        if path:
            self.source_picker.set_path(path)

    def _kind_changed(self, *_):
        merged = self.input_kind.currentData() == "merged"
        self.input_form.setRowVisible(self.filename_column, merged)
        self.source_picker.choose_button.setText("选择表格" if merged else "选择目录")
        self.source_picker.line_edit.setPlaceholderText("选择 .xlsx / .xlsm 合并表" if merged else "递归扫描 .xlsx / .xlsm")
        self.source_picker.clear()
        self._invalidate()

    def _busy(self, busy):
        self._running = busy
        self.inputs.setEnabled(not busy)
        self.table.setEnabled(not busy)
        self.scan_button.setEnabled(not busy)
        self.export_button.setEnabled(not busy and self.result is not None and self.result.can_export)
        self.progress.setVisible(busy)
        if busy:
            self.progress.setRange(0, 0)

    def scan(self):
        if self._running:
            return
        kind = self.input_kind.currentData()
        if kind == "merged":
            from .merged import scan_merged_table as scan_function
        else:
            from .grouping import scan_file_groups as scan_function
        self._invalidate()
        generation = self._generation
        self._busy(True)
        self.status.setText("正在扫描待翻译原文…")

        def ready(result):
            if generation == self._generation:
                self.result = result
                self.model.set_result(result)
                self.counts.setText(f"{len(result.files)} 个文件，{len(result.groups)} 个分组")
                self.status.setText(f'扫描不完整，{len(result.issues)} 处问题。<a href="issues">查看问题</a>'
                                    if result.issues else "扫描完成，点击组名查看文件。")
            self._busy(False)

        options = {"source_header": self.source_header.text(), "target_header": self.target_header.text(),
                   "progress_callback": self.progress_changed.emit}
        if kind == "merged":
            options["filename_column"] = self.filename_column.text()
        self.run_in_background(scan_function, args=(self.source_picker.path(),), kwargs=options,
                               on_success=ready, on_error=self._failed)

    def show_group(self, index):
        if self.result is None or not index.isValid():
            return
        group = self.result.groups[index.row()]
        self.details_dialog.setWindowTitle(f"{group.name}（{len(group.files)} 个文件）")
        self.details_note.setText("各文件独立，可以分别分配。" if group.name == "其他" else "同群文件存在直接或间接的重复原文，建议一起分配。")
        self.details_text.setPlainText("\n".join(path.as_posix() for path in group.files))
        self.details_dialog.open()

    def _open_status(self, link):
        if link == "issues" and self.result:
            self.details_dialog.setWindowTitle("合并表问题" if self.input_kind.currentData() == "merged" else "未扫描文件")
            self.details_note.setText("请修正问题后重新扫描，再导出分组。")
            self.details_text.setPlainText("\n\n".join(f"{issue.path.as_posix()}\n{issue.message}" for issue in self.result.issues))
            self.details_dialog.open()
        elif link == "output" and self.exported:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.exported.output_dir)))

    def export_groups(self):
        if self._running or self.result is None or not self.result.can_export:
            return
        kind = self.input_kind.currentData()
        if kind == "merged":
            from .merged import export_merged_groups as export_function
            initial_dir, name = self.result.input_file.parent, self.result.input_file.stem
        else:
            from .grouping import export_file_groups as export_function
            initial_dir, name = self.result.source_dir.parent, self.result.source_dir.name
        parent = QFileDialog.getExistingDirectory(self, "选择保存位置（将在其中新建分组目录）", str(initial_dir))
        if not parent:
            return
        output = Path(parent) / f"{name}_文件分组_{datetime.now():%Y%m%d_%H%M%S_%f}"
        result, generation = self.result, self._generation
        self._busy(True)
        self.status.setText("正在复核并导出分组…")

        def exported(summary):
            if generation == self._generation:
                self.exported = summary
                unit = "工作簿" if kind == "merged" else "文件"
                self.status.setText(f'已导出 {summary.file_count} 个{unit}。<a href="output">打开输出目录</a>')
                self.status.setToolTip(str(summary.output_dir))
            self._busy(False)

        self.run_in_background(export_function, args=(result, output), kwargs={
            "progress_callback": self.progress_changed.emit}, on_success=exported, on_error=self._failed)

    def _failed(self, message):
        from html import escape
        self.result = self.exported = None
        self.model.set_result(None)
        self.counts.setText("请重新扫描")
        self.status.setText(escape(message))
        self._busy(False)
        show_error(self, "文件分组失败", message)

    def _progress(self, current, total):
        self.progress.setRange(0, max(1, total))
        self.progress.setValue(current)
