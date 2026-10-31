"""Lightweight previews with column selection shared by Qt input forms."""

from dataclasses import dataclass
import weakref

from PySide6.QtCore import QAbstractTableModel, QModelIndex, QObject, Qt
from PySide6.QtWidgets import (
    QComboBox, QDialog, QDialogButtonBox, QFormLayout, QHBoxLayout, QLabel,
    QTableView, QVBoxLayout,
)

from tools.qt_gui_common import muted_label


@dataclass(frozen=True)
class ColumnChoice:
    label: str
    field: object
    kind: str = "letter"
    optional: bool = False
    aliases: tuple[str, ...] = ()
    alias_role: str = ""


def _letter(index):
    result = ""
    while index:
        index, digit = divmod(index - 1, 26)
        result = chr(65 + digit) + result
    return result


class PreviewModel(QAbstractTableModel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.preview = None

    def set_preview(self, preview):
        self.beginResetModel()
        self.preview = preview
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()):
        return 6 if self.preview is not None and not parent.isValid() else 0

    def columnCount(self, parent=QModelIndex()):
        return self.preview.columns if self.preview is not None and not parent.isValid() else 0

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if index.isValid() and self.preview is not None and role in (
            Qt.ItemDataRole.DisplayRole, Qt.ItemDataRole.ToolTipRole,
        ):
            return self.preview.cells.get((index.row(), index.column()), "")
        return None

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal:
                return f"{_letter(section + 1)} ({section + 1})"
            return str(section + 1)
        return None


class ExcelColumnSelector(QObject):
    """Choosers open this explicitly; restoring stored paths never performs I/O."""

    def __init__(self, owner, picker, sheet=None, *, title="预览并选择列", folder=False,
                 first_sheet=False, browse_sheets=False, exclusions=None, exclude_names=(), skip_links=False,
                 choices=(), header_alias_store=None):
        super().__init__(owner)
        self.title = title
        # Qt owns the selector; avoid a Python owner -> selector -> owner cycle.
        self.owner, self.picker, self.sheet = weakref.proxy(owner), picker, sheet
        self.folder, self.first_sheet = folder, first_sheet
        self.browse_sheets, self.exclusions = browse_sheets, exclusions
        self.exclude_names, self.skip_links = exclude_names, skip_links
        self.choices, self.header_alias_store = choices, header_alias_store
        self._generation = 0
        self._loading = False
        self.dialog = None
        self._pending_open = False
        picker.line_edit.textChanged.connect(self.invalidate)
        if sheet is not None:
            signal = sheet.currentTextChanged if isinstance(sheet, QComboBox) else sheet.textChanged
            signal.connect(self.invalidate)

    def invalidate(self, *_):
        self._generation += 1
        if self.dialog is not None:
            self.dialog.hide()
            self.model.set_preview(None)

    def _build_dialog(self):
        self.dialog = QDialog(self.picker)
        self.dialog.setWindowTitle(self.title)
        self.dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.dialog.resize(900, 460 if self.choices else 330)
        layout = QVBoxLayout(self.dialog)
        self.status = muted_label("", word_wrap=True)
        self.status.setTextFormat(Qt.TextFormat.PlainText)
        self.status.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        layout.addWidget(self.status)
        self.sheet_choice = QComboBox()
        if not self.first_sheet and (self.browse_sheets or self.sheet is not None):
            row = QHBoxLayout()
            row.addWidget(QLabel("预览工作表"))
            row.addWidget(self.sheet_choice, 1)
            layout.addLayout(row)
            self.sheet_choice.currentIndexChanged.connect(self._sheet_changed)
        else:
            self.sheet_choice.hide()
        self.table = QTableView()
        self.model = PreviewModel(self.table)
        self.table.setModel(self.model)
        self.table.setEditTriggers(QTableView.EditTrigger.NoEditTriggers)
        self.table.setWordWrap(False)
        self.table.horizontalHeader().setDefaultSectionSize(160)
        self.table.verticalHeader().setDefaultSectionSize(30)
        layout.addWidget(self.table, 1)
        self.column_choices = []
        self.choice_form = form = QFormLayout()
        for choice in self.choices:
            combo = QComboBox()
            combo.setMinimumWidth(260)
            form.addRow(choice.label, combo)
            self.column_choices.append(combo)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Cancel |
                                  QDialogButtonBox.StandardButton.Ok)
        self.apply_button = buttons.button(QDialogButtonBox.StandardButton.Ok)
        self.apply_button.setText("使用所选列")
        self.apply_button.setProperty("primary", True)
        self.apply_button.setEnabled(False)
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("取消")
        buttons.accepted.connect(self._apply_columns)
        buttons.rejected.connect(self.dialog.reject)
        layout.addWidget(buttons)

    def open_preview(self):
        if self.dialog is None:
            self._build_dialog()
        for choice, combo in zip(self.choices, self.column_choices):
            self.choice_form.setRowVisible(combo, not choice.field.isHidden())
        self.dialog.show()
        self.dialog.raise_()
        if self._loading:
            self._generation += 1
            self._pending_open = True
            return
        self._load()

    def _sheet_changed(self, _index):
        self._load(self.sheet_choice.currentData(), use_sample=True)

    def _load(self, preview_sheet=None, *, use_sample=False):
        if self._loading:
            return
        from tools.excel_preview import read_excel_preview

        path = self.picker.path()
        if not path:
            self.status.setText("请先选择文件或目录。")
            return
        self._generation += 1
        generation = self._generation
        self._loading = True
        self.sheet_choice.setEnabled(False)
        self.apply_button.setEnabled(False)
        for combo in self.column_choices:
            combo.setEnabled(False)
        self.model.set_preview(None)
        self.status.setText("正在读取表头及前 5 行…")
        sheet = self.sheet.currentText() if isinstance(self.sheet, QComboBox) else self.sheet.text() if self.sheet else ""
        # Sheet switching keeps the same folder sample.
        sample = getattr(self, "_sample", "") if use_sample else ""
        self._requested_sheet = (sheet if preview_sheet is None else preview_sheet).strip()

        def finish(result=None, error=""):
            self._loading = False
            self.sheet_choice.setEnabled(True)
            if generation != self._generation:
                if self._pending_open and self.dialog.isVisible():
                    self._pending_open = False
                    self._load()
                return
            if error:
                self.status.setText(f"无法预览：{error}")
                return
            self._sample = str(result.path)
            self.model.set_preview(result)
            if result.sheet_fallback:
                self._requested_sheet = ""
            self.sheet_choice.blockSignals(True)
            self.sheet_choice.clear()
            self.sheet_choice.addItem("各文件活动工作表" if self.folder else "活动工作表", "")
            for name in result.sheet_names:
                self.sheet_choice.addItem(name, name)
            self.sheet_choice.setCurrentIndex(max(0, self.sheet_choice.findData(self._requested_sheet)))
            self.sheet_choice.blockSignals(False)
            alias_warning = self._fill_columns(result)
            self.apply_button.setEnabled(result.columns > 0)
            notes = "\n".join(note for note in (result.note, alias_warning) if note)
            self.status.setText(f"{result.path}\n工作表：{result.sheet} · 第 1–6 行（首行为表头）" +
                                (f"\n{notes}" if notes else ""))

        self.owner.run_in_background(
            read_excel_preview, args=(sample or path, self._requested_sheet),
            kwargs={"folder": self.folder and not sample, "first_sheet": self.first_sheet,
                    "exclude": tuple(picker.path() for picker in self.exclusions) if self.exclusions else (),
                    "exclude_names": self.exclude_names, "skip_links": self.skip_links,
                    "allow_sheet_fallback": not use_sample},
            on_success=lambda result: finish(result), on_error=lambda error: finish(error=error),
        )

    def _fill_columns(self, preview):
        from tools.excel_preview import TEXT_LIMIT
        headers = tuple(preview.headers.get(column, preview.cells.get((0, column), "")).strip()
                        for column in range(preview.columns))
        configured_aliases, alias_warning = None, ""
        if self.header_alias_store is not None:
            try:
                configured_aliases = self.header_alias_store.load()
            except (OSError, ValueError) as exc:
                alias_warning = f"表头别名配置读取失败，可手动选择列：{exc}"
        for choice, combo in zip(self.choices, self.column_choices):
            combo.clear()
            combo.addItem("不指定" if choice.optional else "请选择", "" if choice.optional else None)
            current = choice.field.text().strip()
            for column, header in enumerate(headers, 1):
                value = header if choice.kind == "header" else str(column) if choice.kind == "index" else _letter(column)
                if choice.kind == "header" and (not header or header == "[文本未展开]"):
                    continue
                combo.addItem(f"{_letter(column)} ({column}) · {header[:TEXT_LIMIT] or '空表头'}", value)
            selected = combo.findData(current)
            if selected < 0 and (choice.kind == "index" or not current):
                aliases = (current,) if current else choice.aliases
                matches = [index + 1 for index, header in enumerate(headers)
                           if header.casefold() in {alias.casefold() for alias in aliases}]
                if len(matches) == 1:
                    value = str(matches[0]) if choice.kind == "index" else _letter(matches[0])
                    selected = combo.findData(value)
            if not current and choice.alias_role and self.header_alias_store is not None:
                from tools.excel_metadata import _preferred_header_column
                value = _preferred_header_column(tuple(header.casefold() for header in headers),
                                                 getattr(configured_aliases, choice.alias_role, ()), choice.alias_role)
                selected = combo.findData(value) if value else 0
            if selected < 0 and current:
                combo.addItem(f"保留当前设置：{current}", current)
                selected = combo.count() - 1
            combo.setCurrentIndex(max(0, selected))
            combo.setEnabled(True)
        return alias_warning

    def _apply_columns(self):
        preview = self.model.preview
        if preview is None or self._loading:
            return
        values = [combo.currentData() for combo in self.column_choices]
        missing = [choice.label for choice, value in zip(self.choices, values)
                   if not choice.field.isHidden() and not choice.optional and not value]
        if missing:
            self.status.setText("请选择：" + "、".join(missing))
            return
        if self.sheet is not None:
            # Do not trigger the page's synchronous header detection again.
            blocked = self.sheet.blockSignals(True)
            if isinstance(self.sheet, QComboBox):
                self.sheet.clear()
                self.sheet.addItems(preview.sheet_names)
                self.sheet.setCurrentText(preview.sheet)
            else:
                self.sheet.setText(self._requested_sheet)
            self.sheet.blockSignals(blocked)
        for choice, value in zip(self.choices, values):
            if not choice.field.isHidden():
                choice.field.setText(value or "")
        self.dialog.accept()
