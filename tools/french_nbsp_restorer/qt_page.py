"""Unified Qt page for FrenchNbspPage."""

from __future__ import annotations

from PySide6.QtWidgets import (
    QComboBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)
from tools.excel_metadata import detect_source_target_columns, list_workbook_sheets
from tools.header_aliases import HeaderAliasStore
from tools.french_nbsp_restorer.restore_french_nbsp import (
    build_default_output_path as build_nbsp_output_path,
    process_excel as restore_french_nbsp,
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
from tools.qt_page_helpers import _add_action_bar, _choose_excel, _set_combo_values

_show_error = show_error


class FrenchNbspPage(AsyncPage):
    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        header_alias_store: HeaderAliasStore | None = None,
    ) -> None:
        super().__init__(parent)
        self.header_alias_store = header_alias_store or HeaderAliasStore()
        layout = QVBoxLayout(self)
        layout.setContentsMargins(2, 2, 6, 2)
        layout.setSpacing(8)
        input_box, input_layout = section("输入与范围")
        self.input_picker = PathPicker("输入 Excel")
        self.input_picker.choose_button.clicked.connect(self.choose_input)
        input_layout.addWidget(self.input_picker)
        scope = QHBoxLayout()
        self.sheet = QComboBox()
        self.sheet.setMinimumWidth(190)
        self.sheet.currentTextChanged.connect(self.detect_columns)
        self.target_column = QLineEdit("B")
        self.target_column.setPlaceholderText("请指定")
        self.target_column.setToolTip("没有唯一匹配的表头时，请手动填写列字母。")
        self.target_column.setMaximumWidth(80)
        self.start_row = QSpinBox()
        self.start_row.setRange(1, 1_000_000)
        self.start_row.setValue(2)
        self.start_row.setMaximumWidth(94)
        scope.addWidget(QLabel("处理工作表"))
        scope.addWidget(self.sheet)
        scope.addSpacing(8)
        scope.addWidget(QLabel("Target 列"))
        scope.addWidget(self.target_column)
        scope.addSpacing(8)
        scope.addWidget(QLabel("开始行"))
        scope.addWidget(self.start_row)
        scope.addStretch(1)
        input_layout.addLayout(scope)
        layout.addWidget(input_box)

        output_box, output_layout = section("输出设置")
        row = QHBoxLayout()
        self.result_column = QLineEdit()
        self.result_column.setMaximumWidth(100)
        row.addWidget(QLabel("结果列（可选）"))
        row.addWidget(self.result_column)
        row.addStretch(1)
        output_layout.addLayout(row)
        layout.addWidget(output_box)
        self.run_button = primary_button("开始恢复")
        self.run_button.clicked.connect(self.run_restore)
        self.preview = muted_label()
        layout.addWidget(self.preview)
        layout.addStretch(1)
        self.action_bar = _add_action_bar(layout, self.run_button)

    def choose_input(self) -> None:
        if path := _choose_excel(self, "选择 Excel 文件"):
            self.load_input_file(path)

    def load_input_file(self, path: str, *, reset_options: bool = False, show_error: bool = True) -> None:
        if reset_options:
            self.target_column.setText("B")
            self.result_column.clear()
            self.start_row.setValue(2)
        self.input_picker.set_path(path)
        try:
            choices = list_workbook_sheets(path)
        except Exception as exc:  # noqa: BLE001
            _set_combo_values(self.sheet, ())
            if show_error:
                _show_error(self, "读取失败", str(exc))
            return
        chosen = choices.default_sheet or (choices.sheet_names[0] if choices.sheet_names else "")
        _set_combo_values(self.sheet, choices.sheet_names, chosen)
        self.detect_columns(chosen, show_error=show_error)
        self.preview.setText(f"输出文件：{build_nbsp_output_path(path).name}")

    def detect_columns(self, _sheet: str = "", *, show_error: bool = False) -> None:
        if not self.input_picker.path() or not self.sheet.currentText():
            return
        try:
            columns = detect_source_target_columns(
                self.input_picker.path(),
                sheet=self.sheet.currentText(),
                header_aliases=self.header_alias_store.load(),
            )
        except Exception as exc:  # noqa: BLE001
            if show_error:
                _show_error(self, "读取失败", str(exc))
            return
        self.target_column.setText(columns.detected_target_column or "")

    def run_restore(self) -> None:
        if not self.input_picker.path():
            show_error(self, "缺少文件", "请先选择输入 Excel 文件。")
            return
        if not self.target_column.text().strip():
            show_error(self, "缺少列信息", "请填写 target 列。")
            return
        self.run_button.setEnabled(False)
        self.run_in_background(
            restore_french_nbsp,
            kwargs={
                "input_file": self.input_picker.path(),
                "target_column": self.target_column.text().strip(),
                "result_column": self.result_column.text().strip() or None,
                "sheet": self.sheet.currentText() or None,
                "start_row": self.start_row.value(),
                "output_file": None,
            },
            on_success=self._finish,
            on_error=self._fail,
        )

    def _finish(self, summary: object) -> None:
        self.run_button.setEnabled(True)
        show_info(
            self,
            "处理完成",
            "\n".join(
                [
                    "法语 NBSP 恢复已完成。",
                    f"工作表: {summary.worksheet_title}",
                    f"处理行数: {summary.processed_count}",
                    f"修复行数: {summary.changed_count}",
                    f"输出文件: {summary.output_path}",
                ]
            ),
        )

    def _fail(self, message: str) -> None:
        self.run_button.setEnabled(True)
        show_error(self, "处理失败", message)
