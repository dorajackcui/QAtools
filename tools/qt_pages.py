"""PySide6 pages used by the unified Toolshub window."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QFormLayout,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QInputDialog,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSpinBox,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from phraseloom.strings_workflow import (
    default_restored_output_path,
    default_strings_output_path,
    export_strings_workbook,
    restore_strings_workbook,
)
from tools.excel_batcher.excel_batcher import (
    build_default_output_dir as build_batch_output_dir,
    build_default_restore_path,
    restore_batches,
    split_workbook,
)
from tools.excel_merger.merge_active_sheets import (
    build_default_output_path as build_merge_output_path,
    merge_active_sheets,
)
from tools.excel_metadata import detect_source_target_columns, list_workbook_sheets
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
    show_warning,
)
from tools.target_text_checker.check_target_text import (
    ABNORMAL_PUNCTUATION_RULE,
    CONSECUTIVE_SPACES_RULE,
    MIXED_WIDTH_RULE,
)
from tools.tb_projects import TbProject, TbProjectStore
from tools.term_pair_checker.extract_terms_from_excel import (
    TERM_SHEET_NAME,
    detect_history_tb_columns,
)
from tools.workflow.revision_applier import (
    apply_workflow_revisions,
    build_default_revised_output_path,
)
from tools.workflow.workflow_runner import (
    build_default_output_path as build_workflow_output_path,
    run_workflow,
)
from tools.xbench_report_transformer.transform_xbench_report import (
    build_default_output_path as build_xbench_output_path,
    process_excel as transform_xbench_report,
)


EXCEL_FILTER = "Excel 文件 (*.xlsx *.xlsm);;所有文件 (*)"
_show_error = show_error


def _scroll_page(content: QWidget) -> QScrollArea:
    area = QScrollArea()
    area.setWidgetResizable(True)
    area.setFrameShape(QScrollArea.Shape.NoFrame)
    area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
    area.setWidget(content)
    return area


def _add_action_bar(layout: QVBoxLayout, *buttons: QPushButton) -> QFrame:
    """Pin consistently sized page actions to a quiet bottom bar."""

    bar = QFrame()
    bar.setObjectName("pageActionBar")
    bar.setFixedHeight(46)
    row = QHBoxLayout(bar)
    row.setContentsMargins(2, 9, 6, 2)
    row.setSpacing(8)
    row.addStretch(1)
    for button in buttons:
        button.setFixedWidth(136 if button.property("primary") else 108)
        row.addWidget(button)
    layout.addWidget(bar)
    return bar


def _choose_excel(parent: QWidget, title: str, initial_dir: str = "") -> str:
    path, _ = QFileDialog.getOpenFileName(parent, title, initial_dir, EXCEL_FILTER)
    return path


def _set_combo_values(combo: QComboBox, values: tuple[str, ...], selected: str = "") -> str:
    combo.blockSignals(True)
    combo.clear()
    combo.addItems(values)
    chosen = selected if selected in values else (values[0] if values else "")
    if chosen:
        combo.setCurrentText(chosen)
    combo.blockSignals(False)
    return chosen


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


class FrenchNbspPage(AsyncPage):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
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
        output_layout.addWidget(
            muted_label("留空时直接修复 Target 列；恢复 ; : ? ! % 前及 « » 内侧的 NBSP。", word_wrap=True)
        )
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
            columns = detect_source_target_columns(self.input_picker.path(), sheet=self.sheet.currentText())
        except Exception as exc:  # noqa: BLE001
            if show_error:
                _show_error(self, "读取失败", str(exc))
            return
        if columns.detected_target_column:
            self.target_column.setText(columns.detected_target_column)

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
        input_layout.addWidget(muted_label("将 QA 明细整理为文件名 / key / source / target / QA 问题，并按相同内容聚合。", word_wrap=True))
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
        input_layout.addWidget(muted_label(
            "递归读取目录中的 .xlsx/.xlsm，合并每个文件当前活动的工作表，并在首列写入 SourceFile。默认只保留第一份表头。",
            word_wrap=True,
        ))
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


class WorkflowPage(AsyncPage):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.last_workflow_output_path = ""
        self.tb_store = TbProjectStore()

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        content = QWidget()
        self.content_layout = QVBoxLayout(content)
        self.content_layout.setContentsMargins(2, 2, 6, 2)
        self.content_layout.setSpacing(8)
        self._build_input_section()
        self._build_task_section()
        self._build_term_settings()
        self._build_tag_settings()
        self._build_target_text_settings()

        self.run_button = primary_button("开始检查")
        self.revision_button = QPushButton("应用修订")
        self.run_button.clicked.connect(self.run_selected_tasks)
        self.revision_button.clicked.connect(self.apply_revisions)
        self.output_preview = muted_label()
        self.status = muted_label()
        self.content_layout.addWidget(self.output_preview)
        self.content_layout.addWidget(self.status)
        self.content_layout.addStretch(1)
        self.content_scroll = _scroll_page(content)
        outer.addWidget(self.content_scroll, 1)
        self.action_bar = _add_action_bar(
            outer,
            self.revision_button,
            self.run_button,
        )

    def _build_input_section(self) -> None:
        box, layout = section("输入与范围")
        self.input_picker = PathPicker("输入 Excel")
        self.input_picker.choose_button.clicked.connect(self.choose_input_file)
        layout.addWidget(self.input_picker)
        scope = QHBoxLayout()
        self.sheet = QComboBox()
        self.sheet.setMinimumWidth(180)
        self.sheet.currentTextChanged.connect(self.detect_main_columns)
        self.source_column = QLineEdit("A")
        self.source_column.setMaximumWidth(70)
        self.target_column = QLineEdit("B")
        self.target_column.setMaximumWidth(70)
        self.start_row = QSpinBox()
        self.start_row.setRange(1, 1_000_000)
        self.start_row.setValue(2)
        self.start_row.setMaximumWidth(94)
        for label, widget in (
            ("检查工作表", self.sheet),
            ("Source 列", self.source_column),
            ("Target 列", self.target_column),
            ("开始行", self.start_row),
        ):
            scope.addWidget(QLabel(label))
            scope.addWidget(widget)
            scope.addSpacing(8)
        scope.addStretch(1)
        layout.addLayout(scope)
        self.content_layout.addWidget(box)

    def _task_row(self, checkbox: QCheckBox, settings_button: QPushButton | None = None) -> QWidget:
        widget = QWidget()
        row = QHBoxLayout(widget)
        row.setContentsMargins(0, 0, 0, 0)
        row.addWidget(checkbox)
        if settings_button is not None:
            settings_button.setCheckable(True)
            row.addWidget(settings_button)
        row.addStretch(1)
        return widget

    def _build_task_section(self) -> None:
        box, layout = section("质量检查项目")
        header = QHBoxLayout()
        header.addWidget(muted_label("默认全部选中，可按需取消"))
        header.addStretch(1)
        select_all = QPushButton("全选")
        clear_all = QPushButton("取消全选")
        select_all.clicked.connect(lambda: self.set_all_tasks(True))
        clear_all.clicked.connect(lambda: self.set_all_tasks(False))
        header.addWidget(select_all)
        header.addWidget(clear_all)
        layout.addLayout(header)
        grid = QGridLayout()
        self.term_check = QCheckBox("术语检查")
        self.tag_check = QCheckBox("Tag 检查")
        self.line_break_check = QCheckBox("换行数量检查")
        self.consistency_check = QCheckBox("同源译文一致性")
        self.chinese_check = QCheckBox("Target 中文检查")
        self.target_text_check = QCheckBox("Target 文本规范检查")
        self.task_checks = (
            self.term_check,
            self.tag_check,
            self.line_break_check,
            self.consistency_check,
            self.chinese_check,
            self.target_text_check,
        )
        for check in self.task_checks:
            check.setChecked(True)
        self.term_settings_button = QPushButton("展开设置")
        self.tag_settings_button = QPushButton("展开设置")
        self.target_settings_button = QPushButton("展开设置")
        grid.addWidget(self._task_row(self.term_check, self.term_settings_button), 0, 0)
        grid.addWidget(self._task_row(self.tag_check, self.tag_settings_button), 0, 1)
        grid.addWidget(self._task_row(self.line_break_check), 1, 0)
        grid.addWidget(self._task_row(self.consistency_check), 1, 1)
        grid.addWidget(self._task_row(self.chinese_check), 2, 0)
        grid.addWidget(self._task_row(self.target_text_check, self.target_settings_button), 2, 1)
        for column in range(2):
            grid.setColumnStretch(column, 1)
        layout.addLayout(grid)
        self.content_layout.addWidget(box)

        self.term_settings_button.toggled.connect(lambda visible: self._toggle_settings("term", visible))
        self.tag_settings_button.toggled.connect(lambda visible: self._toggle_settings("tag", visible))
        self.target_settings_button.toggled.connect(lambda visible: self._toggle_settings("target", visible))
        self.term_check.toggled.connect(lambda enabled: self._task_toggled(enabled, self.term_settings_button))
        self.tag_check.toggled.connect(lambda enabled: self._task_toggled(enabled, self.tag_settings_button))
        self.target_text_check.toggled.connect(lambda enabled: self._task_toggled(enabled, self.target_settings_button))

    def _build_term_settings(self) -> None:
        self.term_settings_box, layout = section("术语检查设置")
        mark_row = QHBoxLayout()
        mark_row.addWidget(QLabel("术语标记"))
        self.mark_book = QCheckBox("中文方括号【】")
        self.mark_square = QCheckBox("半角方括号 []")
        self.mark_book.setChecked(True)
        self.mark_square.setChecked(True)
        mark_row.addWidget(self.mark_book)
        mark_row.addWidget(self.mark_square)
        mark_row.addStretch(1)
        layout.addLayout(mark_row)

        project_row = QHBoxLayout()
        project_row.addWidget(QLabel("TB 项目"))
        self.tb_project = QComboBox()
        self.tb_project.setMinimumWidth(230)
        self.tb_project.currentTextChanged.connect(self.load_selected_tb_project)
        save_project = QPushButton("保存当前")
        delete_project = QPushButton("删除")
        save_project.clicked.connect(self.save_tb_project)
        delete_project.clicked.connect(self.delete_tb_project)
        project_row.addWidget(self.tb_project)
        project_row.addWidget(save_project)
        project_row.addWidget(delete_project)
        project_row.addStretch(1)
        layout.addLayout(project_row)

        self.history_picker = PathPicker("历史 TB（可选）", allow_clear=True)
        self.history_picker.choose_button.clicked.connect(self.choose_history_file)
        self.history_picker.clear_button.clicked.connect(self.clear_history_fields)
        layout.addWidget(self.history_picker)
        history_scope = QHBoxLayout()
        self.history_sheet = QComboBox()
        self.history_sheet.setMinimumWidth(170)
        self.history_sheet.currentTextChanged.connect(self.detect_history_columns)
        self.history_source = QLineEdit()
        self.history_source.setMaximumWidth(70)
        self.history_target = QLineEdit()
        self.history_target.setMaximumWidth(70)
        self.history_start_row = QSpinBox()
        self.history_start_row.setRange(1, 1_000_000)
        self.history_start_row.setValue(2)
        self.history_start_row.setMaximumWidth(94)
        for label, widget in (
            ("工作表", self.history_sheet),
            ("Source 列", self.history_source),
            ("Target 列", self.history_target),
            ("开始行", self.history_start_row),
        ):
            history_scope.addWidget(QLabel(label))
            history_scope.addWidget(widget)
            history_scope.addSpacing(8)
        history_scope.addStretch(1)
        layout.addLayout(history_scope)
        layout.addWidget(muted_label("未选择术语标记时，必须提供历史 TB。"))
        self.term_settings_box.hide()
        self.content_layout.addWidget(self.term_settings_box)
        self.refresh_tb_projects()

    def _build_tag_settings(self) -> None:
        self.tag_settings_box, layout = section("Tag 检查设置")
        mode_row = QHBoxLayout()
        mode_row.addWidget(QLabel("检查模式"))
        mode_switch = QFrame()
        mode_switch.setObjectName("segmentedControl")
        mode_switch_layout = QHBoxLayout(mode_switch)
        mode_switch_layout.setContentsMargins(1, 1, 1, 1)
        mode_switch_layout.setSpacing(0)
        self.standard_mode = QPushButton("常规 Tag")
        self.memoq_mode = QPushButton("memoQ Marker")
        for mode in (self.standard_mode, self.memoq_mode):
            mode.setCheckable(True)
            mode.setProperty("segmentedMode", True)
            mode.setCursor(Qt.CursorShape.PointingHandCursor)
            mode_switch_layout.addWidget(mode)
        self.standard_mode.setChecked(True)
        self.tag_mode_group = QButtonGroup(self)
        self.tag_mode_group.setExclusive(True)
        self.tag_mode_group.addButton(self.standard_mode, 0)
        self.tag_mode_group.addButton(self.memoq_mode, 1)
        self.standard_mode.toggled.connect(self.update_tag_mode)
        mode_row.addWidget(mode_switch)
        mode_row.addStretch(1)
        layout.addLayout(mode_row)
        layout.addWidget(QLabel("常规类型"))
        type_grid = QGridLayout()
        self.angle_tag = QCheckBox("<...> tag")
        self.color_tag = QCheckBox("[color=...] tag")
        self.brace_tag = QCheckBox("{...} placeholder")
        self.newline_tag = QCheckBox(r"\n mark")
        self.standard_tag_checks = (self.angle_tag, self.color_tag, self.brace_tag, self.newline_tag)
        for index, check in enumerate(self.standard_tag_checks):
            check.setChecked(True)
            type_grid.addWidget(check, index // 2, index % 2)
        for column in range(2):
            type_grid.setColumnStretch(column, 1)
        layout.addLayout(type_grid)
        self.angle_config = PathPicker("尖括号过滤配置", allow_clear=True)
        self.angle_config.choose_button.clicked.connect(self.choose_angle_config)
        layout.addWidget(self.angle_config)
        self.tag_settings_box.hide()
        self.content_layout.addWidget(self.tag_settings_box)

    def _build_target_text_settings(self) -> None:
        self.target_settings_box, layout = section("Target 文本规范检查设置")
        layout.addWidget(QLabel("检查规则"))
        grid = QGridLayout()
        self.abnormal_rule = QCheckBox("异常标点符号（.. / ,, / 。。等）")
        self.spaces_rule = QCheckBox("连续空格（2 个及以上）")
        self.width_rule = QCheckBox("全半角混用")
        self.rule_checks = {
            ABNORMAL_PUNCTUATION_RULE: self.abnormal_rule,
            CONSECUTIVE_SPACES_RULE: self.spaces_rule,
            MIXED_WIDTH_RULE: self.width_rule,
        }
        for index, check in enumerate(self.rule_checks.values()):
            check.setChecked(True)
            grid.addWidget(check, index // 2, index % 2)
        for column in range(2):
            grid.setColumnStretch(column, 1)
        layout.addLayout(grid)
        self.target_settings_box.hide()
        self.content_layout.addWidget(self.target_settings_box)

    def _toggle_settings(self, name: str, visible: bool) -> None:
        mapping = {
            "term": (self.term_settings_box, self.term_settings_button),
            "tag": (self.tag_settings_box, self.tag_settings_button),
            "target": (self.target_settings_box, self.target_settings_button),
        }
        box, button = mapping[name]
        if visible:
            for other_name, (other_box, other_button) in mapping.items():
                if other_name != name:
                    other_button.blockSignals(True)
                    other_button.setChecked(False)
                    other_button.setText("展开设置")
                    other_button.blockSignals(False)
                    other_box.hide()
        box.setVisible(visible)
        button.setText("收起设置" if visible else "展开设置")

    def _task_toggled(self, enabled: bool, button: QPushButton) -> None:
        button.setEnabled(enabled)
        if not enabled:
            button.setChecked(False)

    def set_all_tasks(self, checked: bool) -> None:
        for checkbox in self.task_checks:
            checkbox.setChecked(checked)

    def choose_input_file(self) -> None:
        if path := _choose_excel(self, "选择 Excel 文件"):
            self.load_input_file(path)

    def load_input_file(self, path: str, *, show_error: bool = True) -> None:
        self.input_picker.set_path(path)
        self.last_workflow_output_path = ""
        self.output_preview.setText(f"输出文件：{build_workflow_output_path(path).name}")
        try:
            choices = list_workbook_sheets(path)
        except Exception as exc:  # noqa: BLE001
            _set_combo_values(self.sheet, ())
            if show_error:
                _show_error(self, "读取失败", str(exc))
            return
        selected = choices.default_sheet or (choices.sheet_names[0] if choices.sheet_names else "")
        _set_combo_values(self.sheet, choices.sheet_names, selected)
        self.detect_main_columns(selected, show_error=show_error)

    def detect_main_columns(self, _sheet: str = "", *, show_error: bool = False) -> None:
        if not self.input_picker.path() or not self.sheet.currentText():
            return
        try:
            columns = detect_source_target_columns(self.input_picker.path(), sheet=self.sheet.currentText())
        except Exception as exc:  # noqa: BLE001
            if show_error:
                _show_error(self, "读取失败", str(exc))
            return
        if columns.detected_source_column:
            self.source_column.setText(columns.detected_source_column)
        if columns.detected_target_column:
            self.target_column.setText(columns.detected_target_column)

    def choose_history_file(self) -> None:
        if path := _choose_excel(self, "选择术语历史 TB Excel 文件"):
            self.history_picker.set_path(path)
            self.refresh_history_sheets()

    def clear_history_fields(self) -> None:
        _set_combo_values(self.history_sheet, ())
        self.history_source.clear()
        self.history_target.clear()
        self.history_start_row.setValue(2)

    def refresh_history_sheets(self, *, show_error: bool = True, selected: str = "") -> None:
        if not self.history_picker.path():
            self.clear_history_fields()
            return
        try:
            choices = list_workbook_sheets(self.history_picker.path())
        except Exception as exc:  # noqa: BLE001
            self.clear_history_fields()
            if show_error:
                _show_error(self, "读取失败", str(exc))
            return
        desired = selected or (TERM_SHEET_NAME if TERM_SHEET_NAME in choices.sheet_names else choices.default_sheet or "")
        chosen = _set_combo_values(self.history_sheet, choices.sheet_names, desired)
        self.detect_history_columns(chosen, show_error=show_error)

    def detect_history_columns(self, _sheet: str = "", *, show_error: bool = False) -> None:
        if not self.history_picker.path() or not self.history_sheet.currentText():
            return
        try:
            columns = detect_history_tb_columns(self.history_picker.path(), sheet=self.history_sheet.currentText())
        except Exception as exc:  # noqa: BLE001
            if show_error:
                _show_error(self, "读取失败", str(exc))
            return
        if columns.source_column:
            self.history_source.setText(columns.source_column)
        if columns.target_column:
            self.history_target.setText(columns.target_column)

    def choose_angle_config(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "选择尖括号 Tag 过滤配置", "", "JSON 文件 (*.json);;所有文件 (*)")
        if path:
            self.angle_config.set_path(path)

    def update_tag_mode(self, standard: bool) -> None:
        for check in self.standard_tag_checks:
            check.setEnabled(standard)
        self.angle_config.setEnabled(standard)

    def selected_term_marks(self) -> tuple[str, ...]:
        return tuple(mark for mark, check in (("【】", self.mark_book), ("[]", self.mark_square)) if check.isChecked())

    def selected_tag_types(self) -> tuple[str, ...]:
        if self.memoq_mode.isChecked():
            return ("memoq",)
        return tuple(name for name, check in (
            ("angle", self.angle_tag),
            ("square_color", self.color_tag),
            ("brace", self.brace_tag),
            ("newline", self.newline_tag),
        ) if check.isChecked())

    def selected_target_rules(self) -> tuple[str, ...]:
        return tuple(rule for rule, check in self.rule_checks.items() if check.isChecked())

    def refresh_tb_projects(self, selected: str = "") -> None:
        try:
            names = tuple(project.name for project in self.tb_store.list_projects())
        except ValueError as exc:
            show_error(self, "TB 项目读取失败", str(exc))
            names = ()
        self.tb_project.blockSignals(True)
        self.tb_project.clear()
        self.tb_project.addItem("")
        self.tb_project.addItems(names)
        if selected in names:
            self.tb_project.setCurrentText(selected)
        self.tb_project.blockSignals(False)

    def load_selected_tb_project(self, name: str) -> None:
        if not name:
            return
        try:
            project = self.tb_store.find_project(name)
        except ValueError as exc:
            show_error(self, "TB 项目读取失败", str(exc))
            return
        if project is None:
            return
        self.history_picker.set_path(project.file_path)
        self.refresh_history_sheets(show_error=False, selected=project.sheet)
        self.history_source.setText(project.source_column)
        self.history_target.setText(project.target_column)
        self.history_start_row.setValue(project.start_row)
        if not Path(project.file_path).is_file():
            show_warning(self, "TB 文件不存在", f"项目“{project.name}”对应的 TB 文件已移动或不存在。")

    def _capture_tb_project(self, name: str) -> TbProject:
        if not self.history_picker.path() or not Path(self.history_picker.path()).expanduser().is_file():
            raise ValueError("请先选择有效的历史 TB 文件。")
        if not self.history_sheet.currentText() or not self.history_source.text().strip() or not self.history_target.text().strip():
            raise ValueError("请先确认历史 TB 的工作表及 Source / Target 列。")
        return TbProject(
            name=name,
            file_path=str(Path(self.history_picker.path()).expanduser().absolute()),
            sheet=self.history_sheet.currentText(),
            source_column=self.history_source.text().strip(),
            target_column=self.history_target.text().strip(),
            start_row=self.history_start_row.value(),
        )

    def save_tb_project(self) -> None:
        name, accepted = QInputDialog.getText(self, "保存 TB 项目", "项目名称：", text=self.tb_project.currentText())
        name = name.strip()
        if not accepted:
            return
        if not name:
            show_error(self, "项目名称为空", "请输入项目名称。")
            return
        try:
            project = self._capture_tb_project(name)
            existing = self.tb_store.find_project(name)
            if existing is not None and QMessageBox.question(
                self,
                "更新 TB 项目",
                f"项目“{existing.name}”已存在，是否用当前设置更新？",
            ) != QMessageBox.StandardButton.Yes:
                return
            self.tb_store.save_project(project)
        except (OSError, ValueError) as exc:
            show_error(self, "无法保存 TB 项目", str(exc))
            return
        self.refresh_tb_projects(name)

    def delete_tb_project(self) -> None:
        name = self.tb_project.currentText().strip()
        if not name:
            show_info(self, "未选择项目", "请先选择要删除的 TB 项目。")
            return
        if QMessageBox.question(self, "删除 TB 项目", f"确定删除项目“{name}”吗？\n不会删除原始 TB 文件。") != QMessageBox.StandardButton.Yes:
            return
        try:
            self.tb_store.delete_project(name)
        except (OSError, ValueError) as exc:
            show_error(self, "无法删除 TB 项目", str(exc))
            return
        self.refresh_tb_projects()

    def run_selected_tasks(self) -> None:
        input_file = self.input_picker.path()
        source_column = self.source_column.text().strip()
        target_column = self.target_column.text().strip()
        task_values = tuple(check.isChecked() for check in self.task_checks)
        term_marks = self.selected_term_marks()
        tag_types = self.selected_tag_types()
        target_rules = self.selected_target_rules()
        history_file = self.history_picker.path()
        if not input_file:
            show_error(self, "缺少文件", "请先选择输入 Excel 文件。")
            return
        if not source_column or not target_column:
            show_error(self, "缺少列信息", "请填写 source 列和 target 列。")
            return
        if not any(task_values):
            show_error(self, "缺少任务", "请至少选择一个质量检查项目。")
            return
        if self.term_check.isChecked() and not term_marks and not history_file:
            show_error(self, "缺少术语来源", "术语检查至少需要一种术语 mark，或一个历史 TB。")
            return
        if self.tag_check.isChecked() and not tag_types:
            show_error(self, "缺少检查类型", "Tag检查至少需要一种检查类型。")
            return
        if self.target_text_check.isChecked() and not target_rules:
            show_error(self, "缺少检查规则", "Target 文本规范检查至少需要选择一项规则。")
            return

        self.run_button.setEnabled(False)
        self.revision_button.setEnabled(False)
        self.status.setText("正在执行质量检查…")
        self.run_in_background(
            run_workflow,
            kwargs={
                "input_file": input_file,
                "output_file": None,
                "source_column": source_column,
                "target_column": target_column,
                "sheet": self.sheet.currentText() or None,
                "start_row": self.start_row.value(),
                "run_term_pair_check": self.term_check.isChecked(),
                "term_mark_styles": term_marks,
                "term_history_tb_file": history_file or None,
                "term_history_sheet": self.history_sheet.currentText() or None if history_file else None,
                "term_history_source_column": self.history_source.text().strip() or None if history_file else None,
                "term_history_target_column": self.history_target.text().strip() or None if history_file else None,
                "term_history_start_row": self.history_start_row.value(),
                "run_tag_check": self.tag_check.isChecked(),
                "tag_token_types": tag_types,
                "tag_angle_config_file": self.angle_config.path() or None if "angle" in tag_types else None,
                "run_line_break_check": self.line_break_check.isChecked(),
                "run_source_consistency_check": self.consistency_check.isChecked(),
                "run_chinese_target_check": self.chinese_check.isChecked(),
                "run_target_text_check": self.target_text_check.isChecked(),
                "target_text_rules": target_rules,
            },
            on_success=lambda summary: self._finish_workflow(summary, history_file),
            on_error=self._fail_workflow,
        )

    def _finish_workflow(self, summary: object, history_file: str) -> None:
        self.run_button.setEnabled(True)
        self.revision_button.setEnabled(True)
        self.status.clear()
        self.last_workflow_output_path = str(summary.output_path)
        lines = [
            "一键质量检查完成。",
            f"检查工作表: {summary.worksheet_title}",
            f"source 列: {summary.source_column}",
            f"target 列: {summary.target_column}",
        ]
        if summary.ran_term_pair_check:
            lines.extend((f"术语表条目数: {summary.term_count}", f"术语问题行数: {summary.term_problem_rows}"))
            if history_file:
                lines.append(f"术语历史 TB: {history_file}")
        if summary.ran_tag_check:
            lines.append(f"Tag问题行数: {summary.tag_problem_rows}")
        if summary.ran_line_break_check:
            lines.append(f"换行数量问题行数: {summary.line_break_problem_count}")
        if summary.ran_source_consistency_check:
            lines.extend((
                f"同源译文不一致 source 数: {summary.source_consistency_problem_count}",
                f"同源译文不一致涉及行数: {summary.source_consistency_problem_rows}",
            ))
        if summary.ran_chinese_target_check:
            lines.append(f"Target 中文问题行数: {summary.chinese_target_problem_count}")
        if summary.ran_target_text_check:
            lines.append(f"Target 文本规范问题行数: {summary.target_text_problem_rows}")
        lines.append(f"输出文件: {summary.output_path}")
        show_info(self, "处理完成", "\n".join(lines))

    def _fail_workflow(self, message: str) -> None:
        self.run_button.setEnabled(True)
        self.revision_button.setEnabled(True)
        self.status.clear()
        show_error(self, "处理失败", message)

    def apply_revisions(self) -> None:
        candidate_text = self.last_workflow_output_path or self.input_picker.path()
        candidate = Path(candidate_text).expanduser() if candidate_text else None
        report_file, _ = QFileDialog.getOpenFileName(
            self,
            "选择已填写的问题处理 Excel",
            str(candidate) if candidate else "",
            EXCEL_FILTER,
        )
        if not report_file:
            return
        try:
            default_output = build_default_revised_output_path(report_file)
        except Exception as exc:  # noqa: BLE001
            show_error(self, "读取失败", str(exc))
            return
        output_file, _ = QFileDialog.getSaveFileName(self, "保存修订稿", str(default_output), EXCEL_FILTER)
        if not output_file:
            return
        self.run_button.setEnabled(False)
        self.revision_button.setEnabled(False)
        self.status.setText("正在应用修订…")
        self.run_in_background(
            apply_workflow_revisions,
            args=(report_file,),
            kwargs={"output_file": output_file},
            on_success=self._finish_revision,
            on_error=self._fail_revision,
        )

    def _finish_revision(self, summary: object) -> None:
        self.run_button.setEnabled(True)
        self.revision_button.setEnabled(True)
        self.status.clear()
        lines = [
            "修订稿已生成。",
            f"回填修改: {summary.revised_count} 行",
            f"未填写（忽略）: {summary.ignored_count} 行",
            f"内容未变化: {summary.unchanged_count} 行",
        ]
        if summary.conflict_rows:
            lines.append("因原 target 已变化而跳过: " + "、".join(str(row) for row in summary.conflict_rows))
        lines.append(f"输出文件: {summary.output_path}")
        if summary.conflict_rows:
            show_warning(self, "修订稿已生成（存在冲突）", "\n".join(lines))
        else:
            show_info(self, "修订稿已生成", "\n".join(lines))

    def _fail_revision(self, message: str) -> None:
        self.run_button.setEnabled(True)
        self.revision_button.setEnabled(True)
        self.status.clear()
        show_error(self, "应用失败", message)


PAGE_FACTORIES = {
    "workflow": WorkflowPage,
    "phraseloom": PhraseLoomPage,
    "french_nbsp": FrenchNbspPage,
    "excel_batcher": ExcelBatcherPage,
    "excel_merger": ExcelMergerPage,
    "xbench_report": XbenchPage,
}


__all__ = [
    "ExcelBatcherPage",
    "ExcelMergerPage",
    "FrenchNbspPage",
    "PAGE_FACTORIES",
    "PhraseLoomPage",
    "WorkflowPage",
    "XbenchPage",
]
