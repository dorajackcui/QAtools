"""Unified Qt page for WorkflowPage."""

from __future__ import annotations

from pathlib import Path
from typing import Any
from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QInputDialog,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QToolButton,
    QVBoxLayout,
    QWidget,
)
from tools.excel_metadata import detect_source_target_columns, list_workbook_sheets
from tools.header_aliases import HeaderAliasStore
from tools.qt_gui_common import (
    AsyncPage,
    PathPicker,
    horizontal_rule,
    muted_label,
    primary_button,
    section,
    show_error,
    show_info,
    show_warning,
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
from tools.workflow.gui_options import WorkflowOptionsStore
from tools.workflow.workflow_runner import (
    build_default_output_path as build_workflow_output_path,
    run_workflow,
)
from tools.qt_page_helpers import (
    _scroll_page,
    _add_action_bar,
    _choose_excel,
    _set_combo_values,
    EXCEL_FILTER,
)
from tools.workflow.qt_settings import WorkflowSettingsMixin

_show_error = show_error


class WorkflowPage(WorkflowSettingsMixin, AsyncPage):
    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        header_alias_store: HeaderAliasStore | None = None,
        options_store: WorkflowOptionsStore | None = None,
    ) -> None:
        super().__init__(parent)
        self.last_workflow_output_path = ""
        self.tb_store = TbProjectStore()
        self.header_alias_store = header_alias_store or HeaderAliasStore()
        self.options_store = options_store or WorkflowOptionsStore()
        self._settings_snapshots: dict[str, dict[str, Any]] = {}

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
        self._build_substring_settings()

        self.run_button = primary_button("开始检查")
        self.revision_button = QPushButton("应用修订")
        self.open_report_button = QPushButton("打开报告")
        self.open_report_button.setEnabled(False)
        self.open_report_button.setToolTip("使用系统默认应用打开本次检查报告。")
        self.run_button.clicked.connect(self.run_selected_tasks)
        self.revision_button.clicked.connect(self.apply_revisions)
        self.open_report_button.clicked.connect(self.open_report)
        self.input_picker.path_changed.connect(self._reset_workflow_report)
        self.output_preview = muted_label()
        self.status = muted_label()
        self.content_layout.addWidget(self.output_preview)
        self.content_layout.addWidget(self.status)
        self.content_layout.addStretch(1)
        self.content_scroll = _scroll_page(content)
        outer.addWidget(self.content_scroll, 1)
        self.action_bar = _add_action_bar(
            outer,
            self.open_report_button,
            self.revision_button,
            self.run_button,
        )
        try:
            self._restore_options(self.options_store.load(self._capture_options()))
        except (OSError, ValueError) as exc:
            self.status.setText(f"未能恢复记住的选项，已使用默认设置：{exc}")

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
        for column_input in (self.source_column, self.target_column):
            column_input.setPlaceholderText("请指定")
            column_input.setToolTip("没有唯一匹配的表头时，请手动填写列字母。")
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

    def _task_row(
        self,
        checkbox: QCheckBox,
        settings_button: QToolButton | None = None,
    ) -> QWidget:
        widget = QWidget()
        row = QHBoxLayout(widget)
        row.setContentsMargins(0, 0, 0, 0)
        row.addWidget(checkbox)
        if settings_button is not None:
            row.addSpacing(6)
            row.addWidget(settings_button)
        row.addStretch(1)
        return widget

    @staticmethod
    def _settings_button(tool_tip: str) -> QToolButton:
        button = QToolButton()
        button.setArrowType(Qt.ArrowType.DownArrow)
        button.setToolTip(tool_tip)
        button.setAccessibleName(tool_tip)
        button.setProperty("settingsButton", True)
        button.setCursor(Qt.CursorShape.PointingHandCursor)
        button.setFixedSize(30, 30)
        return button

    def _build_task_section(self) -> None:
        box, layout = section()
        header = QHBoxLayout()
        header.addStretch(1)
        select_all = QPushButton("全选")
        clear_all = QPushButton("清空")
        self.remember_options_button = QPushButton("记住选项")
        self.remember_options_button.setToolTip("记住当前输入范围、检查项目及详细设置，下次打开时恢复。")
        self.remember_options_button.clicked.connect(self.remember_options)
        select_all.clicked.connect(lambda: self.set_all_tasks(True))
        clear_all.clicked.connect(lambda: self.set_all_tasks(False))
        header.addWidget(select_all)
        header.addWidget(clear_all)
        header.addWidget(self.remember_options_button)
        layout.addLayout(header)

        self.term_check = QCheckBox("术语检查")
        self.consistency_check = QCheckBox("同 Source 不同 Target")
        self.target_consistency_check = QCheckBox("同 Target 不同 Source")
        self.substring_consistency_check = QCheckBox("子串译文一致性")
        self.substring_consistency_check.setToolTip(
            "检查长原文是否沿用其包含短句的参考译文；结果为疑似问题，需要人工复核。"
        )
        self.tag_check = QCheckBox("Tag / Placeholder")
        self.line_break_check = QCheckBox("换行数量")
        self.number_check = QCheckBox("数字一致性")
        self.url_check = QCheckBox("URL 一致性")
        self.chinese_check = QCheckBox("Target 中文")
        self.target_text_check = QCheckBox("Target 文本规范")
        self.task_checks = (
            self.term_check,
            self.consistency_check,
            self.target_consistency_check,
            self.tag_check,
            self.line_break_check,
            self.number_check,
            self.url_check,
            self.chinese_check,
            self.target_text_check,
            self.substring_consistency_check,  # Append to preserve saved option positions.
        )
        for check in self.task_checks:
            check.setChecked(True)
        self.target_consistency_check.setChecked(False)
        self.substring_consistency_check.setChecked(False)

        self.substring_settings_button = self._settings_button("子串译文一致性设置")
        self.substring_settings_button.setEnabled(self.substring_consistency_check.isChecked())
        self.term_settings_button = self._settings_button("术语检查设置")
        self.tag_settings_button = self._settings_button("Tag / Placeholder 设置")
        self.target_settings_button = self._settings_button("Target 文本规范设置")

        translation_heading = QLabel("术语与翻译一致性")
        translation_heading.setObjectName("sectionTitle")
        layout.addWidget(translation_heading)
        translation_grid = QGridLayout()
        translation_grid.addWidget(
            self._task_row(self.term_check, self.term_settings_button), 0, 0
        )
        translation_grid.addWidget(self._task_row(self.substring_consistency_check, self.substring_settings_button), 0, 1)
        translation_grid.addWidget(self._task_row(self.consistency_check), 1, 0)
        translation_grid.addWidget(
            self._task_row(self.target_consistency_check), 1, 1
        )
        for column in range(2):
            translation_grid.setColumnStretch(column, 1)
        layout.addLayout(translation_grid)

        layout.addWidget(horizontal_rule())
        fidelity_heading = QLabel("内容保真检查")
        fidelity_heading.setObjectName("sectionTitle")
        layout.addWidget(fidelity_heading)
        fidelity_grid = QGridLayout()
        fidelity_grid.addWidget(
            self._task_row(self.tag_check, self.tag_settings_button), 0, 0
        )
        fidelity_grid.addWidget(self._task_row(self.line_break_check), 0, 1)
        fidelity_grid.addWidget(self._task_row(self.number_check), 1, 0)
        fidelity_grid.addWidget(self._task_row(self.url_check), 1, 1)
        for column in range(2):
            fidelity_grid.setColumnStretch(column, 1)
        layout.addLayout(fidelity_grid)

        layout.addWidget(horizontal_rule())
        target_heading = QLabel("Target 文本质量")
        target_heading.setObjectName("sectionTitle")
        layout.addWidget(target_heading)
        target_grid = QGridLayout()
        target_grid.addWidget(self._task_row(self.chinese_check), 0, 0)
        target_grid.addWidget(
            self._task_row(self.target_text_check, self.target_settings_button),
            0,
            1,
        )
        for column in range(2):
            target_grid.setColumnStretch(column, 1)
        layout.addLayout(target_grid)
        self.content_layout.addWidget(box)

        self.substring_settings_button.clicked.connect(
            lambda: self._open_settings_dialog("substring")
        )
        self.substring_consistency_check.toggled.connect(self.substring_settings_button.setEnabled)
        self.term_settings_button.clicked.connect(
            lambda: self._open_settings_dialog("term")
        )
        self.tag_settings_button.clicked.connect(
            lambda: self._open_settings_dialog("tag")
        )
        self.target_settings_button.clicked.connect(
            lambda: self._open_settings_dialog("target")
        )
        self.term_check.toggled.connect(
            lambda enabled: self.term_settings_button.setEnabled(enabled)
        )
        self.tag_check.toggled.connect(
            lambda enabled: self.tag_settings_button.setEnabled(enabled)
        )
        self.target_text_check.toggled.connect(
            lambda enabled: self.target_settings_button.setEnabled(enabled)
        )


    def _capture_options(self) -> dict[str, Any]:
        return {
            "input": {
                "path": self.input_picker.path(),
                "sheets": tuple(self.sheet.itemText(index) for index in range(self.sheet.count())),
                "sheet": self.sheet.currentText(),
                "source": self.source_column.text(),
                "target": self.target_column.text(),
                "start_row": self.start_row.value(),
            },
            "checks": tuple(check.isChecked() for check in self.task_checks),
            "settings": {
                name: self._capture_settings_state(name)
                for name in ("term", "tag", "target", "substring")
            },
        }

    def _restore_options(self, options: dict[str, Any]) -> None:
        scope = options["input"]
        self.input_picker.set_path(scope["path"])
        self._restore_combo(self.sheet, scope["sheets"], scope["sheet"])
        self.source_column.setText(scope["source"])
        self.target_column.setText(scope["target"])
        self.start_row.setValue(scope["start_row"])
        for check, checked in zip(self.task_checks, options["checks"], strict=True):
            check.setChecked(checked)
        for name, snapshot in options["settings"].items():
            self._restore_settings_state(name, snapshot)
        if scope["path"]:
            self.output_preview.setText(f"输出文件：{build_workflow_output_path(scope['path']).name}")

    def remember_options(self) -> None:
        try:
            self.options_store.save(self._capture_options())
        except OSError as exc:
            show_error(self, "记住选项失败", str(exc))
            return
        self.status.setText("已记住当前选项，下次打开时自动恢复。")

    def set_all_tasks(self, checked: bool) -> None:
        for checkbox in self.task_checks:
            checkbox.setChecked(checked)

    def choose_input_file(self) -> None:
        if path := _choose_excel(self, "选择 Excel 文件"):
            self.load_input_file(path)

    def _settings_parent(self, name: str) -> QWidget:
        dialogs = {
            "term": self.term_settings_dialog,
            "tag": self.tag_settings_dialog,
            "target": self.target_settings_dialog,
        }
        dialog = dialogs[name]
        return dialog if dialog.isVisible() else self

    def load_input_file(self, path: str, *, show_error: bool = True) -> None:
        self.input_picker.set_path(path)
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
            columns = detect_source_target_columns(
                self.input_picker.path(),
                sheet=self.sheet.currentText(),
                header_aliases=self.header_alias_store.load(),
            )
        except Exception as exc:  # noqa: BLE001
            if show_error:
                _show_error(self, "读取失败", str(exc))
            return
        self.source_column.setText(columns.detected_source_column or "")
        self.target_column.setText(columns.detected_target_column or "")

    def choose_history_file(self) -> None:
        if path := _choose_excel(
            self._settings_parent("term"),
            "选择术语历史 TB Excel 文件",
        ):
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
        path, _ = QFileDialog.getOpenFileName(
            self._settings_parent("tag"),
            "选择尖括号 Tag 过滤配置",
            "",
            "JSON 文件 (*.json);;所有文件 (*)",
        )
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
        parent = self._settings_parent("term")
        name, accepted = QInputDialog.getText(
            parent,
            "保存 TB 项目",
            "项目名称：",
            text=self.tb_project.currentText(),
        )
        name = name.strip()
        if not accepted:
            return
        if not name:
            show_error(parent, "项目名称为空", "请输入项目名称。")
            return
        try:
            project = self._capture_tb_project(name)
            existing = self.tb_store.find_project(name)
            if existing is not None and QMessageBox.question(
                parent,
                "更新 TB 项目",
                f"项目“{existing.name}”已存在，是否用当前设置更新？",
            ) != QMessageBox.StandardButton.Yes:
                return
            self.tb_store.save_project(project)
        except (OSError, ValueError) as exc:
            show_error(parent, "无法保存 TB 项目", str(exc))
            return
        self.refresh_tb_projects(name)

    def delete_tb_project(self) -> None:
        parent = self._settings_parent("term")
        name = self.tb_project.currentText().strip()
        if not name:
            show_info(parent, "未选择项目", "请先选择要删除的 TB 项目。")
            return
        if QMessageBox.question(
            parent,
            "删除 TB 项目",
            f"确定删除项目“{name}”吗？\n不会删除原始 TB 文件。",
        ) != QMessageBox.StandardButton.Yes:
            return
        try:
            self.tb_store.delete_project(name)
        except (OSError, ValueError) as exc:
            show_error(parent, "无法删除 TB 项目", str(exc))
            return
        self.refresh_tb_projects()

    def run_selected_tasks(self) -> None:
        input_file = self.input_picker.path()
        source_column = self.source_column.text().strip()
        target_column = self.target_column.text().strip()
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
        if self.term_check.isChecked() and not term_marks and not history_file:
            show_error(self, "缺少术语来源", "术语检查至少需要一种术语 mark，或一个历史 TB。")
            return
        if self.tag_check.isChecked() and not tag_types:
            show_error(self, "缺少检查类型", "Tag检查至少需要一种检查类型。")
            return
        if self.target_text_check.isChecked() and not target_rules:
            show_error(self, "缺少检查规则", "Target 文本规范检查至少需要选择一项规则。")
            return

        self._reset_workflow_report()
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
                "tag_check_order": self.tag_check_order.isChecked(),
                "tag_angle_config_file": self.angle_config.path() or None if "angle" in tag_types else None,
                "run_line_break_check": self.line_break_check.isChecked(),
                "run_source_consistency_check": self.consistency_check.isChecked(),
                "run_target_consistency_check": self.target_consistency_check.isChecked(),
                "run_substring_consistency_check": self.substring_consistency_check.isChecked(),
                "substring_min_cjk_chars": self.substring_min_cjk_chars.value(),
                "substring_min_other_chars": self.substring_min_other_chars.value(),
                "run_number_check": self.number_check.isChecked(),
                "run_url_check": self.url_check.isChecked(),
                "run_chinese_target_check": self.chinese_check.isChecked(),
                "run_target_text_check": self.target_text_check.isChecked(),
                "target_text_rules": target_rules,
            },
            on_success=lambda summary: self._finish_workflow(summary, history_file, input_file),
            on_error=self._fail_workflow,
        )

    def _finish_workflow(self, summary: object, history_file: str, input_file: str) -> None:
        self.run_button.setEnabled(True)
        self.revision_button.setEnabled(True)
        self.status.clear()
        if self.input_picker.path() == input_file:
            self.last_workflow_output_path = str(summary.output_path)
            self.open_report_button.setEnabled(True)
            self.open_report_button.setToolTip(self.last_workflow_output_path)
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
                f"同 Source 不同 Target 组数: {summary.source_consistency_problem_count}",
                f"同 Source 不同 Target 涉及行数: {summary.source_consistency_problem_rows}",
            ))
        if summary.ran_target_consistency_check:
            lines.extend((
                f"同 Target 不同 Source 组数: {summary.target_consistency_problem_count}",
                f"同 Target 不同 Source 涉及行数: {summary.target_consistency_problem_rows}",
            ))
        if summary.ran_substring_consistency_check:
            lines.append(f"子串译文一致性疑似问题行数: {summary.substring_consistency_problem_rows}")
        if summary.ran_number_check:
            lines.append(f"数字一致性问题行数: {summary.number_problem_rows}")
        if summary.ran_url_check:
            lines.append(f"URL 一致性问题行数: {summary.url_problem_rows}")
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

    def _reset_workflow_report(self) -> None:
        self.last_workflow_output_path = ""
        self.open_report_button.setEnabled(False)
        self.open_report_button.setToolTip("检查成功后可打开报告。")

    def open_report(self) -> None:
        if not self.last_workflow_output_path:
            return
        report_path = Path(self.last_workflow_output_path).expanduser()
        try:
            if not report_path.is_file():
                self._reset_workflow_report()
                show_warning(self, "报告不存在", f"报告可能已被移动或删除，请重新检查。\n{report_path}")
                return
            opened = QDesktopServices.openUrl(QUrl.fromLocalFile(str(report_path)))
        except OSError as exc:
            show_error(self, "无法打开报告", f"{exc}\n报告：{report_path}")
            return
        if not opened:
            show_error(
                self,
                "无法打开报告",
                f"请尝试从文件管理器打开，或检查 Excel / WPS 的默认打开设置。\n报告：{report_path}",
            )

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
            lines.append("因行号对应的 source 不匹配而跳过: " + "、".join(str(row) for row in summary.conflict_rows))
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
