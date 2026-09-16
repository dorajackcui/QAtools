"""QA settings dialogs, separated from page execution and file handling."""

from __future__ import annotations

from typing import Any
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QSpinBox,
    QVBoxLayout,
)
from tools.qt_gui_common import PathPicker, horizontal_rule
from tools.substring_consistency_checker.check_substring_consistency import (
    DEFAULT_MIN_CJK_CHARS, DEFAULT_MIN_OTHER_CHARS, MAX_MIN_CHARS,
)
from tools.target_text_checker.check_target_text import (
    ABNORMAL_PUNCTUATION_RULE,
    CONSECUTIVE_SPACES_RULE,
    LEADING_TRAILING_SPACES_RULE,
    MIXED_WIDTH_RULE,
    PAIRED_SYMBOLS_RULE,
)


class WorkflowSettingsMixin:
    """QA dialog construction and cancel/restore state on the owning page."""

    def _create_settings_dialog(
        self,
        name: str,
        title: str,
        *,
        minimum_width: int,
    ) -> tuple[QDialog, QVBoxLayout]:
        dialog = QDialog(self)
        dialog.setObjectName("settingsDialog")
        dialog.setWindowTitle(title)
        dialog.setModal(True)
        dialog.setMinimumWidth(minimum_width)
        dialog.setSizeGripEnabled(False)

        outer = QVBoxLayout(dialog)
        outer.setContentsMargins(16, 14, 16, 14)
        outer.setSpacing(12)
        content = QVBoxLayout()
        content.setContentsMargins(0, 0, 0, 0)
        content.setSpacing(9)
        outer.addLayout(content)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok
            | QDialogButtonBox.StandardButton.Cancel
        )
        ok_button = buttons.button(QDialogButtonBox.StandardButton.Ok)
        cancel_button = buttons.button(QDialogButtonBox.StandardButton.Cancel)
        ok_button.setText("确定")
        ok_button.setProperty("primary", True)
        cancel_button.setText("取消")
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        outer.addWidget(horizontal_rule())
        outer.addWidget(buttons)
        dialog.finished.connect(
            lambda result, dialog_name=name: self._settings_dialog_finished(
                dialog_name,
                result,
            )
        )
        return dialog, content


    def _build_term_settings(self) -> None:
        self.term_settings_dialog, layout = self._create_settings_dialog(
            "term",
            "术语检查设置",
            minimum_width=780,
        )
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
        self.refresh_tb_projects()


    def _build_tag_settings(self) -> None:
        self.tag_settings_dialog, layout = self._create_settings_dialog(
            "tag",
            "Tag / Placeholder 设置",
            minimum_width=620,
        )
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
        self.tag_check_order = QCheckBox("检查 Tag 顺序（类型、内容、数量逐一对应）")
        self.tag_check_order.setChecked(False)
        layout.addWidget(self.tag_check_order)
        self.angle_config = PathPicker("尖括号过滤配置", allow_clear=True)
        self.angle_config.choose_button.clicked.connect(self.choose_angle_config)
        layout.addWidget(self.angle_config)


    def _build_target_text_settings(self) -> None:
        self.target_settings_dialog, layout = self._create_settings_dialog(
            "target",
            "Target 文本规范设置",
            minimum_width=620,
        )
        layout.addWidget(QLabel("检查规则"))
        grid = QGridLayout()
        self.abnormal_rule = QCheckBox("异常标点符号（.. / ,, / 。。等）")
        self.spaces_rule = QCheckBox("连续空格（2 个及以上）")
        self.edge_spaces_rule = QCheckBox("首尾空格")
        self.width_rule = QCheckBox("全半角混用")
        self.paired_symbols_rule = QCheckBox("括号与引号配对")
        self.rule_checks = {
            ABNORMAL_PUNCTUATION_RULE: self.abnormal_rule,
            CONSECUTIVE_SPACES_RULE: self.spaces_rule,
            LEADING_TRAILING_SPACES_RULE: self.edge_spaces_rule,
            MIXED_WIDTH_RULE: self.width_rule,
            PAIRED_SYMBOLS_RULE: self.paired_symbols_rule,
        }
        for index, check in enumerate(self.rule_checks.values()):
            check.setChecked(True)
            grid.addWidget(check, index // 2, index % 2)
        for column in range(2):
            grid.setColumnStretch(column, 1)
        layout.addLayout(grid)


    def _build_substring_settings(self) -> None:
        self.substring_settings_dialog, layout = self._create_settings_dialog(
            "substring", "子串译文一致性设置", minimum_width=420,
        )
        grid = QGridLayout()
        self.substring_min_cjk_chars = QSpinBox()
        self.substring_min_other_chars = QSpinBox()
        for row, (label, widget, default) in enumerate((
            ("CJK 子串最小有效字符数", self.substring_min_cjk_chars, DEFAULT_MIN_CJK_CHARS),
            ("其他文字子串最小有效字符数", self.substring_min_other_chars, DEFAULT_MIN_OTHER_CHARS),
        )):
            widget.setRange(1, MAX_MIN_CHARS)
            widget.setValue(default)
            widget.setToolTip("统计文字和字母；空白、标点、数字、Tag 和占位符不计数。含 CJK 文字时使用 CJK 下限。")
            grid.addWidget(QLabel(label), row, 0)
            grid.addWidget(widget, row, 1)
        layout.addLayout(grid)

    def _open_settings_dialog(self, name: str) -> None:
        dialogs = {
            "term": self.term_settings_dialog,
            "tag": self.tag_settings_dialog,
            "target": self.target_settings_dialog,
            "substring": self.substring_settings_dialog,
        }
        self._settings_snapshots[name] = self._capture_settings_state(name)
        dialog = dialogs[name]
        dialog.adjustSize()
        dialog.open()
        dialog.raise_()
        dialog.activateWindow()


    def _capture_settings_state(self, name: str) -> dict[str, Any]:
        if name == "substring":
            return {
                "min_cjk_chars": self.substring_min_cjk_chars.value(),
                "min_other_chars": self.substring_min_other_chars.value(),
            }
        if name == "term":
            return {
                "mark_book": self.mark_book.isChecked(),
                "mark_square": self.mark_square.isChecked(),
                "tb_project": self.tb_project.currentText(),
                "history_path": self.history_picker.path(),
                "history_sheets": tuple(
                    self.history_sheet.itemText(index)
                    for index in range(self.history_sheet.count())
                ),
                "history_sheet": self.history_sheet.currentText(),
                "history_source": self.history_source.text(),
                "history_target": self.history_target.text(),
                "history_start_row": self.history_start_row.value(),
            }
        if name == "tag":
            return {
                "memoq_mode": self.memoq_mode.isChecked(),
                "tag_checks": tuple(
                    check.isChecked() for check in self.standard_tag_checks
                ),
                "angle_config": self.angle_config.path(),
                "check_order": self.tag_check_order.isChecked(),
            }
        if name == "target":
            return {
                "rules": {
                    rule: check.isChecked()
                    for rule, check in self.rule_checks.items()
                }
            }
        raise ValueError(f"未知设置窗口: {name}")


    @staticmethod
    def _restore_combo(
        combo: QComboBox,
        values: tuple[str, ...],
        selected: str,
    ) -> None:
        combo.blockSignals(True)
        combo.clear()
        combo.addItems(values)
        if selected in values:
            combo.setCurrentText(selected)
        elif not selected:
            combo.setCurrentIndex(-1)
        combo.blockSignals(False)


    def _restore_settings_state(self, name: str, snapshot: dict[str, Any]) -> None:
        if name == "substring":
            self.substring_min_cjk_chars.setValue(snapshot["min_cjk_chars"])
            self.substring_min_other_chars.setValue(snapshot["min_other_chars"])
            return
        if name == "term":
            self.mark_book.setChecked(snapshot["mark_book"])
            self.mark_square.setChecked(snapshot["mark_square"])
            self.history_picker.set_path(snapshot["history_path"])
            self._restore_combo(
                self.history_sheet,
                snapshot["history_sheets"],
                snapshot["history_sheet"],
            )
            self.history_source.setText(snapshot["history_source"])
            self.history_target.setText(snapshot["history_target"])
            self.history_start_row.setValue(snapshot["history_start_row"])
            self.refresh_tb_projects(snapshot["tb_project"])
            return
        if name == "tag":
            if snapshot["memoq_mode"]:
                self.memoq_mode.setChecked(True)
            else:
                self.standard_mode.setChecked(True)
            for check, checked in zip(
                self.standard_tag_checks,
                snapshot["tag_checks"],
                strict=True,
            ):
                check.setChecked(checked)
            self.angle_config.set_path(snapshot["angle_config"])
            self.tag_check_order.setChecked(snapshot["check_order"])
            self.update_tag_mode(not snapshot["memoq_mode"])
            return
        if name == "target":
            for rule, checked in snapshot["rules"].items():
                self.rule_checks[rule].setChecked(checked)


    def _settings_dialog_finished(self, name: str, result: int) -> None:
        snapshot = self._settings_snapshots.pop(name, None)
        if result == int(QDialog.DialogCode.Rejected) and snapshot is not None:
            self._restore_settings_state(name, snapshot)
