#!/usr/bin/env python3
"""GUI for running multiple checks against one Excel file."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from tools.excel_metadata import detect_source_target_columns, list_workbook_sheets
from tools.gui_common import (
    APP_MAIN_BACKGROUND,
    MUTED_LABEL_STYLE,
    PRIMARY_BUTTON_STYLE,
    SECTION_FRAME_STYLE,
    add_optional_status_label,
    configure_tool_page_style,
    create_application_root,
    create_file_path_display,
    parse_positive_int,
    refresh_top_aligned_scroll_region,
)
from tools.tb_project_ui import TbProjectControls
from tools.tb_projects import TbProject
from tools.target_text_checker.check_target_text import (
    ABNORMAL_PUNCTUATION_RULE,
    CONSECUTIVE_SPACES_RULE,
    MIXED_WIDTH_RULE,
)
from tools.term_pair_checker.extract_terms_from_excel import (
    TERM_SHEET_NAME,
    detect_history_tb_columns,
)

from .workflow_runner import build_default_output_path, run_workflow
from .revision_applier import (
    apply_workflow_revisions,
    build_default_revised_output_path,
)


class WorkflowRunnerApp(ttk.Frame):
    def __init__(self, master: tk.Misc) -> None:
        super().__init__(master, padding=16)
        self.input_file_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.term_history_tb_file_var = tk.StringVar()
        self.term_history_sheet_var = tk.StringVar()
        self.term_history_source_column_var = tk.StringVar()
        self.term_history_target_column_var = tk.StringVar()
        self.term_history_start_row_var = tk.StringVar(value="2")
        self.source_column_var = tk.StringVar(value="A")
        self.target_column_var = tk.StringVar(value="B")
        self.start_row_var = tk.StringVar(value="2")
        self.run_term_pair_var = tk.BooleanVar(value=True)
        self.run_tag_check_var = tk.BooleanVar(value=True)
        self.run_line_break_check_var = tk.BooleanVar(value=True)
        self.run_source_consistency_check_var = tk.BooleanVar(value=True)
        self.run_chinese_target_check_var = tk.BooleanVar(value=True)
        self.run_target_text_check_var = tk.BooleanVar(value=True)
        self.output_preview_var = tk.StringVar()
        self.term_settings_button_text_var = tk.StringVar(value="展开设置")
        self.tag_settings_button_text_var = tk.StringVar(value="展开设置")
        self.target_text_settings_button_text_var = tk.StringVar(value="展开设置")
        self.tag_mode_var = tk.StringVar(value="standard")
        self.tag_angle_config_file_var = tk.StringVar()
        self.term_settings_expanded = False
        self.tag_settings_expanded = False
        self.target_text_settings_expanded = False
        self.last_workflow_output_path = ""
        self.term_mark_style_vars = {
            "【】": tk.BooleanVar(value=True),
            "[]": tk.BooleanVar(value=True),
        }
        self.angle_var = tk.BooleanVar(value=True)
        self.square_color_var = tk.BooleanVar(value=True)
        self.brace_var = tk.BooleanVar(value=True)
        self.newline_var = tk.BooleanVar(value=True)
        self.target_text_rule_vars = {
            ABNORMAL_PUNCTUATION_RULE: tk.BooleanVar(value=True),
            CONSECUTIVE_SPACES_RULE: tk.BooleanVar(value=True),
            MIXED_WIDTH_RULE: tk.BooleanVar(value=True),
        }

        self._build_ui()

    def _build_ui(self) -> None:
        configure_tool_page_style(self)
        self.scroll_canvas = tk.Canvas(
            self,
            width=900,
            height=480,
            background=APP_MAIN_BACKGROUND,
            borderwidth=0,
            highlightthickness=0,
            yscrollincrement=16,
        )
        self.scroll_canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(
            self,
            orient="vertical",
            command=self.scroll_canvas.yview,
        )
        scrollbar.grid(row=0, column=1, sticky="ns", padx=(8, 0))
        self.scroll_canvas.configure(yscrollcommand=scrollbar.set)

        self.scroll_content = ttk.Frame(self.scroll_canvas)
        self.scroll_content.columnconfigure(0, weight=1)
        self.scroll_content_window = self.scroll_canvas.create_window(
            (0, 0),
            window=self.scroll_content,
            anchor="nw",
        )
        self.scroll_content.bind("<Configure>", self.handle_scroll_content_configure)
        self.scroll_canvas.bind("<Configure>", self.handle_scroll_canvas_configure)
        self.scroll_canvas.bind("<Enter>", self.bind_workflow_mousewheel)
        self.scroll_canvas.bind("<Leave>", self.unbind_workflow_mousewheel)

        input_frame = ttk.LabelFrame(
            self.scroll_content,
            text="输入与范围",
            padding=12,
            style=SECTION_FRAME_STYLE,
        )
        input_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        input_frame.columnconfigure(1, weight=0)

        ttk.Label(input_frame, text="输入 Excel").grid(row=0, column=0, sticky="w")
        self.input_file_display = create_file_path_display(
            input_frame,
            variable=self.input_file_var,
        )
        self.input_file_display.grid(row=0, column=1, sticky="w", padx=(12, 8))
        ttk.Button(input_frame, text="选择文件", command=self.choose_input_file).grid(
            row=0, column=2, sticky="ew"
        )

        scope_frame = ttk.Frame(input_frame)
        scope_frame.grid(row=1, column=0, columnspan=3, sticky="w", pady=(12, 0))
        ttk.Label(scope_frame, text="检查工作表").grid(row=0, column=0, sticky="w")
        self.sheet_combobox = ttk.Combobox(
            scope_frame,
            textvariable=self.sheet_var,
            width=18,
            state="readonly",
        )
        self.sheet_combobox.grid(row=0, column=1, sticky="w", padx=(8, 18))
        self.sheet_combobox.bind("<<ComboboxSelected>>", self.handle_sheet_selected)
        ttk.Label(scope_frame, text="Source 列").grid(row=0, column=2, sticky="w")
        ttk.Entry(scope_frame, textvariable=self.source_column_var, width=7).grid(
            row=0, column=3, sticky="w", padx=(8, 18)
        )
        ttk.Label(scope_frame, text="Target 列").grid(row=0, column=4, sticky="w")
        ttk.Entry(scope_frame, textvariable=self.target_column_var, width=7).grid(
            row=0, column=5, sticky="w", padx=(8, 18)
        )
        ttk.Label(scope_frame, text="开始行").grid(row=0, column=6, sticky="w")
        ttk.Spinbox(
            scope_frame,
            textvariable=self.start_row_var,
            width=7,
            from_=1,
            to=1_000_000,
        ).grid(row=0, column=7, sticky="w", padx=(8, 0))

        task_frame = ttk.LabelFrame(
            self.scroll_content,
            text="质量检查项目",
            padding=12,
            style=SECTION_FRAME_STYLE,
        )
        task_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        for column in range(3):
            task_frame.columnconfigure(column, weight=1, uniform="quality-checks")

        ttk.Label(task_frame, text="默认全部选中，可按需取消").grid(
            row=0, column=0, sticky="w"
        )
        selection_buttons = ttk.Frame(task_frame)
        selection_buttons.grid(row=0, column=2, sticky="e")
        ttk.Button(selection_buttons, text="全选", command=self.select_all_tasks).grid(
            row=0, column=0
        )
        ttk.Button(selection_buttons, text="取消全选", command=self.clear_all_tasks).grid(
            row=0, column=1, padx=(6, 0)
        )

        term_item = ttk.Frame(task_frame)
        term_item.grid(row=1, column=0, sticky="w", pady=(12, 4))
        ttk.Checkbutton(
            term_item,
            text="术语检查",
            variable=self.run_term_pair_var,
            command=self.handle_term_check_toggled,
        ).grid(row=0, column=0, sticky="w")
        self.term_settings_button = ttk.Button(
            term_item,
            textvariable=self.term_settings_button_text_var,
            command=self.toggle_term_settings,
            width=8,
        )
        self.term_settings_button.grid(row=0, column=1, padx=(8, 0))

        tag_item = ttk.Frame(task_frame)
        tag_item.grid(row=1, column=1, sticky="w", pady=(12, 4))
        ttk.Checkbutton(
            tag_item,
            text="Tag 检查",
            variable=self.run_tag_check_var,
            command=self.handle_tag_check_toggled,
        ).grid(row=0, column=0, sticky="w")
        self.tag_settings_button = ttk.Button(
            tag_item,
            textvariable=self.tag_settings_button_text_var,
            command=self.toggle_tag_settings,
            width=8,
        )
        self.tag_settings_button.grid(row=0, column=1, padx=(8, 0))

        ttk.Checkbutton(
            task_frame,
            text="换行数量检查",
            variable=self.run_line_break_check_var,
        ).grid(row=1, column=2, sticky="w", pady=(12, 4))
        ttk.Checkbutton(
            task_frame,
            text="同源译文一致性",
            variable=self.run_source_consistency_check_var,
        ).grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Checkbutton(
            task_frame,
            text="Target 中文检查",
            variable=self.run_chinese_target_check_var,
        ).grid(row=2, column=1, sticky="w", pady=(8, 0))

        target_text_item = ttk.Frame(task_frame)
        target_text_item.grid(row=2, column=2, sticky="w", pady=(8, 0))
        ttk.Checkbutton(
            target_text_item,
            text="Target 文本规范检查",
            variable=self.run_target_text_check_var,
            command=self.handle_target_text_check_toggled,
        ).grid(row=0, column=0, sticky="w")
        self.target_text_settings_button = ttk.Button(
            target_text_item,
            textvariable=self.target_text_settings_button_text_var,
            command=self.toggle_target_text_settings,
            width=8,
        )
        self.target_text_settings_button.grid(row=0, column=1, padx=(8, 0))

        self.term_settings_frame = ttk.LabelFrame(
            self.scroll_content,
            text="术语检查设置",
            padding=12,
            style=SECTION_FRAME_STYLE,
        )
        self.term_settings_frame.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        self.term_settings_frame.columnconfigure(1, weight=0)
        ttk.Label(self.term_settings_frame, text="术语标记").grid(
            row=0, column=0, sticky="w"
        )
        term_mark_frame = ttk.Frame(self.term_settings_frame)
        term_mark_frame.grid(row=0, column=1, columnspan=2, sticky="w", padx=(12, 0))
        ttk.Checkbutton(
            term_mark_frame,
            text="中文方括号【】",
            variable=self.term_mark_style_vars["【】"],
        ).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(
            term_mark_frame,
            text="半角方括号 []",
            variable=self.term_mark_style_vars["[]"],
        ).grid(row=0, column=1, sticky="w", padx=(16, 0))

        self.tb_project_controls = TbProjectControls(
            self.term_settings_frame,
            capture_project=self.capture_tb_project,
            apply_project=self.apply_tb_project,
        )
        self.tb_project_controls.grid(
            row=1,
            column=0,
            columnspan=3,
            sticky="ew",
            pady=(12, 0),
        )

        ttk.Label(self.term_settings_frame, text="历史 TB（可选）").grid(
            row=2, column=0, sticky="w", pady=(12, 0)
        )
        create_file_path_display(
            self.term_settings_frame,
            variable=self.term_history_tb_file_var,
        ).grid(row=2, column=1, sticky="w", padx=(12, 8), pady=(12, 0))
        history_buttons = ttk.Frame(self.term_settings_frame)
        history_buttons.grid(row=2, column=2, sticky="e", pady=(12, 0))
        ttk.Button(
            history_buttons,
            text="选择文件",
            command=self.choose_term_history_tb_file,
        ).grid(row=0, column=0)
        ttk.Button(
            history_buttons,
            text="清空",
            command=self.clear_term_history_tb_file,
        ).grid(row=0, column=1, padx=(6, 0))

        history_scope = ttk.Frame(self.term_settings_frame)
        history_scope.grid(row=3, column=0, columnspan=3, sticky="w", pady=(12, 0))
        ttk.Label(history_scope, text="工作表").grid(row=0, column=0, sticky="w")
        self.term_history_sheet_combobox = ttk.Combobox(
            history_scope,
            textvariable=self.term_history_sheet_var,
            width=16,
            state="readonly",
        )
        self.term_history_sheet_combobox.grid(row=0, column=1, sticky="w", padx=(8, 18))
        self.term_history_sheet_combobox.bind(
            "<<ComboboxSelected>>",
            self.handle_term_history_sheet_selected,
        )
        ttk.Label(history_scope, text="Source 列").grid(row=0, column=2, sticky="w")
        ttk.Entry(
            history_scope,
            textvariable=self.term_history_source_column_var,
            width=7,
        ).grid(row=0, column=3, sticky="w", padx=(8, 18))
        ttk.Label(history_scope, text="Target 列").grid(row=0, column=4, sticky="w")
        ttk.Entry(
            history_scope,
            textvariable=self.term_history_target_column_var,
            width=7,
        ).grid(row=0, column=5, sticky="w", padx=(8, 18))
        ttk.Label(history_scope, text="开始行").grid(row=0, column=6, sticky="w")
        ttk.Spinbox(
            history_scope,
            textvariable=self.term_history_start_row_var,
            width=7,
            from_=1,
            to=1_000_000,
        ).grid(row=0, column=7, sticky="w", padx=(8, 0))
        ttk.Label(
            self.term_settings_frame,
            text="未选择术语标记时，必须提供历史 TB。",
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(10, 0))
        self.term_settings_frame.grid_remove()

        self.tag_settings_frame = ttk.LabelFrame(
            self.scroll_content,
            text="Tag 检查设置",
            padding=12,
            style=SECTION_FRAME_STYLE,
        )
        self.tag_settings_frame.grid(row=3, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(self.tag_settings_frame, text="检查模式").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Radiobutton(
            self.tag_settings_frame,
            text="常规 Tag",
            variable=self.tag_mode_var,
            value="standard",
            command=self.handle_tag_mode_changed,
        ).grid(row=0, column=1, sticky="w", padx=(12, 0))
        ttk.Radiobutton(
            self.tag_settings_frame,
            text="memoQ Tag",
            variable=self.tag_mode_var,
            value="memoq",
            command=self.handle_tag_mode_changed,
        ).grid(row=0, column=2, sticky="w", padx=(16, 0))

        ttk.Label(self.tag_settings_frame, text="常规类型").grid(
            row=1, column=0, sticky="w", pady=(12, 0)
        )
        standard_tag_frame = ttk.Frame(self.tag_settings_frame)
        standard_tag_frame.grid(row=1, column=1, columnspan=3, sticky="w", padx=(12, 0), pady=(12, 0))
        self.standard_tag_checkbuttons = []
        for column, (label, variable) in enumerate(
            (
                ("<...> tag", self.angle_var),
                ("[color=...] tag", self.square_color_var),
                ("{...} placeholder", self.brace_var),
                (r"\n mark", self.newline_var),
            )
        ):
            checkbutton = ttk.Checkbutton(
                standard_tag_frame,
                text=label,
                variable=variable,
            )
            checkbutton.grid(row=0, column=column, sticky="w", padx=(0 if column == 0 else 16, 0))
            self.standard_tag_checkbuttons.append(checkbutton)

        ttk.Label(self.tag_settings_frame, text="尖括号过滤配置").grid(
            row=2, column=0, sticky="w", pady=(12, 0)
        )
        self.tag_angle_config_entry = create_file_path_display(
            self.tag_settings_frame,
            variable=self.tag_angle_config_file_var,
        )
        self.tag_angle_config_entry.grid(
            row=2,
            column=1,
            columnspan=2,
            sticky="w",
            padx=(12, 8),
            pady=(12, 0),
        )
        tag_config_actions = ttk.Frame(self.tag_settings_frame)
        tag_config_actions.grid(row=2, column=3, sticky="ew", pady=(12, 0))
        self.tag_angle_config_button = ttk.Button(
            tag_config_actions,
            text="选择文件",
            command=self.choose_tag_angle_config_file,
        )
        self.tag_angle_config_button.grid(row=0, column=0)
        self.tag_angle_config_clear_button = ttk.Button(
            tag_config_actions,
            text="清空",
            command=self.clear_tag_angle_config_file,
        )
        self.tag_angle_config_clear_button.grid(row=0, column=1, padx=(6, 0))
        self.tag_settings_frame.columnconfigure(2, weight=0)
        self.tag_settings_frame.grid_remove()

        self.target_text_settings_frame = ttk.LabelFrame(
            self.scroll_content,
            text="Target 文本规范检查设置",
            padding=12,
            style=SECTION_FRAME_STYLE,
        )
        self.target_text_settings_frame.grid(
            row=4,
            column=0,
            sticky="ew",
            pady=(0, 10),
        )
        ttk.Label(self.target_text_settings_frame, text="检查规则").grid(
            row=0,
            column=0,
            sticky="w",
        )
        target_text_rule_frame = ttk.Frame(self.target_text_settings_frame)
        target_text_rule_frame.grid(row=0, column=1, sticky="w", padx=(12, 0))
        for column, (rule, label) in enumerate(
            (
                (
                    ABNORMAL_PUNCTUATION_RULE,
                    "异常标点符号（.. / ,, / 。。等）",
                ),
                (CONSECUTIVE_SPACES_RULE, "连续空格（2 个及以上）"),
                (MIXED_WIDTH_RULE, "全半角混用"),
            )
        ):
            ttk.Checkbutton(
                target_text_rule_frame,
                text=label,
                variable=self.target_text_rule_vars[rule],
            ).grid(
                row=0,
                column=column,
                sticky="w",
                padx=(0 if column == 0 else 16, 0),
            )
        self.target_text_settings_frame.grid_remove()

        action_frame = ttk.Frame(self)
        action_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        action_frame.columnconfigure(0, weight=3)
        action_frame.columnconfigure(1, weight=1)
        ttk.Button(
            action_frame,
            text="开始检查",
            command=self.run_selected_tasks,
            style=PRIMARY_BUTTON_STYLE,
        ).grid(row=0, column=0, sticky="ew")
        ttk.Button(
            action_frame,
            text="应用修订",
            command=self.apply_revisions,
        ).grid(row=0, column=1, sticky="ew", padx=(8, 0))
        self.output_preview_label = add_optional_status_label(
            self,
            variable=self.output_preview_var,
            row=2,
            columnspan=2,
        )

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

    def handle_scroll_content_configure(self, _event: object | None = None) -> None:
        self.refresh_scroll_region()

    def handle_scroll_canvas_configure(self, event: object) -> None:
        width = getattr(event, "width", self.scroll_canvas.winfo_width())
        self.scroll_canvas.itemconfigure(self.scroll_content_window, width=width)
        self.refresh_scroll_region()

    def refresh_scroll_region(self) -> None:
        refresh_top_aligned_scroll_region(self.scroll_canvas)

    def bind_workflow_mousewheel(self, _event: object | None = None) -> None:
        self.bind_all("<MouseWheel>", self.handle_workflow_mousewheel)

    def unbind_workflow_mousewheel(self, _event: object | None = None) -> None:
        self.unbind_all("<MouseWheel>")

    def handle_workflow_mousewheel(self, event: object) -> None:
        delta = getattr(event, "delta", 0)
        if delta:
            self.scroll_canvas.yview_scroll(-1 if delta > 0 else 1, "units")

    def reveal_settings_frame(self, frame: ttk.LabelFrame) -> None:
        self.update_idletasks()
        self.refresh_scroll_region()
        content_height = max(self.scroll_content.winfo_reqheight(), 1)
        target_fraction = max(0.0, min(frame.winfo_y() / content_height, 1.0))
        self.scroll_canvas.yview_moveto(target_fraction)

    def task_vars(self) -> tuple[tk.BooleanVar, ...]:
        return (
            self.run_term_pair_var,
            self.run_tag_check_var,
            self.run_line_break_check_var,
            self.run_source_consistency_check_var,
            self.run_chinese_target_check_var,
            self.run_target_text_check_var,
        )

    def select_all_tasks(self) -> None:
        for variable in self.task_vars():
            variable.set(True)
        self.handle_term_check_toggled()
        self.handle_tag_check_toggled()
        self.handle_target_text_check_toggled()

    def clear_all_tasks(self) -> None:
        for variable in self.task_vars():
            variable.set(False)
        self.handle_term_check_toggled()
        self.handle_tag_check_toggled()
        self.handle_target_text_check_toggled()

    def set_term_settings_expanded(self, expanded: bool) -> None:
        self.term_settings_expanded = expanded
        self.term_settings_button_text_var.set("收起设置" if expanded else "展开设置")
        if expanded:
            self.term_settings_frame.grid()
            if hasattr(self, "scroll_canvas"):
                self.after_idle(
                    lambda: self.reveal_settings_frame(self.term_settings_frame)
                )
        else:
            self.term_settings_frame.grid_remove()

    def toggle_term_settings(self) -> None:
        if self.run_term_pair_var.get():
            expanded = not self.term_settings_expanded
            if expanded and self.tag_settings_expanded:
                self.set_tag_settings_expanded(False)
            if expanded and self.target_text_settings_expanded:
                self.set_target_text_settings_expanded(False)
            self.set_term_settings_expanded(expanded)

    def set_tag_settings_expanded(self, expanded: bool) -> None:
        self.tag_settings_expanded = expanded
        self.tag_settings_button_text_var.set("收起设置" if expanded else "展开设置")
        if expanded:
            self.tag_settings_frame.grid()
            if hasattr(self, "scroll_canvas"):
                self.after_idle(
                    lambda: self.reveal_settings_frame(self.tag_settings_frame)
                )
        else:
            self.tag_settings_frame.grid_remove()

    def toggle_tag_settings(self) -> None:
        if self.run_tag_check_var.get():
            expanded = not self.tag_settings_expanded
            if expanded and self.term_settings_expanded:
                self.set_term_settings_expanded(False)
            if expanded and self.target_text_settings_expanded:
                self.set_target_text_settings_expanded(False)
            self.set_tag_settings_expanded(expanded)

    def set_target_text_settings_expanded(self, expanded: bool) -> None:
        self.target_text_settings_expanded = expanded
        self.target_text_settings_button_text_var.set(
            "收起设置" if expanded else "展开设置"
        )
        if expanded:
            self.target_text_settings_frame.grid()
            if hasattr(self, "scroll_canvas"):
                self.after_idle(
                    lambda: self.reveal_settings_frame(
                        self.target_text_settings_frame
                    )
                )
        else:
            self.target_text_settings_frame.grid_remove()

    def toggle_target_text_settings(self) -> None:
        if self.run_target_text_check_var.get():
            expanded = not self.target_text_settings_expanded
            if expanded and self.term_settings_expanded:
                self.set_term_settings_expanded(False)
            if expanded and self.tag_settings_expanded:
                self.set_tag_settings_expanded(False)
            self.set_target_text_settings_expanded(expanded)

    def handle_term_check_toggled(self) -> None:
        enabled = self.run_term_pair_var.get()
        self.term_settings_button.configure(state="normal" if enabled else "disabled")
        if not enabled:
            self.set_term_settings_expanded(False)

    def handle_tag_check_toggled(self) -> None:
        enabled = self.run_tag_check_var.get()
        self.tag_settings_button.configure(state="normal" if enabled else "disabled")
        if not enabled:
            self.set_tag_settings_expanded(False)
        self.handle_tag_mode_changed()

    def handle_target_text_check_toggled(self) -> None:
        enabled = self.run_target_text_check_var.get()
        self.target_text_settings_button.configure(
            state="normal" if enabled else "disabled"
        )
        if not enabled:
            self.set_target_text_settings_expanded(False)

    def handle_tag_mode_changed(self) -> None:
        standard_enabled = (
            self.run_tag_check_var.get() and self.tag_mode_var.get() == "standard"
        )
        state = "normal" if standard_enabled else "disabled"
        for checkbutton in getattr(self, "standard_tag_checkbuttons", ()):
            checkbutton.configure(state=state)
        for widget_name in (
            "tag_angle_config_entry",
            "tag_angle_config_button",
            "tag_angle_config_clear_button",
        ):
            widget = getattr(self, widget_name, None)
            if widget is not None:
                widget.configure(state=state)

    def update_output_preview(self) -> None:
        input_file = self.input_file_var.get().strip()
        if not input_file:
            self.output_preview_var.set("")
            return
        output_name = build_default_output_path(input_file).name
        self.output_preview_var.set(f"输出文件：{output_name}")

    def handle_input_file_focus_out(self, _event: object | None = None) -> None:
        self.refresh_sheet_choices(show_error=False)

    def choose_input_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if not file_path:
            return

        self.load_input_file(file_path)

    def load_input_file(self, file_path: str, *, show_error: bool = True) -> None:
        """Load an Excel path supplied by the picker or Finder quick action."""

        self.input_file_var.set(file_path)
        self.last_workflow_output_path = ""
        self.refresh_sheet_choices(show_error=show_error)

    def choose_term_history_tb_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="选择术语历史 TB Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if not file_path:
            return
        if hasattr(self, "tb_project_controls"):
            self.tb_project_controls.mark_current_settings_modified()
        self.term_history_tb_file_var.set(file_path)
        self.refresh_term_history_sheet_choices()

    def choose_tag_angle_config_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="选择尖括号 Tag 过滤配置",
            filetypes=[("JSON 文件", "*.json"), ("所有文件", "*.*")],
        )
        if file_path:
            self.tag_angle_config_file_var.set(file_path)

    def clear_tag_angle_config_file(self) -> None:
        self.tag_angle_config_file_var.set("")

    def clear_term_history_tb_file(self) -> None:
        if hasattr(self, "tb_project_controls"):
            self.tb_project_controls.clear_selection()
        self.term_history_tb_file_var.set("")
        self.term_history_source_column_var.set("")
        self.term_history_target_column_var.set("")
        self.term_history_start_row_var.set("2")
        self.clear_term_history_sheet_choices()

    def capture_tb_project(self, project_name: str) -> TbProject:
        file_path = self.term_history_tb_file_var.get().strip()
        if not file_path:
            raise ValueError("请先选择历史 TB 文件。")
        if not Path(file_path).expanduser().is_file():
            raise ValueError(f"历史 TB 文件不存在：{file_path}")
        sheet = self.term_history_sheet_var.get().strip()
        source_column = self.term_history_source_column_var.get().strip()
        target_column = self.term_history_target_column_var.get().strip()
        if not sheet or not source_column or not target_column:
            raise ValueError("请先确认历史 TB 的工作表及 Source / Target 列。")
        start_row = parse_positive_int(
            self.term_history_start_row_var.get(),
            default=2,
            field_name="术语历史开始行",
        )
        return TbProject(
            name=project_name,
            file_path=str(Path(file_path).expanduser().absolute()),
            sheet=sheet,
            source_column=source_column,
            target_column=target_column,
            start_row=start_row,
        )

    def apply_tb_project(self, project: TbProject) -> None:
        self.term_history_tb_file_var.set(project.file_path)
        self.term_history_sheet_var.set(project.sheet)
        self.refresh_term_history_sheet_choices(show_error=False)
        self.term_history_sheet_var.set(project.sheet)
        self.term_history_source_column_var.set(project.source_column)
        self.term_history_target_column_var.set(project.target_column)
        self.term_history_start_row_var.set(str(project.start_row))

    def clear_term_history_sheet_choices(self) -> None:
        self.term_history_sheet_combobox["values"] = ()
        self.term_history_sheet_var.set("")

    def refresh_sheet_choices(self, show_error: bool = True) -> None:
        file_path = self.input_file_var.get().strip()
        self.update_output_preview()
        if not file_path:
            self.sheet_combobox["values"] = ()
            self.sheet_var.set("")
            return

        try:
            sheet_choices = list_workbook_sheets(file_path)
        except Exception as exc:
            self.sheet_combobox["values"] = ()
            self.sheet_var.set("")
            if show_error:
                messagebox.showerror("读取失败", str(exc))
            return

        self.sheet_combobox["values"] = sheet_choices.sheet_names
        selected_sheet = self.sheet_var.get().strip()
        if selected_sheet not in sheet_choices.sheet_names:
            selected_sheet = sheet_choices.default_sheet or (sheet_choices.sheet_names[0] if sheet_choices.sheet_names else "")
            self.sheet_var.set(selected_sheet)

        self.handle_sheet_selected(show_error=show_error)

    def refresh_term_history_sheet_choices(self, show_error: bool = True) -> None:
        file_path = self.term_history_tb_file_var.get().strip()
        if not file_path:
            self.clear_term_history_sheet_choices()
            return

        try:
            sheet_choices = list_workbook_sheets(file_path)
        except Exception as exc:
            self.clear_term_history_sheet_choices()
            if show_error:
                messagebox.showerror("读取失败", str(exc))
            return

        self.term_history_sheet_combobox["values"] = sheet_choices.sheet_names
        selected_sheet = self.term_history_sheet_var.get().strip()
        if selected_sheet not in sheet_choices.sheet_names:
            selected_sheet = (
                TERM_SHEET_NAME
                if TERM_SHEET_NAME in sheet_choices.sheet_names
                else sheet_choices.default_sheet or (sheet_choices.sheet_names[0] if sheet_choices.sheet_names else "")
            )
            self.term_history_sheet_var.set(selected_sheet)

        self.handle_term_history_sheet_selected(show_error=show_error)

    def handle_sheet_selected(self, _event: object | None = None, show_error: bool = True) -> None:
        file_path = self.input_file_var.get().strip()
        sheet_name = self.sheet_var.get().strip() or None
        if not file_path or not sheet_name:
            return

        try:
            detected_columns = detect_source_target_columns(file_path, sheet=sheet_name)
        except Exception as exc:
            if show_error:
                messagebox.showerror("读取失败", str(exc))
            return

        if detected_columns.detected_source_column:
            self.source_column_var.set(detected_columns.detected_source_column)
        if detected_columns.detected_target_column:
            self.target_column_var.set(detected_columns.detected_target_column)

    def handle_term_history_sheet_selected(self, _event: object | None = None, show_error: bool = True) -> None:
        file_path = self.term_history_tb_file_var.get().strip()
        sheet_name = self.term_history_sheet_var.get().strip() or None
        if not file_path or not sheet_name:
            return

        try:
            detected_columns = detect_history_tb_columns(file_path, sheet=sheet_name)
        except Exception as exc:
            if show_error:
                messagebox.showerror("读取失败", str(exc))
            return

        if detected_columns.source_column:
            self.term_history_source_column_var.set(detected_columns.source_column)
        if detected_columns.target_column:
            self.term_history_target_column_var.set(detected_columns.target_column)

    def get_selected_term_mark_styles(self) -> tuple[str, ...]:
        return tuple(
            mark_style
            for mark_style, variable in self.term_mark_style_vars.items()
            if variable.get()
        )

    def get_selected_tag_token_types(self) -> tuple[str, ...]:
        if self.tag_mode_var.get() == "memoq":
            return ("memoq",)

        token_types: list[str] = []
        if self.angle_var.get():
            token_types.append("angle")
        if self.square_color_var.get():
            token_types.append("square_color")
        if self.brace_var.get():
            token_types.append("brace")
        if self.newline_var.get():
            token_types.append("newline")
        return tuple(token_types)

    def get_selected_target_text_rules(self) -> tuple[str, ...]:
        return tuple(
            rule
            for rule, variable in self.target_text_rule_vars.items()
            if variable.get()
        )

    def apply_revisions(self) -> None:
        candidate_path = self.last_workflow_output_path or self.input_file_var.get().strip()
        candidate = Path(candidate_path).expanduser() if candidate_path else None
        report_file = filedialog.askopenfilename(
            title="选择已填写的问题处理 Excel",
            initialdir=str(candidate.parent) if candidate else None,
            initialfile=candidate.name if candidate else None,
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if not report_file:
            return

        try:
            default_output = build_default_revised_output_path(report_file)
        except Exception as exc:
            messagebox.showerror("读取失败", str(exc))
            return
        output_file = filedialog.asksaveasfilename(
            title="保存修订稿",
            initialdir=str(default_output.parent),
            initialfile=default_output.name,
            defaultextension=default_output.suffix or ".xlsx",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if not output_file:
            return

        try:
            summary = apply_workflow_revisions(report_file, output_file=output_file)
        except Exception as exc:
            messagebox.showerror("应用失败", str(exc))
            return

        lines = [
            "修订稿已生成。",
            f"回填修改: {summary.revised_count} 行",
            f"未填写（忽略）: {summary.ignored_count} 行",
            f"内容未变化: {summary.unchanged_count} 行",
        ]
        if summary.conflict_rows:
            lines.append(
                "因原 target 已变化而跳过: "
                + "、".join(str(row) for row in summary.conflict_rows)
            )
        lines.append(f"输出文件: {summary.output_path}")
        if summary.conflict_rows:
            messagebox.showwarning("修订稿已生成（存在冲突）", "\n".join(lines))
        else:
            messagebox.showinfo("修订稿已生成", "\n".join(lines))

    def run_selected_tasks(self) -> None:
        input_file = self.input_file_var.get().strip()
        term_history_tb_file = self.term_history_tb_file_var.get().strip()
        term_history_sheet = self.term_history_sheet_var.get().strip() or None
        term_history_source_column = self.term_history_source_column_var.get().strip() or None
        term_history_target_column = self.term_history_target_column_var.get().strip() or None
        source_column = self.source_column_var.get().strip()
        target_column = self.target_column_var.get().strip()
        run_term_pair_check = self.run_term_pair_var.get()
        run_tag_check = self.run_tag_check_var.get()
        run_line_break_check = self.run_line_break_check_var.get()
        run_source_consistency_check = self.run_source_consistency_check_var.get()
        run_chinese_target_check = self.run_chinese_target_check_var.get()
        run_target_text_check = self.run_target_text_check_var.get()
        term_mark_styles = self.get_selected_term_mark_styles()
        tag_token_types = self.get_selected_tag_token_types()
        target_text_rules = self.get_selected_target_text_rules()
        tag_angle_config_var = getattr(self, "tag_angle_config_file_var", None)
        tag_angle_config_file = (
            tag_angle_config_var.get().strip()
            if "angle" in tag_token_types and tag_angle_config_var is not None
            else ""
        )
        term_history_start_row = 2

        if not input_file:
            messagebox.showerror("缺少文件", "请先选择输入 Excel 文件。")
            return
        if not source_column or not target_column:
            messagebox.showerror("缺少列信息", "请填写 source 列和 target 列。")
            return
        if not any(
            (
                run_term_pair_check,
                run_tag_check,
                run_line_break_check,
                run_source_consistency_check,
                run_chinese_target_check,
                run_target_text_check,
            )
        ):
            messagebox.showerror("缺少任务", "请至少选择一个质量检查项目。")
            return
        if run_term_pair_check and not term_mark_styles and not term_history_tb_file:
            messagebox.showerror(
                "缺少术语来源",
                "术语检查至少需要一种术语 mark，或一个历史 TB。",
            )
            return
        if run_tag_check and not tag_token_types:
            messagebox.showerror("缺少检查类型", "Tag检查至少需要一种检查类型。")
            return
        if run_target_text_check and not target_text_rules:
            messagebox.showerror(
                "缺少检查规则",
                "Target 文本规范检查至少需要选择一项规则。",
            )
            return

        try:
            start_row = parse_positive_int(self.start_row_var.get(), default=2, field_name="开始行")
        except ValueError as exc:
            messagebox.showerror("开始行错误", str(exc))
            return
        if term_history_tb_file:
            try:
                term_history_start_row = parse_positive_int(
                    self.term_history_start_row_var.get(),
                    default=2,
                    field_name="术语历史开始行",
                )
            except ValueError as exc:
                messagebox.showerror("术语历史开始行错误", str(exc))
                return
        else:
            term_history_sheet = None
            term_history_source_column = None
            term_history_target_column = None

        try:
            summary = run_workflow(
                input_file=input_file,
                output_file=None,
                source_column=source_column,
                target_column=target_column,
                sheet=self.sheet_var.get().strip() or None,
                start_row=start_row,
                run_term_pair_check=run_term_pair_check,
                term_mark_styles=term_mark_styles,
                term_history_tb_file=term_history_tb_file or None,
                term_history_sheet=term_history_sheet,
                term_history_source_column=term_history_source_column,
                term_history_target_column=term_history_target_column,
                term_history_start_row=term_history_start_row,
                run_tag_check=run_tag_check,
                tag_token_types=tag_token_types,
                tag_angle_config_file=tag_angle_config_file or None,
                run_line_break_check=run_line_break_check,
                run_source_consistency_check=run_source_consistency_check,
                run_chinese_target_check=run_chinese_target_check,
                run_target_text_check=run_target_text_check,
                target_text_rules=target_text_rules,
            )
        except Exception as exc:
            messagebox.showerror("处理失败", str(exc))
            return

        self.last_workflow_output_path = str(summary.output_path)

        lines = [
            "一键质量检查完成。",
            f"检查工作表: {summary.worksheet_title}",
            f"source 列: {summary.source_column}",
            f"target 列: {summary.target_column}",
        ]
        if summary.ran_term_pair_check:
            lines.append(f"术语表条目数: {summary.term_count}")
            lines.append(f"术语问题行数: {summary.term_problem_rows}")
            if term_history_tb_file:
                lines.append(f"术语历史 TB: {term_history_tb_file}")
        if summary.ran_tag_check:
            lines.append(f"Tag问题行数: {summary.tag_problem_rows}")
        if summary.ran_line_break_check:
            lines.append(f"换行数量问题行数: {summary.line_break_problem_count}")
        if summary.ran_source_consistency_check:
            lines.append(
                f"同源译文不一致 source 数: {summary.source_consistency_problem_count}"
            )
            lines.append(
                f"同源译文不一致涉及行数: {summary.source_consistency_problem_rows}"
            )
        if summary.ran_chinese_target_check:
            lines.append(f"Target 中文问题行数: {summary.chinese_target_problem_count}")
        if summary.ran_target_text_check:
            lines.append(
                f"Target 文本规范问题行数: {summary.target_text_problem_rows}"
            )
        lines.append(f"输出文件: {summary.output_path}")
        messagebox.showinfo("处理完成", "\n".join(lines))


def main() -> None:
    root = create_application_root()
    root.title("一键质量检查")
    root.resizable(True, True)
    app = WorkflowRunnerApp(root)
    app.grid(row=0, column=0, sticky="nsew")
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    root.mainloop()


if __name__ == "__main__":
    main()
