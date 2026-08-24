#!/usr/bin/env python3
"""Desktop UI for tag and placeholder checking in Excel."""

from __future__ import annotations

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from tools.excel_metadata import detect_source_target_columns, list_workbook_sheets
from tools.gui_common import (
    MUTED_LABEL_STYLE,
    PRIMARY_BUTTON_STYLE,
    OutputPreviewMixin,
    add_optional_status_label,
    add_file_picker_row,
    configure_tool_page_style,
    create_application_root,
    create_file_path_display,
    create_section,
    parse_positive_int,
)

try:
    from .check_tags_and_placeholders import build_default_output_path, process_excel
except ImportError:
    from check_tags_and_placeholders import build_default_output_path, process_excel


class TagPlaceholderCheckerApp(OutputPreviewMixin, ttk.Frame):
    output_path_builder = staticmethod(build_default_output_path)

    def __init__(self, master: tk.Misc) -> None:
        super().__init__(master, padding=16)
        self.input_file_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.source_column_var = tk.StringVar(value="A")
        self.target_column_var = tk.StringVar(value="B")
        self.start_row_var = tk.StringVar(value="2")
        self.angle_var = tk.BooleanVar(value=True)
        self.square_color_var = tk.BooleanVar(value=True)
        self.brace_var = tk.BooleanVar(value=True)
        self.newline_var = tk.BooleanVar(value=True)
        self.tag_mode_var = tk.StringVar(value="standard")
        self.angle_config_file_var = tk.StringVar()
        self.output_preview_var = tk.StringVar()
        self.standard_tag_checkbuttons: list[ttk.Checkbutton] = []
        self._build_ui()

    def _build_ui(self) -> None:
        configure_tool_page_style(self)
        input_frame = create_section(self, title="输入与范围", row=0)
        add_file_picker_row(
            input_frame,
            label="输入 Excel",
            variable=self.input_file_var,
            command=self.choose_input_file,
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

        settings_frame = create_section(self, title="检查设置", row=1)
        ttk.Label(settings_frame, text="检查模式").grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(
            settings_frame,
            text="常规 Tag",
            variable=self.tag_mode_var,
            value="standard",
            command=self.handle_tag_mode_changed,
        ).grid(row=0, column=1, sticky="w", padx=(12, 0))
        ttk.Radiobutton(
            settings_frame,
            text="memoQ Marker",
            variable=self.tag_mode_var,
            value="memoq",
            command=self.handle_tag_mode_changed,
        ).grid(row=0, column=2, sticky="w", padx=(16, 0))

        ttk.Label(settings_frame, text="常规类型").grid(
            row=1, column=0, sticky="w", pady=(12, 0)
        )
        token_type_frame = ttk.Frame(settings_frame)
        token_type_frame.grid(
            row=1,
            column=1,
            columnspan=4,
            sticky="w",
            padx=(12, 0),
            pady=(12, 0),
        )
        for column, (label, variable) in enumerate(
            (
                ("<...> tag", self.angle_var),
                ("[color=...] tag", self.square_color_var),
                ("{...} placeholder", self.brace_var),
                (r"\n mark", self.newline_var),
            )
        ):
            checkbutton = ttk.Checkbutton(
                token_type_frame,
                text=label,
                variable=variable,
            )
            checkbutton.grid(
                row=0,
                column=column,
                sticky="w",
                padx=(0 if column == 0 else 16, 0),
            )
            self.standard_tag_checkbuttons.append(checkbutton)

        ttk.Label(settings_frame, text="尖括号过滤配置").grid(
            row=2, column=0, sticky="w", pady=(12, 0)
        )
        self.angle_config_entry = create_file_path_display(
            settings_frame,
            variable=self.angle_config_file_var,
        )
        self.angle_config_entry.grid(
            row=2,
            column=1,
            columnspan=3,
            sticky="w",
            padx=(12, 8),
            pady=(12, 0),
        )
        angle_config_actions = ttk.Frame(settings_frame)
        angle_config_actions.grid(row=2, column=4, sticky="ew", pady=(12, 0))
        self.angle_config_button = ttk.Button(
            angle_config_actions,
            text="选择文件",
            command=self.choose_angle_config_file,
        )
        self.angle_config_button.grid(row=0, column=0)
        self.angle_config_clear_button = ttk.Button(
            angle_config_actions,
            text="清空",
            command=self.clear_angle_config_file,
        )
        self.angle_config_clear_button.grid(row=0, column=1, padx=(6, 0))
        settings_frame.columnconfigure(3, weight=0)

        ttk.Label(
            settings_frame,
            text=(
                "逐行比较 source / target；“常规 Tag”检查所选常规类型，"
                "“memoQ Marker”只检查 memoQ 数字标记；尖括号过滤配置为可选 JSON。"
            ),
            style=MUTED_LABEL_STYLE,
            wraplength=760,
        ).grid(row=3, column=0, columnspan=5, sticky="w", pady=(12, 0))

        ttk.Button(
            self,
            text="开始检查",
            command=self.run_check,
            style=PRIMARY_BUTTON_STYLE,
        ).grid(row=2, column=0, sticky="ew")
        self.output_preview_label = add_optional_status_label(
            self,
            variable=self.output_preview_var,
            row=3,
        )
        self.columnconfigure(0, weight=1)

    def handle_tag_mode_changed(self) -> None:
        state = "normal" if self.tag_mode_var.get() == "standard" else "disabled"
        for checkbutton in self.standard_tag_checkbuttons:
            checkbutton.configure(state=state)
        for widget_name in (
            "angle_config_entry",
            "angle_config_button",
            "angle_config_clear_button",
        ):
            widget = getattr(self, widget_name, None)
            if widget is not None:
                widget.configure(state=state)

    def choose_angle_config_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="选择尖括号 Tag 过滤配置",
            filetypes=[("JSON 文件", "*.json"), ("所有文件", "*.*")],
        )
        if file_path:
            self.angle_config_file_var.set(file_path)

    def clear_angle_config_file(self) -> None:
        self.angle_config_file_var.set("")

    def choose_input_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="选择检查 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if not file_path:
            return
        self.input_file_var.set(file_path)
        self.refresh_sheet_choices()
        self.update_output_preview()

    def refresh_sheet_choices(self, show_error: bool = True) -> None:
        file_path = self.input_file_var.get().strip()
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
            selected_sheet = sheet_choices.default_sheet or (
                sheet_choices.sheet_names[0] if sheet_choices.sheet_names else ""
            )
            self.sheet_var.set(selected_sheet)
        self.handle_sheet_selected(show_error=show_error)

    def handle_sheet_selected(
        self,
        _event: object | None = None,
        show_error: bool = True,
    ) -> None:
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

    def get_selected_token_types(self) -> tuple[str, ...]:
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

    def run_check(self) -> None:
        input_file = self.input_file_var.get().strip()
        source_column = self.source_column_var.get().strip()
        target_column = self.target_column_var.get().strip()
        token_types = self.get_selected_token_types()

        if not input_file:
            messagebox.showerror("缺少文件", "请先选择检查 Excel 文件。")
            return
        if not source_column or not target_column:
            messagebox.showerror("缺少列信息", "请填写 source 列和 target 列。")
            return
        if not token_types:
            messagebox.showerror("缺少检查类型", "请至少选择一种检查类型。")
            return

        try:
            start_row = parse_positive_int(
                self.start_row_var.get(), default=2, field_name="开始行"
            )
            summary = process_excel(
                input_file=input_file,
                sheet=self.sheet_var.get().strip() or None,
                source_column=source_column,
                target_column=target_column,
                start_row=start_row,
                token_types=token_types,
                angle_config_file=(
                    self.angle_config_file_var.get().strip() or None
                    if "angle" in token_types
                    else None
                ),
                output_file=None,
            )
        except Exception as exc:
            messagebox.showerror("处理失败", str(exc))
            return

        labels = {
            "angle": "<...> tag",
            "square_color": "[color=...] tag",
            "brace": "{...} placeholder",
            "newline": r"\n mark",
            "memoq": "memoQ marker",
        }
        selected_labels = [labels[token] for token in summary.selected_token_types]
        messagebox.showinfo(
            "处理完成",
            "\n".join(
                [
                    "Tag 检查已完成。",
                    f"检查工作表: {summary.worksheet_title}",
                    f"检查类型: {'、'.join(selected_labels)}",
                    f"总行数: {summary.total_rows_checked}",
                    f"问题行数: {summary.problem_rows}",
                    f"问题条数: {summary.problem_count}",
                    f"输出文件: {summary.output_path}",
                ]
            ),
        )


def main() -> None:
    root = create_application_root()
    root.title("Excel Tag / Placeholder 检查")
    root.resizable(True, True)
    app = TagPlaceholderCheckerApp(root)
    app.grid(row=0, column=0, sticky="nsew")
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    root.mainloop()


if __name__ == "__main__":
    main()
