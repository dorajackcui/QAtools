#!/usr/bin/env python3
"""Desktop UI for source/target line-break count checking."""

from __future__ import annotations

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from tools.excel_metadata import detect_source_target_columns, list_workbook_sheets
from tools.gui_common import (
    MUTED_LABEL_STYLE,
    PRIMARY_BUTTON_STYLE,
    OutputPreviewMixin,
    add_file_picker_row,
    configure_tool_page_style,
    create_section,
    parse_positive_int,
)

try:
    from .check_line_breaks import build_default_output_path, process_excel
except ImportError:
    from check_line_breaks import build_default_output_path, process_excel


class LineBreakCheckerApp(OutputPreviewMixin, ttk.Frame):
    output_path_builder = staticmethod(build_default_output_path)

    def __init__(self, master: tk.Misc) -> None:
        super().__init__(master, padding=16)
        self.input_file_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.source_column_var = tk.StringVar(value="A")
        self.target_column_var = tk.StringVar(value="B")
        self.start_row_var = tk.StringVar(value="2")
        self.output_preview_var = tk.StringVar(
            value="输出文件：选择输入 Excel 后自动生成"
        )
        self._build_ui()

    def _build_ui(self) -> None:
        configure_tool_page_style(self)
        input_frame = create_section(self, title="输入与范围", row=0)
        add_file_picker_row(
            input_frame,
            label="输入 Excel",
            variable=self.input_file_var,
            command=self.choose_input_file,
            focus_out_command=self.handle_input_file_focus_out,
        )

        scope_frame = ttk.Frame(input_frame)
        scope_frame.grid(row=1, column=0, columnspan=3, sticky="w", pady=(12, 0))
        ttk.Label(scope_frame, text="检查工作表").grid(row=0, column=0, sticky="w")
        self.sheet_combobox = ttk.Combobox(
            scope_frame, textvariable=self.sheet_var, width=18, state="readonly"
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

        ttk.Label(
            input_frame,
            text="逐行比较 source / target 单元格中的真实换行数量；CRLF 按一个换行计算。",
            style=MUTED_LABEL_STYLE,
            wraplength=760,
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(12, 0))

        ttk.Button(
            self,
            text="开始检查",
            command=self.run_check,
            style=PRIMARY_BUTTON_STYLE,
        ).grid(row=1, column=0, sticky="ew")
        ttk.Label(
            self,
            textvariable=self.output_preview_var,
            style=MUTED_LABEL_STYLE,
        ).grid(row=2, column=0, sticky="w", pady=(8, 0))
        self.columnconfigure(0, weight=1)

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
        self, _event: object | None = None, show_error: bool = True
    ) -> None:
        file_path = self.input_file_var.get().strip()
        sheet_name = self.sheet_var.get().strip() or None
        if not file_path or not sheet_name:
            return
        try:
            detected = detect_source_target_columns(file_path, sheet=sheet_name)
        except Exception as exc:
            if show_error:
                messagebox.showerror("读取失败", str(exc))
            return
        if detected.detected_source_column:
            self.source_column_var.set(detected.detected_source_column)
        if detected.detected_target_column:
            self.target_column_var.set(detected.detected_target_column)

    def run_check(self) -> None:
        input_file = self.input_file_var.get().strip()
        source_column = self.source_column_var.get().strip()
        target_column = self.target_column_var.get().strip()
        if not input_file:
            messagebox.showerror("缺少文件", "请先选择检查 Excel 文件。")
            return
        if not source_column or not target_column:
            messagebox.showerror("缺少列信息", "请填写 source 列和 target 列。")
            return
        try:
            start_row = parse_positive_int(
                self.start_row_var.get(), default=2, field_name="开始行"
            )
            summary = process_excel(
                input_file=input_file,
                source_column=source_column,
                target_column=target_column,
                sheet=self.sheet_var.get().strip() or None,
                start_row=start_row,
                output_file=None,
            )
        except Exception as exc:
            messagebox.showerror("处理失败", str(exc))
            return

        messagebox.showinfo(
            "处理完成",
            "\n".join(
                [
                    "换行数量检查已完成。",
                    f"检查工作表: {summary.worksheet_title}",
                    f"总行数: {summary.total_rows_checked}",
                    f"问题行数: {summary.problem_rows}",
                    f"输出文件: {summary.output_path}",
                ]
            ),
        )


def main() -> None:
    root = tk.Tk()
    root.title("Source / Target 换行数量检查")
    root.resizable(True, True)
    app = LineBreakCheckerApp(root)
    app.grid(row=0, column=0, sticky="nsew")
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    root.mainloop()


if __name__ == "__main__":
    main()
