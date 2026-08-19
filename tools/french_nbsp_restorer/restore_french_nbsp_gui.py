#!/usr/bin/env python3
"""Desktop UI for restoring French non-breaking spaces in Excel."""

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
    create_section,
    parse_positive_int,
)

try:
    from .restore_french_nbsp import build_default_output_path, process_excel
except ImportError:
    from restore_french_nbsp import build_default_output_path, process_excel


class FrenchNbspRestorerApp(OutputPreviewMixin, ttk.Frame):
    output_path_builder = staticmethod(build_default_output_path)

    def __init__(self, master: tk.Misc) -> None:
        super().__init__(master, padding=16)
        self.input_file_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.target_column_var = tk.StringVar(value="B")
        self.result_column_var = tk.StringVar()
        self.start_row_var = tk.StringVar(value="2")
        self.output_preview_var = tk.StringVar()
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
        ttk.Label(scope_frame, text="处理工作表").grid(row=0, column=0, sticky="w")
        self.sheet_combobox = ttk.Combobox(
            scope_frame,
            textvariable=self.sheet_var,
            width=18,
            state="readonly",
        )
        self.sheet_combobox.grid(row=0, column=1, sticky="w", padx=(8, 18))
        self.sheet_combobox.bind("<<ComboboxSelected>>", self.handle_sheet_selected)
        ttk.Label(scope_frame, text="Target 列").grid(row=0, column=2, sticky="w")
        ttk.Entry(scope_frame, textvariable=self.target_column_var, width=7).grid(
            row=0, column=3, sticky="w", padx=(8, 18)
        )
        ttk.Label(scope_frame, text="开始行").grid(row=0, column=4, sticky="w")
        ttk.Spinbox(
            scope_frame,
            textvariable=self.start_row_var,
            width=7,
            from_=1,
            to=1_000_000,
        ).grid(row=0, column=5, sticky="w", padx=(8, 0))

        settings_frame = create_section(self, title="输出设置", row=1)
        ttk.Label(settings_frame, text="结果列（可选）").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Entry(settings_frame, textvariable=self.result_column_var, width=10).grid(
            row=0, column=1, sticky="w", padx=(12, 0)
        )
        ttk.Label(
            settings_frame,
            text="留空时直接修复 Target 列；恢复 ; : ? ! % 前及 « » 内侧的 NBSP。",
            style=MUTED_LABEL_STYLE,
            wraplength=760,
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(12, 0))

        ttk.Button(
            self,
            text="开始恢复",
            command=self.run_restore,
            style=PRIMARY_BUTTON_STYLE,
        ).grid(row=2, column=0, sticky="ew")
        self.output_preview_label = add_optional_status_label(
            self,
            variable=self.output_preview_var,
            row=3,
        )
        self.columnconfigure(0, weight=1)

    def choose_input_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if not file_path:
            return
        self.load_input_file(file_path)

    def load_input_file(
        self,
        file_path: str,
        *,
        reset_options: bool = False,
        show_error: bool = True,
    ) -> None:
        """Load an Excel path supplied by the picker or Finder quick action."""

        if reset_options:
            self.sheet_var.set("")
            self.target_column_var.set("B")
            self.result_column_var.set("")
            self.start_row_var.set("2")
        self.input_file_var.set(file_path)
        self.refresh_sheet_choices(show_error=show_error)
        self.update_output_preview()

    def clear_sheet_choices(self) -> None:
        self.sheet_combobox["values"] = ()
        self.sheet_var.set("")

    def refresh_sheet_choices(self, show_error: bool = True) -> None:
        input_file = self.input_file_var.get().strip()
        if not input_file:
            self.clear_sheet_choices()
            return
        try:
            sheet_choices = list_workbook_sheets(input_file)
        except Exception as exc:
            self.clear_sheet_choices()
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
        input_file = self.input_file_var.get().strip()
        sheet_name = self.sheet_var.get().strip() or None
        if not input_file or not sheet_name:
            return
        try:
            detected_columns = detect_source_target_columns(input_file, sheet=sheet_name)
        except Exception as exc:
            if show_error:
                messagebox.showerror("读取失败", str(exc))
            return
        if detected_columns.detected_target_column:
            self.target_column_var.set(detected_columns.detected_target_column)

    def run_restore(self) -> None:
        input_file = self.input_file_var.get().strip()
        sheet = self.sheet_var.get().strip() or None
        target_column = self.target_column_var.get().strip()
        result_column = self.result_column_var.get().strip() or None
        if not input_file:
            messagebox.showerror("缺少文件", "请先选择输入 Excel 文件。")
            return
        if not target_column:
            messagebox.showerror("缺少列信息", "请填写 target 列。")
            return

        try:
            start_row = parse_positive_int(
                self.start_row_var.get(), default=2, field_name="开始行"
            )
            summary = process_excel(
                input_file=input_file,
                target_column=target_column,
                result_column=result_column,
                sheet=sheet,
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
                    "法语 NBSP 恢复已完成。",
                    f"工作表: {summary.worksheet_title}",
                    f"处理行数: {summary.processed_count}",
                    f"修复行数: {summary.changed_count}",
                    f"输出文件: {summary.output_path}",
                ]
            ),
        )


def main() -> None:
    root = tk.Tk()
    root.title("法语 NBSP 恢复")
    root.resizable(True, True)
    app = FrenchNbspRestorerApp(root)
    app.grid(row=0, column=0, sticky="nsew")
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    root.mainloop()


if __name__ == "__main__":
    main()
