#!/usr/bin/env python3
"""Desktop UI for transforming Xbench QA reports."""

from __future__ import annotations

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from tools.excel_metadata import list_workbook_sheets
from tools.gui_common import (
    MUTED_LABEL_STYLE,
    PRIMARY_BUTTON_STYLE,
    OutputPreviewMixin,
    add_optional_status_label,
    add_file_picker_row,
    configure_tool_page_style,
    create_application_root,
    create_section,
)

try:
    from .transform_xbench_report import build_default_output_path, process_excel
except ImportError:
    from transform_xbench_report import build_default_output_path, process_excel


class XbenchReportTransformerApp(OutputPreviewMixin, ttk.Frame):
    output_path_builder = staticmethod(build_default_output_path)

    def __init__(self, master: tk.Misc) -> None:
        super().__init__(master, padding=16)
        self.input_file_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.output_preview_var = tk.StringVar()
        self._build_ui()

    def _build_ui(self) -> None:
        configure_tool_page_style(self)
        input_frame = create_section(self, title="输入与范围", row=0)
        add_file_picker_row(
            input_frame,
            label="Xbench QA Report",
            variable=self.input_file_var,
            command=self.choose_input_file,
        )

        scope_frame = ttk.Frame(input_frame)
        scope_frame.grid(row=1, column=0, columnspan=3, sticky="w", pady=(12, 0))
        ttk.Label(scope_frame, text="报告工作表").grid(row=0, column=0, sticky="w")
        self.sheet_combobox = ttk.Combobox(
            scope_frame,
            textvariable=self.sheet_var,
            width=22,
            state="readonly",
        )
        self.sheet_combobox.grid(row=0, column=1, sticky="w", padx=(8, 0))
        ttk.Label(
            input_frame,
            text="将 QA 明细整理为文件名 / key / source / target / QA 问题，并按相同内容聚合。",
            style=MUTED_LABEL_STYLE,
            wraplength=760,
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(12, 0))

        ttk.Button(
            self,
            text="开始转换",
            command=self.run_transform,
            style=PRIMARY_BUTTON_STYLE,
        ).grid(row=1, column=0, sticky="ew")
        self.output_preview_label = add_optional_status_label(
            self,
            variable=self.output_preview_var,
            row=2,
        )
        self.columnconfigure(0, weight=1)

    def choose_input_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="选择 Xbench QA Report Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if not file_path:
            return
        self.input_file_var.set(file_path)
        self.refresh_sheet_choices()
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

    def run_transform(self) -> None:
        input_file = self.input_file_var.get().strip()
        sheet = self.sheet_var.get().strip() or None
        if not input_file:
            messagebox.showerror(
                "缺少文件", "请先选择 Xbench QA Report Excel 文件。"
            )
            return
        try:
            summary = process_excel(
                input_file=input_file,
                sheet=sheet,
                output_file=None,
            )
        except Exception as exc:
            messagebox.showerror("处理失败", str(exc))
            return

        messagebox.showinfo(
            "处理完成",
            "\n".join(
                [
                    "Xbench QA Report 转换已完成。",
                    f"工作表: {summary.worksheet_title}",
                    f"读取明细数: {summary.detail_count}",
                    f"输出行数: {summary.grouped_count}",
                    f"输出文件: {summary.output_path}",
                ]
            ),
        )


def main() -> None:
    root = create_application_root()
    root.title("Xbench QA Report 转换")
    root.resizable(True, True)
    app = XbenchReportTransformerApp(root)
    app.grid(row=0, column=0, sticky="nsew")
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    root.mainloop()


if __name__ == "__main__":
    main()
