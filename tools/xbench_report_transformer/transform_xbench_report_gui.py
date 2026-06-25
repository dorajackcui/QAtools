#!/usr/bin/env python3
"""Minimal desktop UI for transforming Xbench QA reports."""

from __future__ import annotations

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from tools.excel_metadata import list_workbook_sheets

try:
    from .transform_xbench_report import process_excel
except ImportError:
    from transform_xbench_report import process_excel


class XbenchReportTransformerApp(ttk.Frame):
    def __init__(self, master: tk.Misc) -> None:
        super().__init__(master, padding=16)
        self.input_file_var = tk.StringVar()
        self.sheet_var = tk.StringVar()

        self._build_ui()

    def _build_ui(self) -> None:
        ttk.Label(self, text="Xbench QA Report").grid(row=0, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.input_file_var, width=42).grid(
            row=0, column=1, sticky="ew", pady=(0, 8)
        )
        ttk.Button(self, text="选择", command=self.choose_input_file).grid(
            row=0, column=2, padx=(8, 0), pady=(0, 8)
        )

        ttk.Label(self, text="工作表名").grid(row=2, column=0, sticky="w", pady=(0, 8))
        self.sheet_combobox = ttk.Combobox(self, textvariable=self.sheet_var, width=20, state="readonly")
        self.sheet_combobox.grid(row=2, column=1, sticky="w", pady=(0, 8))

        ttk.Button(self, text="开始转换", command=self.run_transform).grid(
            row=3, column=0, columnspan=3, sticky="ew"
        )

        note = "规则：把 Xbench QA 明细整理为 文件名 / key / source / target / QA问题；默认输出到原文件同目录。"
        ttk.Label(self, text=note, wraplength=720).grid(
            row=4, column=0, columnspan=3, sticky="w", pady=(12, 0)
        )

        self.columnconfigure(1, weight=1)

    def choose_input_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="选择 Xbench QA Report Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if not file_path:
            return

        self.input_file_var.set(file_path)
        self.refresh_sheet_choices()

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
            selected_sheet = sheet_choices.default_sheet or (sheet_choices.sheet_names[0] if sheet_choices.sheet_names else "")
            self.sheet_var.set(selected_sheet)

    def run_transform(self) -> None:
        input_file = self.input_file_var.get().strip()
        sheet = self.sheet_var.get().strip() or None

        if not input_file:
            messagebox.showerror("缺少文件", "请先选择 Xbench QA Report Excel 文件。")
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
    root = tk.Tk()
    root.title("Xbench QA Report 转换")
    root.resizable(False, False)
    app = XbenchReportTransformerApp(root)
    app.grid(row=0, column=0, sticky="nsew")
    root.mainloop()


if __name__ == "__main__":
    main()
