#!/usr/bin/env python3
"""Desktop UI for same-source target consistency checking."""

from __future__ import annotations

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from tools.excel_metadata import detect_source_target_columns, list_workbook_sheets
from tools.gui_common import parse_positive_int

try:
    from .check_source_consistency import process_excel
except ImportError:
    from check_source_consistency import process_excel


class SourceConsistencyCheckerApp(ttk.Frame):
    def __init__(self, master: tk.Misc) -> None:
        super().__init__(master, padding=16)
        self.input_file_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.source_column_var = tk.StringVar(value="A")
        self.target_column_var = tk.StringVar(value="B")
        self.start_row_var = tk.StringVar(value="2")
        self._build_ui()

    def _build_ui(self) -> None:
        ttk.Label(self, text="输入 Excel").grid(row=0, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.input_file_var, width=42).grid(
            row=0, column=1, sticky="ew", pady=(0, 8)
        )
        ttk.Button(self, text="选择", command=self.choose_input_file).grid(
            row=0, column=2, padx=(8, 0), pady=(0, 8)
        )

        ttk.Label(self, text="检查工作表").grid(row=1, column=0, sticky="w", pady=(0, 8))
        self.sheet_combobox = ttk.Combobox(
            self, textvariable=self.sheet_var, width=20, state="readonly"
        )
        self.sheet_combobox.grid(row=1, column=1, sticky="w", pady=(0, 8))
        self.sheet_combobox.bind("<<ComboboxSelected>>", self.handle_sheet_selected)

        ttk.Label(self, text="Source 列").grid(row=2, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.source_column_var, width=10).grid(
            row=2, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Label(self, text="Target 列").grid(row=3, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.target_column_var, width=10).grid(
            row=3, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Label(self, text="开始行").grid(row=4, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.start_row_var, width=10).grid(
            row=4, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Button(self, text="开始检查一致性", command=self.run_check).grid(
            row=5, column=0, columnspan=3, sticky="ew"
        )
        ttk.Label(
            self,
            text="规则：按 source 单元格文本精确分组，检查同一 source 是否对应多个不同 target。",
        ).grid(row=6, column=0, columnspan=3, sticky="w", pady=(12, 0))
        self.columnconfigure(1, weight=1)

    def choose_input_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="选择检查 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if not file_path:
            return
        self.input_file_var.set(file_path)
        self.refresh_sheet_choices()

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
                    "相同 source 译文一致性检查已完成。",
                    f"检查工作表: {summary.worksheet_title}",
                    f"重复 source 数: {summary.repeated_source_count}",
                    f"不一致 source 数: {summary.inconsistent_source_count}",
                    f"问题行数: {summary.problem_rows}",
                    f"输出文件: {summary.output_path}",
                ]
            ),
        )


def main() -> None:
    root = tk.Tk()
    root.title("相同 Source 译文一致性检查")
    root.resizable(False, False)
    app = SourceConsistencyCheckerApp(root)
    app.grid(row=0, column=0, sticky="nsew")
    root.mainloop()


if __name__ == "__main__":
    main()
