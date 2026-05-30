#!/usr/bin/env python3
"""Minimal desktop UI for checking Chinese characters in Excel target text."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from tools.excel_metadata import detect_source_target_columns, list_workbook_sheets

try:
    from .check_chinese_target import build_default_output_path, process_excel
except ImportError:
    from check_chinese_target import build_default_output_path, process_excel


class ChineseTargetCheckerApp(ttk.Frame):
    def __init__(self, master: tk.Misc) -> None:
        super().__init__(master, padding=16)
        self.input_file_var = tk.StringVar()
        self.output_file_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.target_column_var = tk.StringVar(value="B")
        self.result_column_var = tk.StringVar()
        self.start_row_var = tk.StringVar(value="2")
        self.problem_sheet_var = tk.BooleanVar(value=False)

        self._build_ui()

    def _build_ui(self) -> None:
        ttk.Label(self, text="输入 Excel").grid(row=0, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.input_file_var, width=42).grid(
            row=0, column=1, sticky="ew", pady=(0, 8)
        )
        ttk.Button(self, text="选择", command=self.choose_input_file).grid(
            row=0, column=2, padx=(8, 0), pady=(0, 8)
        )

        ttk.Label(self, text="输出 Excel").grid(row=1, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.output_file_var, width=42).grid(
            row=1, column=1, sticky="ew", pady=(0, 8)
        )
        ttk.Button(self, text="另存为", command=self.choose_output_file).grid(
            row=1, column=2, padx=(8, 0), pady=(0, 8)
        )

        ttk.Label(self, text="工作表名").grid(row=2, column=0, sticky="w", pady=(0, 8))
        self.sheet_combobox = ttk.Combobox(self, textvariable=self.sheet_var, width=20, state="readonly")
        self.sheet_combobox.grid(row=2, column=1, sticky="w", pady=(0, 8))
        self.sheet_combobox.bind("<<ComboboxSelected>>", self.handle_sheet_selected)

        ttk.Label(self, text="Target 列").grid(row=3, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.target_column_var, width=10).grid(
            row=3, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Label(self, text="结果列").grid(row=4, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.result_column_var, width=10).grid(
            row=4, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Label(self, text="开始行").grid(row=5, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.start_row_var, width=10).grid(
            row=5, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Checkbutton(
            self,
            text="新增问题工作表",
            variable=self.problem_sheet_var,
        ).grid(row=6, column=1, sticky="w", pady=(0, 12))

        ttk.Button(self, text="开始检查中文", command=self.run_check).grid(
            row=7, column=0, columnspan=3, sticky="ew"
        )

        note = "规则：检查 target 列是否包含中文字符；结果列为空时默认写入 target 右侧一列。"
        ttk.Label(self, text=note).grid(row=8, column=0, columnspan=3, sticky="w", pady=(12, 0))

        self.columnconfigure(1, weight=1)

    def choose_input_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if not file_path:
            return

        self.input_file_var.set(file_path)
        if not self.output_file_var.get().strip():
            self.output_file_var.set(str(build_default_output_path(Path(file_path))))
        self.refresh_sheet_choices()

    def choose_output_file(self) -> None:
        input_file = self.input_file_var.get().strip()
        initial_file = ""
        initial_dir = ""

        if input_file:
            input_path = Path(input_file)
            initial_dir = str(input_path.parent)
            initial_file = build_default_output_path(input_path).name

        file_path = filedialog.asksaveasfilename(
            title="选择输出 Excel 文件",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx"), ("Excel 启用宏", "*.xlsm"), ("所有文件", "*.*")],
            initialdir=initial_dir or None,
            initialfile=initial_file or None,
        )
        if file_path:
            self.output_file_var.set(file_path)

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

        self.handle_sheet_selected(show_error=show_error)

    def handle_sheet_selected(self, _event: object | None = None, show_error: bool = True) -> None:
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

    def run_check(self) -> None:
        input_file = self.input_file_var.get().strip()
        output_file = self.output_file_var.get().strip()
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
            start_row = int(self.start_row_var.get().strip() or "2")
        except ValueError:
            messagebox.showerror("开始行错误", "开始行必须是整数。")
            return

        try:
            summary = process_excel(
                input_file=input_file,
                target_column=target_column,
                result_column=result_column,
                sheet=sheet,
                start_row=start_row,
                create_problem_sheet=self.problem_sheet_var.get(),
                output_file=output_file or None,
            )
        except Exception as exc:
            messagebox.showerror("处理失败", str(exc))
            return

        self.output_file_var.set(str(summary.output_path))
        messagebox.showinfo(
            "处理完成",
            "\n".join(
                [
                    "target 中文检查已完成。",
                    f"工作表: {summary.worksheet_title}",
                    f"target 列: {summary.target_column}",
                    f"结果列: {summary.result_column}",
                    f"开始行: {summary.start_row}",
                    f"处理行数: {summary.processed_count}",
                    f"含中文行数: {summary.matched_count}",
                    f"问题工作表: {'已生成' if summary.problem_sheet_created else '未生成'}",
                    f"输出文件: {summary.output_path}",
                ]
            ),
        )


def main() -> None:
    root = tk.Tk()
    root.title("Target 中文检查")
    root.resizable(False, False)
    app = ChineseTargetCheckerApp(root)
    app.grid(row=0, column=0, sticky="nsew")
    root.mainloop()


if __name__ == "__main__":
    main()
