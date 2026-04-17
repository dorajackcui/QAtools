#!/usr/bin/env python3
"""Minimal desktop UI for validating source/target term pairs in Excel."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from tools.excel_metadata import detect_source_target_columns, list_workbook_sheets

try:
    from .extract_terms_from_excel import DEFAULT_EXCLUSION_CONFIG_NAME, build_default_output_path, process_excel
except ImportError:
    from extract_terms_from_excel import DEFAULT_EXCLUSION_CONFIG_NAME, build_default_output_path, process_excel


class ExtractTermsApp(ttk.Frame):
    def __init__(self, master: tk.Misc) -> None:
        super().__init__(master, padding=16)
        self.input_file_var = tk.StringVar()
        self.output_file_var = tk.StringVar()
        self.source_column_var = tk.StringVar(value="A")
        self.target_column_var = tk.StringVar(value="B")
        self.sheet_var = tk.StringVar()
        self.start_row_var = tk.StringVar(value="2")
        self.mark_style_vars = {
            "【】": tk.BooleanVar(value=True),
            "[]": tk.BooleanVar(value=False),
            "<>": tk.BooleanVar(value=False),
        }

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

        ttk.Label(self, text="Source 列").grid(row=2, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.source_column_var, width=10).grid(
            row=2, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Label(self, text="Target 列").grid(row=3, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.target_column_var, width=10).grid(
            row=3, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Label(self, text="工作表名").grid(row=4, column=0, sticky="w", pady=(0, 8))
        self.sheet_combobox = ttk.Combobox(self, textvariable=self.sheet_var, width=20, state="readonly")
        self.sheet_combobox.grid(row=4, column=1, sticky="w", pady=(0, 8))
        self.sheet_combobox.bind("<<ComboboxSelected>>", self.handle_sheet_selected)

        ttk.Label(self, text="开始行").grid(row=5, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.start_row_var, width=10).grid(
            row=5, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Label(self, text="Tag 类型").grid(row=6, column=0, sticky="nw", pady=(0, 12))
        mark_frame = ttk.Frame(self)
        mark_frame.grid(row=6, column=1, columnspan=2, sticky="w", pady=(0, 12))
        ttk.Checkbutton(mark_frame, text="【】", variable=self.mark_style_vars["【】"]).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Checkbutton(mark_frame, text="[]", variable=self.mark_style_vars["[]"]).grid(
            row=0, column=1, sticky="w", padx=(12, 0)
        )
        ttk.Checkbutton(mark_frame, text="<>", variable=self.mark_style_vars["<>"]).grid(
            row=0, column=2, sticky="w", padx=(12, 0)
        )

        ttk.Button(self, text="开始检查", command=self.run_extraction).grid(
            row=7, column=0, columnspan=3, sticky="ew"
        )

        note = (
            "规则：术语表保留 tag，术语检查忽略 tag；"
            f"伪标签排除规则请维护在 {DEFAULT_EXCLUSION_CONFIG_NAME}。"
        )
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
        if not file_path:
            return

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
        selected_sheet = self.sheet_var.get().strip() or None
        if not input_file or not selected_sheet:
            return

        try:
            detected_columns = detect_source_target_columns(input_file, sheet=selected_sheet)
        except Exception as exc:
            if show_error:
                messagebox.showerror("读取失败", str(exc))
            return

        if detected_columns.detected_source_column:
            self.source_column_var.set(detected_columns.detected_source_column)
        if detected_columns.detected_target_column:
            self.target_column_var.set(detected_columns.detected_target_column)

    def run_extraction(self) -> None:
        input_file = self.input_file_var.get().strip()
        output_file = self.output_file_var.get().strip()
        source_column = self.source_column_var.get().strip()
        target_column = self.target_column_var.get().strip()
        sheet = self.sheet_var.get().strip() or None
        selected_mark_styles = [
            mark_style for mark_style, variable in self.mark_style_vars.items() if variable.get()
        ]

        if not input_file:
            messagebox.showerror("缺少文件", "请先选择输入 Excel 文件。")
            return
        if not source_column or not target_column:
            messagebox.showerror("缺少列信息", "请填写 source 列和 target 列。")
            return
        if not selected_mark_styles:
            messagebox.showerror("缺少 tag 类型", "请至少选择一种 tag 类型。")
            return

        try:
            start_row = int(self.start_row_var.get().strip() or "2")
        except ValueError:
            messagebox.showerror("开始行错误", "开始行必须是整数。")
            return

        try:
            (
                worksheet_title,
                source_col,
                target_col,
                saved_path,
                term_count,
                problem_count,
            ) = process_excel(
                input_file=input_file,
                source_column=source_column,
                target_column=target_column,
                sheet=sheet,
                start_row=start_row,
                mark_styles=selected_mark_styles,
                output_file=output_file or None,
            )
        except Exception as exc:
            messagebox.showerror("处理失败", str(exc))
            return

        self.output_file_var.set(str(saved_path))
        messagebox.showinfo(
            "处理完成",
            "\n".join(
                [
                    "术语检查已完成。",
                    f"工作表: {worksheet_title}",
                    f"source 列: {source_col}",
                    f"target 列: {target_col}",
                    f"tag 类型: {'、'.join(selected_mark_styles)}",
                    f"术语表条目数: {term_count}",
                    f"问题行数: {problem_count}",
                    f"输出文件: {saved_path}",
                ]
            ),
        )


def main() -> None:
    root = tk.Tk()
    root.title("Excel 术语对校验")
    root.resizable(False, False)
    app = ExtractTermsApp(root)
    app.grid(row=0, column=0, sticky="nsew")
    root.mainloop()


if __name__ == "__main__":
    main()
