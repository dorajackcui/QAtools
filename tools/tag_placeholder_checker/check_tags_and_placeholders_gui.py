#!/usr/bin/env python3
"""Minimal desktop UI for tag and placeholder checking in Excel."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from tools.excel_metadata import detect_source_target_columns, list_workbook_sheets

try:
    from .check_tags_and_placeholders import (
        build_default_output_path,
        process_excel,
    )
except ImportError:
    from check_tags_and_placeholders import build_default_output_path, process_excel


class TagPlaceholderCheckerApp(ttk.Frame):
    def __init__(self, master: tk.Misc) -> None:
        super().__init__(master, padding=16)
        self.input_file_var = tk.StringVar()
        self.output_file_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.source_column_var = tk.StringVar(value="A")
        self.target_column_var = tk.StringVar(value="B")
        self.start_row_var = tk.StringVar(value="2")
        self.angle_var = tk.BooleanVar(value=True)
        self.brace_var = tk.BooleanVar(value=True)
        self.newline_var = tk.BooleanVar(value=True)

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

        ttk.Label(self, text="检查工作表").grid(row=2, column=0, sticky="w", pady=(0, 8))
        self.sheet_combobox = ttk.Combobox(
            self,
            textvariable=self.sheet_var,
            width=20,
            state="readonly",
        )
        self.sheet_combobox.grid(row=2, column=1, sticky="w", pady=(0, 8))
        self.sheet_combobox.bind("<<ComboboxSelected>>", self.handle_sheet_selected)

        ttk.Label(self, text="Source 列").grid(row=3, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.source_column_var, width=10).grid(
            row=3, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Label(self, text="Target 列").grid(row=4, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.target_column_var, width=10).grid(
            row=4, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Label(self, text="开始行").grid(row=5, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.start_row_var, width=10).grid(
            row=5, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Label(self, text="检查类型").grid(row=6, column=0, sticky="w", pady=(0, 8))
        token_type_frame = ttk.Frame(self)
        token_type_frame.grid(row=6, column=1, columnspan=2, sticky="w", pady=(0, 8))
        ttk.Checkbutton(token_type_frame, text="<...> tag", variable=self.angle_var).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Checkbutton(token_type_frame, text="{...} placeholder", variable=self.brace_var).grid(
            row=0, column=1, sticky="w", padx=(12, 0)
        )
        ttk.Checkbutton(token_type_frame, text="\\n mark", variable=self.newline_var).grid(
            row=0, column=2, sticky="w", padx=(12, 0)
        )

        ttk.Button(self, text="开始检查", command=self.run_check).grid(
            row=7, column=0, columnspan=3, sticky="ew"
        )

        note = r"规则：逐行比对 source / target 中的 <...>、{...} 和 \n，检查缺失、多出和数量不一致。"
        ttk.Label(self, text=note).grid(row=8, column=0, columnspan=3, sticky="w", pady=(12, 0))
        ttk.Label(
            self,
            text="说明：<...> 默认只检查配置文件里定义的 tag 模式，避免把普通尖括号内容误判成 tag。",
        ).grid(row=9, column=0, columnspan=3, sticky="w", pady=(4, 0))

        self.columnconfigure(1, weight=1)

    def choose_input_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="选择检查 Excel 文件",
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
            selected_sheet = sheet_choices.default_sheet or (sheet_choices.sheet_names[0] if sheet_choices.sheet_names else "")
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
        token_types: list[str] = []
        if self.angle_var.get():
            token_types.append("angle")
        if self.brace_var.get():
            token_types.append("brace")
        if self.newline_var.get():
            token_types.append("newline")
        return tuple(token_types)

    def run_check(self) -> None:
        input_file = self.input_file_var.get().strip()
        output_file = self.output_file_var.get().strip()
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
            start_row = int(self.start_row_var.get().strip() or "2")
        except ValueError:
            messagebox.showerror("开始行错误", "开始行必须是整数。")
            return

        try:
            summary = process_excel(
                input_file=input_file,
                sheet=self.sheet_var.get().strip() or None,
                source_column=source_column,
                target_column=target_column,
                start_row=start_row,
                token_types=token_types,
                output_file=output_file or None,
            )
        except Exception as exc:
            messagebox.showerror("处理失败", str(exc))
            return

        self.output_file_var.set(str(summary.output_path))
        selected_labels = []
        if "angle" in summary.selected_token_types:
            selected_labels.append("<...> tag")
        if "brace" in summary.selected_token_types:
            selected_labels.append("{...} placeholder")
        if "newline" in summary.selected_token_types:
            selected_labels.append(r"\n mark")

        messagebox.showinfo(
            "处理完成",
            "\n".join(
                [
                    "tag / placeholder 检查已完成。",
                    f"检查工作表: {summary.worksheet_title}",
                    f"检查类型: {'、'.join(selected_labels)}",
                    f"总行数: {summary.total_rows_checked}",
                    f"命中检查类型行数: {summary.rows_with_selected_tokens}",
                    f"含尖括号tag行数: {summary.angle_rows}",
                    f"含花括号placeholder行数: {summary.brace_rows}",
                    rf"含\n mark行数: {summary.newline_rows}",
                    f"问题行数: {summary.problem_rows}",
                    f"问题条数: {summary.problem_count}",
                    f"输出文件: {summary.output_path}",
                ]
            ),
        )


def main() -> None:
    root = tk.Tk()
    root.title("Excel Tag / Placeholder 检查")
    root.resizable(False, False)
    app = TagPlaceholderCheckerApp(root)
    app.grid(row=0, column=0, sticky="nsew")
    root.mainloop()


if __name__ == "__main__":
    main()
