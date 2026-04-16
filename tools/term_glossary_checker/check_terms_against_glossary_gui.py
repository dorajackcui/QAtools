#!/usr/bin/env python3
"""Minimal desktop UI for glossary-based term checking in Excel."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from tools.excel_metadata import detect_source_target_columns, list_workbook_sheets

try:
    from .check_terms_against_glossary import build_default_output_path, process_excel
except ImportError:
    from check_terms_against_glossary import build_default_output_path, process_excel


class TermGlossaryCheckerApp(ttk.Frame):
    def __init__(self, master: tk.Misc) -> None:
        super().__init__(master, padding=16)
        self.glossary_file_var = tk.StringVar()
        self.data_file_var = tk.StringVar()
        self.output_file_var = tk.StringVar()
        self.glossary_sheet_var = tk.StringVar()
        self.glossary_source_column_var = tk.StringVar(value="A")
        self.glossary_target_column_var = tk.StringVar(value="B")
        self.data_sheet_var = tk.StringVar()
        self.data_source_column_var = tk.StringVar(value="A")
        self.data_target_column_var = tk.StringVar(value="B")
        self.start_row_var = tk.StringVar(value="2")
        self.case_sensitive_var = tk.BooleanVar(value=False)
        self.match_mode_var = tk.StringVar(value="hybrid-boundary")

        self._build_ui()

    def _build_ui(self) -> None:
        ttk.Label(self, text="术语表 Excel").grid(row=0, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.glossary_file_var, width=42).grid(
            row=0, column=1, sticky="ew", pady=(0, 8)
        )
        ttk.Button(self, text="选择", command=self.choose_glossary_file).grid(
            row=0, column=2, padx=(8, 0), pady=(0, 8)
        )

        ttk.Label(self, text="检查文本 Excel").grid(row=1, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.data_file_var, width=42).grid(
            row=1, column=1, sticky="ew", pady=(0, 8)
        )
        ttk.Button(self, text="选择", command=self.choose_data_file).grid(
            row=1, column=2, padx=(8, 0), pady=(0, 8)
        )

        ttk.Label(self, text="输出 Excel").grid(row=2, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.output_file_var, width=42).grid(
            row=2, column=1, sticky="ew", pady=(0, 8)
        )
        ttk.Button(self, text="另存为", command=self.choose_output_file).grid(
            row=2, column=2, padx=(8, 0), pady=(0, 8)
        )

        ttk.Label(self, text="术语表工作表").grid(row=3, column=0, sticky="w", pady=(0, 8))
        self.glossary_sheet_combobox = ttk.Combobox(
            self,
            textvariable=self.glossary_sheet_var,
            width=20,
            state="readonly",
        )
        self.glossary_sheet_combobox.grid(row=3, column=1, sticky="w", pady=(0, 8))
        self.glossary_sheet_combobox.bind("<<ComboboxSelected>>", self.handle_glossary_sheet_selected)

        ttk.Label(self, text="术语表 Source 列").grid(row=4, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.glossary_source_column_var, width=10).grid(
            row=4, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Label(self, text="术语表 Target 列").grid(row=5, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.glossary_target_column_var, width=10).grid(
            row=5, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Label(self, text="检查工作表").grid(row=6, column=0, sticky="w", pady=(0, 8))
        self.data_sheet_combobox = ttk.Combobox(
            self,
            textvariable=self.data_sheet_var,
            width=20,
            state="readonly",
        )
        self.data_sheet_combobox.grid(row=6, column=1, sticky="w", pady=(0, 8))
        self.data_sheet_combobox.bind("<<ComboboxSelected>>", self.handle_data_sheet_selected)

        ttk.Label(self, text="检查 Source 列").grid(row=7, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.data_source_column_var, width=10).grid(
            row=7, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Label(self, text="检查 Target 列").grid(row=8, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.data_target_column_var, width=10).grid(
            row=8, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Label(self, text="开始行").grid(row=9, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.start_row_var, width=10).grid(
            row=9, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Label(self, text="匹配模式").grid(row=10, column=0, sticky="w", pady=(0, 8))
        ttk.Combobox(
            self,
            textvariable=self.match_mode_var,
            values=("hybrid-boundary", "substring"),
            width=20,
            state="readonly",
        ).grid(row=10, column=1, sticky="w", pady=(0, 8))

        ttk.Checkbutton(self, text="大小写敏感", variable=self.case_sensitive_var).grid(
            row=11, column=0, columnspan=3, sticky="w", pady=(0, 12)
        )

        ttk.Button(self, text="开始检查", command=self.run_check).grid(
            row=12, column=0, columnspan=3, sticky="ew"
        )

        note = "规则：默认使用混合边界匹配，避免 rain 命中 training、ACC 命中 account。"
        ttk.Label(self, text=note).grid(row=13, column=0, columnspan=3, sticky="w", pady=(12, 0))

        self.columnconfigure(1, weight=1)

    def choose_glossary_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="选择术语表 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if file_path:
            self.glossary_file_var.set(file_path)
            self.refresh_glossary_sheet_choices()

    def choose_data_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="选择检查文本 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if file_path:
            self.data_file_var.set(file_path)
            if not self.output_file_var.get().strip():
                self.output_file_var.set(str(build_default_output_path(Path(file_path))))
            self.refresh_data_sheet_choices()

    def choose_output_file(self) -> None:
        data_file = self.data_file_var.get().strip()
        initial_file = ""
        initial_dir = ""

        if data_file:
            data_path = Path(data_file)
            initial_dir = str(data_path.parent)
            initial_file = build_default_output_path(data_path).name

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

    def clear_glossary_sheet_choices(self) -> None:
        self.glossary_sheet_combobox["values"] = ()
        self.glossary_sheet_var.set("")

    def clear_data_sheet_choices(self) -> None:
        self.data_sheet_combobox["values"] = ()
        self.data_sheet_var.set("")

    def _refresh_sheet_choices(
        self,
        *,
        file_path: str,
        sheet_var: tk.StringVar,
        combobox: ttk.Combobox,
        on_selected,
        show_error: bool,
    ) -> None:
        if not file_path:
            combobox["values"] = ()
            sheet_var.set("")
            return

        try:
            sheet_choices = list_workbook_sheets(file_path)
        except Exception as exc:
            combobox["values"] = ()
            sheet_var.set("")
            if show_error:
                messagebox.showerror("读取失败", str(exc))
            return

        combobox["values"] = sheet_choices.sheet_names
        selected_sheet = sheet_var.get().strip()
        if selected_sheet not in sheet_choices.sheet_names:
            selected_sheet = sheet_choices.default_sheet or (sheet_choices.sheet_names[0] if sheet_choices.sheet_names else "")
            sheet_var.set(selected_sheet)

        on_selected(show_error=show_error)

    def _autofill_detected_columns(
        self,
        *,
        file_path: str,
        sheet_name: str | None,
        source_column_var: tk.StringVar,
        target_column_var: tk.StringVar,
        show_error: bool,
    ) -> None:
        if not file_path or not sheet_name:
            return

        try:
            detected_columns = detect_source_target_columns(file_path, sheet=sheet_name)
        except Exception as exc:
            if show_error:
                messagebox.showerror("读取失败", str(exc))
            return

        if detected_columns.detected_source_column:
            source_column_var.set(detected_columns.detected_source_column)
        if detected_columns.detected_target_column:
            target_column_var.set(detected_columns.detected_target_column)

    def refresh_glossary_sheet_choices(self, show_error: bool = True) -> None:
        self._refresh_sheet_choices(
            file_path=self.glossary_file_var.get().strip(),
            sheet_var=self.glossary_sheet_var,
            combobox=self.glossary_sheet_combobox,
            on_selected=self.handle_glossary_sheet_selected,
            show_error=show_error,
        )

    def refresh_data_sheet_choices(self, show_error: bool = True) -> None:
        self._refresh_sheet_choices(
            file_path=self.data_file_var.get().strip(),
            sheet_var=self.data_sheet_var,
            combobox=self.data_sheet_combobox,
            on_selected=self.handle_data_sheet_selected,
            show_error=show_error,
        )

    def handle_glossary_sheet_selected(
        self,
        _event: object | None = None,
        show_error: bool = True,
    ) -> None:
        self._autofill_detected_columns(
            file_path=self.glossary_file_var.get().strip(),
            sheet_name=self.glossary_sheet_var.get().strip() or None,
            source_column_var=self.glossary_source_column_var,
            target_column_var=self.glossary_target_column_var,
            show_error=show_error,
        )

    def handle_data_sheet_selected(
        self,
        _event: object | None = None,
        show_error: bool = True,
    ) -> None:
        self._autofill_detected_columns(
            file_path=self.data_file_var.get().strip(),
            sheet_name=self.data_sheet_var.get().strip() or None,
            source_column_var=self.data_source_column_var,
            target_column_var=self.data_target_column_var,
            show_error=show_error,
        )

    def run_check(self) -> None:
        glossary_file = self.glossary_file_var.get().strip()
        data_file = self.data_file_var.get().strip()
        output_file = self.output_file_var.get().strip()
        glossary_source_column = self.glossary_source_column_var.get().strip()
        glossary_target_column = self.glossary_target_column_var.get().strip()
        data_source_column = self.data_source_column_var.get().strip()
        data_target_column = self.data_target_column_var.get().strip()

        if not glossary_file:
            messagebox.showerror("缺少文件", "请先选择术语表 Excel 文件。")
            return
        if not data_file:
            messagebox.showerror("缺少文件", "请先选择检查文本 Excel 文件。")
            return
        if not glossary_source_column or not glossary_target_column:
            messagebox.showerror("缺少列信息", "请填写术语表 source 列和 target 列。")
            return
        if not data_source_column or not data_target_column:
            messagebox.showerror("缺少列信息", "请填写检查文本 source 列和 target 列。")
            return

        try:
            start_row = int(self.start_row_var.get().strip() or "2")
        except ValueError:
            messagebox.showerror("开始行错误", "开始行必须是整数。")
            return

        try:
            summary = process_excel(
                glossary_file=glossary_file,
                data_file=data_file,
                glossary_sheet=self.glossary_sheet_var.get().strip() or None,
                glossary_source_column=glossary_source_column,
                glossary_target_column=glossary_target_column,
                data_sheet=self.data_sheet_var.get().strip() or None,
                data_source_column=data_source_column,
                data_target_column=data_target_column,
                start_row=start_row,
                case_sensitive=self.case_sensitive_var.get(),
                match_mode=self.match_mode_var.get().strip() or "hybrid-boundary",
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
                    "术语检查已完成。",
                    f"术语表工作表: {summary.glossary_sheet_title}",
                    f"检查工作表: {summary.data_sheet_title}",
                    f"大小写模式: {'严格区分' if summary.case_sensitive else '忽略大小写'}",
                    f"匹配模式: {'混合边界' if summary.match_mode == 'hybrid-boundary' else '纯包含'}",
                    f"术语表条数: {summary.glossary_term_count}",
                    f"冲突术语数: {summary.conflict_count}",
                    f"命中术语行数: {summary.matched_rows}",
                    f"问题行数: {summary.problem_rows}",
                    f"问题条数: {summary.problem_count}",
                    f"输出文件: {summary.output_path}",
                ]
            ),
        )


def main() -> None:
    root = tk.Tk()
    root.title("Excel 术语表命中检查")
    root.resizable(False, False)
    app = TermGlossaryCheckerApp(root)
    app.grid(row=0, column=0, sticky="nsew")
    root.mainloop()


if __name__ == "__main__":
    main()
