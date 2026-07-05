#!/usr/bin/env python3
"""Minimal desktop UI for validating source/target term pairs in Excel."""

from __future__ import annotations

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from tools.excel_metadata import detect_source_target_columns, list_workbook_sheets
from tools.false_positive_review import review_clusters_with_codex
from tools.gui_common import parse_positive_int

try:
    from .extract_terms_from_excel import (
        TERM_SHEET_NAME,
        detect_history_tb_columns,
        process_excel,
    )
except ImportError:
    from extract_terms_from_excel import (
        TERM_SHEET_NAME,
        detect_history_tb_columns,
        process_excel,
    )


class ExtractTermsApp(ttk.Frame):
    def __init__(self, master: tk.Misc) -> None:
        super().__init__(master, padding=16)
        self.input_file_var = tk.StringVar()
        self.history_tb_file_var = tk.StringVar()
        self.history_sheet_var = tk.StringVar()
        self.history_source_column_var = tk.StringVar()
        self.history_target_column_var = tk.StringVar()
        self.history_start_row_var = tk.StringVar(value="2")
        self.source_column_var = tk.StringVar(value="A")
        self.target_column_var = tk.StringVar(value="B")
        self.sheet_var = tk.StringVar()
        self.start_row_var = tk.StringVar(value="2")
        self.mark_style_vars = {
            "【】": tk.BooleanVar(value=True),
            "[]": tk.BooleanVar(value=True),
        }
        self.codex_fp_review_var = tk.BooleanVar(value=False)

        self._build_ui()

    def _build_ui(self) -> None:
        ttk.Label(self, text="输入 Excel").grid(row=0, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.input_file_var, width=42).grid(
            row=0, column=1, sticky="ew", pady=(0, 8)
        )
        ttk.Button(self, text="选择", command=self.choose_input_file).grid(
            row=0, column=2, padx=(8, 0), pady=(0, 8)
        )

        ttk.Label(self, text="历史 TB Excel").grid(row=2, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.history_tb_file_var, width=42).grid(
            row=2, column=1, sticky="ew", pady=(0, 8)
        )
        history_buttons = ttk.Frame(self)
        history_buttons.grid(row=2, column=2, padx=(8, 0), pady=(0, 8), sticky="ew")
        ttk.Button(history_buttons, text="选择", command=self.choose_history_tb_file).grid(
            row=0, column=0, sticky="ew"
        )
        ttk.Button(history_buttons, text="清空", command=self.clear_history_tb_file).grid(
            row=0, column=1, sticky="ew", padx=(6, 0)
        )

        ttk.Label(self, text="历史 TB 工作表").grid(row=3, column=0, sticky="w", pady=(0, 8))
        self.history_sheet_combobox = ttk.Combobox(
            self,
            textvariable=self.history_sheet_var,
            width=20,
            state="readonly",
        )
        self.history_sheet_combobox.grid(row=3, column=1, sticky="w", pady=(0, 8))
        self.history_sheet_combobox.bind("<<ComboboxSelected>>", self.handle_history_sheet_selected)

        ttk.Label(self, text="历史 TB Source 列").grid(row=4, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.history_source_column_var, width=10).grid(
            row=4, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Label(self, text="历史 TB Target 列").grid(row=5, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.history_target_column_var, width=10).grid(
            row=5, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Label(self, text="历史 TB 开始行").grid(row=6, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.history_start_row_var, width=10).grid(
            row=6, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Label(self, text="Source 列").grid(row=7, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.source_column_var, width=10).grid(
            row=7, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Label(self, text="Target 列").grid(row=8, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.target_column_var, width=10).grid(
            row=8, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Label(self, text="工作表名").grid(row=9, column=0, sticky="w", pady=(0, 8))
        self.sheet_combobox = ttk.Combobox(self, textvariable=self.sheet_var, width=20, state="readonly")
        self.sheet_combobox.grid(row=9, column=1, sticky="w", pady=(0, 8))
        self.sheet_combobox.bind("<<ComboboxSelected>>", self.handle_sheet_selected)

        ttk.Label(self, text="开始行").grid(row=10, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.start_row_var, width=10).grid(
            row=10, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Label(self, text="术语 mark").grid(row=11, column=0, sticky="nw", pady=(0, 12))
        mark_frame = ttk.Frame(self)
        mark_frame.grid(row=11, column=1, columnspan=2, sticky="w", pady=(0, 12))
        ttk.Checkbutton(mark_frame, text="【】", variable=self.mark_style_vars["【】"]).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Checkbutton(mark_frame, text="[]", variable=self.mark_style_vars["[]"]).grid(
            row=0, column=1, sticky="w", padx=(12, 0)
        )
        ttk.Checkbutton(
            self,
            text="使用 Codex 筛查术语误报",
            variable=self.codex_fp_review_var,
        ).grid(row=12, column=0, columnspan=3, sticky="w", pady=(0, 12))

        ttk.Button(self, text="开始检查", command=self.run_extraction).grid(
            row=13, column=0, columnspan=3, sticky="ew"
        )

        note = (
            "规则：术语表保留 tag，术语检查忽略 tag；"
            "历史 TB 命中时优先使用历史 target；"
            "<...> 和 {...} 不作为术语 mark。"
        )
        ttk.Label(self, text=note).grid(row=14, column=0, columnspan=3, sticky="w", pady=(12, 0))

        self.columnconfigure(1, weight=1)

    def choose_input_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if not file_path:
            return
        self.input_file_var.set(file_path)
        self.refresh_sheet_choices()

    def choose_history_tb_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="选择历史 TB Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if not file_path:
            return
        self.history_tb_file_var.set(file_path)
        self.refresh_history_sheet_choices()

    def clear_history_tb_file(self) -> None:
        self.history_tb_file_var.set("")
        self.history_source_column_var.set("")
        self.history_target_column_var.set("")
        self.history_start_row_var.set("2")
        self.clear_history_sheet_choices()

    def clear_sheet_choices(self) -> None:
        self.sheet_combobox["values"] = ()
        self.sheet_var.set("")

    def clear_history_sheet_choices(self) -> None:
        self.history_sheet_combobox["values"] = ()
        self.history_sheet_var.set("")

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

    def refresh_history_sheet_choices(self, show_error: bool = True) -> None:
        history_tb_file = self.history_tb_file_var.get().strip()
        if not history_tb_file:
            self.clear_history_sheet_choices()
            return

        try:
            sheet_choices = list_workbook_sheets(history_tb_file)
        except Exception as exc:
            self.clear_history_sheet_choices()
            if show_error:
                messagebox.showerror("读取失败", str(exc))
            return

        self.history_sheet_combobox["values"] = sheet_choices.sheet_names
        selected_sheet = self.history_sheet_var.get().strip()
        if selected_sheet not in sheet_choices.sheet_names:
            selected_sheet = (
                TERM_SHEET_NAME
                if TERM_SHEET_NAME in sheet_choices.sheet_names
                else sheet_choices.default_sheet or (sheet_choices.sheet_names[0] if sheet_choices.sheet_names else "")
            )
            self.history_sheet_var.set(selected_sheet)

        self.handle_history_sheet_selected(show_error=show_error)

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

    def handle_history_sheet_selected(self, _event: object | None = None, show_error: bool = True) -> None:
        history_tb_file = self.history_tb_file_var.get().strip()
        selected_sheet = self.history_sheet_var.get().strip() or None
        if not history_tb_file or not selected_sheet:
            return

        try:
            detected_columns = detect_history_tb_columns(history_tb_file, sheet=selected_sheet)
        except Exception as exc:
            if show_error:
                messagebox.showerror("读取失败", str(exc))
            return

        if detected_columns.source_column:
            self.history_source_column_var.set(detected_columns.source_column)
        if detected_columns.target_column:
            self.history_target_column_var.set(detected_columns.target_column)

    def run_extraction(self) -> None:
        input_file = self.input_file_var.get().strip()
        history_tb_file = self.history_tb_file_var.get().strip()
        history_sheet = self.history_sheet_var.get().strip() or None
        history_source_column = self.history_source_column_var.get().strip() or None
        history_target_column = self.history_target_column_var.get().strip() or None
        history_start_row = 2
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
            messagebox.showerror("缺少术语 mark", "请至少选择一种术语 mark。")
            return

        try:
            start_row = parse_positive_int(self.start_row_var.get(), default=2, field_name="开始行")
        except ValueError as exc:
            messagebox.showerror("开始行错误", str(exc))
            return
        if history_tb_file:
            try:
                history_start_row = parse_positive_int(
                    self.history_start_row_var.get(),
                    default=2,
                    field_name="历史 TB 开始行",
                )
            except ValueError as exc:
                messagebox.showerror("历史 TB 开始行错误", str(exc))
                return
        else:
            history_sheet = None
            history_source_column = None
            history_target_column = None

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
                history_tb_file=history_tb_file or None,
                history_sheet=history_sheet,
                history_source_column=history_source_column,
                history_target_column=history_target_column,
                history_start_row=history_start_row,
                output_file=None,
                false_positive_reviewer=review_clusters_with_codex if self.codex_fp_review_var.get() else None,
            )
        except Exception as exc:
            messagebox.showerror("处理失败", str(exc))
            return

        message_lines = [
            "术语检查已完成。",
            f"工作表: {worksheet_title}",
            f"source 列: {source_col}",
            f"target 列: {target_col}",
            f"术语 mark: {'、'.join(selected_mark_styles)}",
        ]
        if history_tb_file:
            message_lines.append(f"历史 TB: {history_tb_file}")
        if self.codex_fp_review_var.get():
            message_lines.append("Codex 假阳性筛查: 已写入 fp_* 辅助列")
        message_lines.extend(
            [
                f"术语表条目数: {term_count}",
                f"问题条数: {problem_count}",
                f"输出文件: {saved_path}",
            ]
        )
        messagebox.showinfo(
            "处理完成",
            "\n".join(message_lines),
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
