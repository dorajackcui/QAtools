#!/usr/bin/env python3
"""GUI for running multiple checks against one Excel file."""

from __future__ import annotations

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from tools.excel_metadata import detect_source_target_columns, list_workbook_sheets
from tools.false_positive_review import review_clusters_with_codex
from tools.gui_common import parse_positive_int
from tools.term_pair_checker.extract_terms_from_excel import (
    TERM_SHEET_NAME,
    detect_history_tb_columns,
)

from .workflow_runner import run_workflow


class WorkflowRunnerApp(ttk.Frame):
    def __init__(self, master: tk.Misc) -> None:
        super().__init__(master, padding=16)
        self.input_file_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.term_history_tb_file_var = tk.StringVar()
        self.term_history_sheet_var = tk.StringVar()
        self.term_history_source_column_var = tk.StringVar()
        self.term_history_target_column_var = tk.StringVar()
        self.term_history_start_row_var = tk.StringVar(value="2")
        self.source_column_var = tk.StringVar(value="A")
        self.target_column_var = tk.StringVar(value="B")
        self.start_row_var = tk.StringVar(value="2")
        self.run_term_pair_var = tk.BooleanVar(value=True)
        self.run_tag_check_var = tk.BooleanVar(value=True)
        self.codex_fp_review_var = tk.BooleanVar(value=False)
        self.term_mark_style_vars = {
            "【】": tk.BooleanVar(value=False),
            "[]": tk.BooleanVar(value=True),
            "<>": tk.BooleanVar(value=True),
        }
        self.angle_var = tk.BooleanVar(value=True)
        self.square_color_var = tk.BooleanVar(value=True)
        self.brace_var = tk.BooleanVar(value=True)
        self.newline_var = tk.BooleanVar(value=True)
        self.numeric_var = tk.BooleanVar(value=True)

        self._build_ui()

    def _build_ui(self) -> None:
        ttk.Label(self, text="输入 Excel").grid(row=0, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.input_file_var, width=42).grid(
            row=0, column=1, sticky="ew", pady=(0, 8)
        )
        ttk.Button(self, text="选择", command=self.choose_input_file).grid(
            row=0, column=2, padx=(8, 0), pady=(0, 8)
        )

        ttk.Label(self, text="术语历史 TB Excel").grid(row=2, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.term_history_tb_file_var, width=42).grid(
            row=2, column=1, sticky="ew", pady=(0, 8)
        )
        history_buttons = ttk.Frame(self)
        history_buttons.grid(row=2, column=2, padx=(8, 0), pady=(0, 8), sticky="ew")
        ttk.Button(history_buttons, text="选择", command=self.choose_term_history_tb_file).grid(
            row=0, column=0, sticky="ew"
        )
        ttk.Button(history_buttons, text="清空", command=self.clear_term_history_tb_file).grid(
            row=0, column=1, sticky="ew", padx=(6, 0)
        )

        ttk.Label(self, text="术语历史 TB 工作表").grid(row=3, column=0, sticky="w", pady=(0, 8))
        self.term_history_sheet_combobox = ttk.Combobox(
            self,
            textvariable=self.term_history_sheet_var,
            width=20,
            state="readonly",
        )
        self.term_history_sheet_combobox.grid(row=3, column=1, sticky="w", pady=(0, 8))
        self.term_history_sheet_combobox.bind("<<ComboboxSelected>>", self.handle_term_history_sheet_selected)

        ttk.Label(self, text="术语历史 Source 列").grid(row=4, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.term_history_source_column_var, width=10).grid(
            row=4, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Label(self, text="术语历史 Target 列").grid(row=5, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.term_history_target_column_var, width=10).grid(
            row=5, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Label(self, text="术语历史开始行").grid(row=6, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.term_history_start_row_var, width=10).grid(
            row=6, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Label(self, text="检查工作表").grid(row=7, column=0, sticky="w", pady=(0, 8))
        self.sheet_combobox = ttk.Combobox(self, textvariable=self.sheet_var, width=20, state="readonly")
        self.sheet_combobox.grid(row=7, column=1, sticky="w", pady=(0, 8))
        self.sheet_combobox.bind("<<ComboboxSelected>>", self.handle_sheet_selected)

        ttk.Label(self, text="Source 列").grid(row=8, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.source_column_var, width=10).grid(
            row=8, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Label(self, text="Target 列").grid(row=9, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.target_column_var, width=10).grid(
            row=9, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Label(self, text="开始行").grid(row=10, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.start_row_var, width=10).grid(
            row=10, column=1, sticky="w", pady=(0, 8)
        )

        task_frame = ttk.LabelFrame(self, text="Workflow 任务", padding=10)
        task_frame.grid(row=11, column=0, columnspan=3, sticky="ew", pady=(4, 8))

        ttk.Checkbutton(task_frame, text="术语对检查", variable=self.run_term_pair_var).grid(
            row=0, column=0, sticky="w"
        )
        term_mark_frame = ttk.Frame(task_frame)
        term_mark_frame.grid(row=1, column=0, sticky="w", pady=(4, 8))
        ttk.Checkbutton(term_mark_frame, text="【】", variable=self.term_mark_style_vars["【】"]).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Checkbutton(term_mark_frame, text="[]", variable=self.term_mark_style_vars["[]"]).grid(
            row=0, column=1, sticky="w", padx=(12, 0)
        )
        ttk.Checkbutton(term_mark_frame, text="<>", variable=self.term_mark_style_vars["<>"]).grid(
            row=0, column=2, sticky="w", padx=(12, 0)
        )

        ttk.Checkbutton(task_frame, text="Tag检查", variable=self.run_tag_check_var).grid(
            row=2, column=0, sticky="w"
        )
        tag_type_frame = ttk.Frame(task_frame)
        tag_type_frame.grid(row=3, column=0, sticky="w", pady=(4, 0))
        ttk.Checkbutton(tag_type_frame, text="<...> tag", variable=self.angle_var).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Checkbutton(tag_type_frame, text="[color=...] tag", variable=self.square_color_var).grid(
            row=0, column=1, sticky="w", padx=(12, 0)
        )
        ttk.Checkbutton(tag_type_frame, text="{...} placeholder", variable=self.brace_var).grid(
            row=0, column=2, sticky="w", padx=(12, 0)
        )
        ttk.Checkbutton(tag_type_frame, text="\\n mark", variable=self.newline_var).grid(
            row=0, column=3, sticky="w", padx=(12, 0)
        )
        ttk.Checkbutton(tag_type_frame, text="数字tag", variable=self.numeric_var).grid(
            row=0, column=4, sticky="w", padx=(12, 0)
        )

        ttk.Checkbutton(
            self,
            text="使用 Codex 筛查术语误报",
            variable=self.codex_fp_review_var,
        ).grid(row=12, column=0, columnspan=3, sticky="w", pady=(0, 12))

        ttk.Button(self, text="开始执行 Workflow", command=self.run_selected_tasks).grid(
            row=13, column=0, columnspan=3, sticky="ew"
        )

        ttk.Label(
            self,
            text="说明：按顺序复用现有 checker，把术语对检查和 Tag检查结果写进同一份结果文件。",
        ).grid(row=14, column=0, columnspan=3, sticky="w", pady=(12, 0))

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

    def choose_term_history_tb_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="选择术语历史 TB Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if not file_path:
            return
        self.term_history_tb_file_var.set(file_path)
        self.refresh_term_history_sheet_choices()

    def clear_term_history_tb_file(self) -> None:
        self.term_history_tb_file_var.set("")
        self.term_history_source_column_var.set("")
        self.term_history_target_column_var.set("")
        self.term_history_start_row_var.set("2")
        self.clear_term_history_sheet_choices()

    def clear_term_history_sheet_choices(self) -> None:
        self.term_history_sheet_combobox["values"] = ()
        self.term_history_sheet_var.set("")

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

    def refresh_term_history_sheet_choices(self, show_error: bool = True) -> None:
        file_path = self.term_history_tb_file_var.get().strip()
        if not file_path:
            self.clear_term_history_sheet_choices()
            return

        try:
            sheet_choices = list_workbook_sheets(file_path)
        except Exception as exc:
            self.clear_term_history_sheet_choices()
            if show_error:
                messagebox.showerror("读取失败", str(exc))
            return

        self.term_history_sheet_combobox["values"] = sheet_choices.sheet_names
        selected_sheet = self.term_history_sheet_var.get().strip()
        if selected_sheet not in sheet_choices.sheet_names:
            selected_sheet = (
                TERM_SHEET_NAME
                if TERM_SHEET_NAME in sheet_choices.sheet_names
                else sheet_choices.default_sheet or (sheet_choices.sheet_names[0] if sheet_choices.sheet_names else "")
            )
            self.term_history_sheet_var.set(selected_sheet)

        self.handle_term_history_sheet_selected(show_error=show_error)

    def handle_sheet_selected(self, _event: object | None = None, show_error: bool = True) -> None:
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

    def handle_term_history_sheet_selected(self, _event: object | None = None, show_error: bool = True) -> None:
        file_path = self.term_history_tb_file_var.get().strip()
        sheet_name = self.term_history_sheet_var.get().strip() or None
        if not file_path or not sheet_name:
            return

        try:
            detected_columns = detect_history_tb_columns(file_path, sheet=sheet_name)
        except Exception as exc:
            if show_error:
                messagebox.showerror("读取失败", str(exc))
            return

        if detected_columns.source_column:
            self.term_history_source_column_var.set(detected_columns.source_column)
        if detected_columns.target_column:
            self.term_history_target_column_var.set(detected_columns.target_column)

    def get_selected_term_mark_styles(self) -> tuple[str, ...]:
        return tuple(
            mark_style
            for mark_style, variable in self.term_mark_style_vars.items()
            if variable.get()
        )

    def get_selected_tag_token_types(self) -> tuple[str, ...]:
        token_types: list[str] = []
        if self.angle_var.get():
            token_types.append("angle")
        if self.square_color_var.get():
            token_types.append("square_color")
        if self.brace_var.get():
            token_types.append("brace")
        if self.newline_var.get():
            token_types.append("newline")
        if self.numeric_var.get():
            token_types.append("numeric")
        return tuple(token_types)

    def run_selected_tasks(self) -> None:
        input_file = self.input_file_var.get().strip()
        term_history_tb_file = self.term_history_tb_file_var.get().strip()
        term_history_sheet = self.term_history_sheet_var.get().strip() or None
        term_history_source_column = self.term_history_source_column_var.get().strip() or None
        term_history_target_column = self.term_history_target_column_var.get().strip() or None
        source_column = self.source_column_var.get().strip()
        target_column = self.target_column_var.get().strip()
        run_term_pair_check = self.run_term_pair_var.get()
        run_tag_check = self.run_tag_check_var.get()
        term_mark_styles = self.get_selected_term_mark_styles()
        tag_token_types = self.get_selected_tag_token_types()
        term_history_start_row = 2

        if not input_file:
            messagebox.showerror("缺少文件", "请先选择输入 Excel 文件。")
            return
        if not source_column or not target_column:
            messagebox.showerror("缺少列信息", "请填写 source 列和 target 列。")
            return
        if run_term_pair_check and not term_mark_styles:
            messagebox.showerror("缺少 tag 类型", "术语对检查至少需要一种 tag 类型。")
            return
        if run_tag_check and not tag_token_types:
            messagebox.showerror("缺少检查类型", "Tag检查至少需要一种检查类型。")
            return

        try:
            start_row = parse_positive_int(self.start_row_var.get(), default=2, field_name="开始行")
        except ValueError as exc:
            messagebox.showerror("开始行错误", str(exc))
            return
        if term_history_tb_file:
            try:
                term_history_start_row = parse_positive_int(
                    self.term_history_start_row_var.get(),
                    default=2,
                    field_name="术语历史开始行",
                )
            except ValueError as exc:
                messagebox.showerror("术语历史开始行错误", str(exc))
                return
        else:
            term_history_sheet = None
            term_history_source_column = None
            term_history_target_column = None

        try:
            summary = run_workflow(
                input_file=input_file,
                output_file=None,
                source_column=source_column,
                target_column=target_column,
                sheet=self.sheet_var.get().strip() or None,
                start_row=start_row,
                run_term_pair_check=run_term_pair_check,
                term_mark_styles=term_mark_styles,
                term_history_tb_file=term_history_tb_file or None,
                term_history_sheet=term_history_sheet,
                term_history_source_column=term_history_source_column,
                term_history_target_column=term_history_target_column,
                term_history_start_row=term_history_start_row,
                run_tag_check=run_tag_check,
                tag_token_types=tag_token_types,
                false_positive_reviewer=review_clusters_with_codex if self.codex_fp_review_var.get() else None,
            )
        except Exception as exc:
            messagebox.showerror("处理失败", str(exc))
            return

        lines = [
            "Workflow 执行完成。",
            f"检查工作表: {summary.worksheet_title}",
            f"source 列: {summary.source_column}",
            f"target 列: {summary.target_column}",
        ]
        if summary.ran_term_pair_check:
            lines.append(f"术语表条目数: {summary.term_count}")
            lines.append(f"术语问题条数: {summary.term_problem_count}")
            if self.codex_fp_review_var.get():
                lines.append("Codex 假阳性筛查: 已写入 fp_* 辅助列")
            if term_history_tb_file:
                lines.append(f"术语历史 TB: {term_history_tb_file}")
        if summary.ran_tag_check:
            lines.append(f"Tag问题条数: {summary.tag_problem_count}")
        lines.append(f"输出文件: {summary.output_path}")
        messagebox.showinfo("处理完成", "\n".join(lines))


def main() -> None:
    root = tk.Tk()
    root.title("Excel Workflow 编排")
    root.resizable(False, False)
    app = WorkflowRunnerApp(root)
    app.grid(row=0, column=0, sticky="nsew")
    root.mainloop()


if __name__ == "__main__":
    main()
