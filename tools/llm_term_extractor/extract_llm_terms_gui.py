#!/usr/bin/env python3
"""Desktop UI for LLM-assisted term extraction from Excel workbooks."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from tools.excel_metadata import detect_source_target_columns, list_workbook_sheets
from tools.gui_common import parse_positive_int

try:
    from .extract_llm_terms import (
        DEFAULT_CODEX_MODEL,
        DEFAULT_CONFLICT_PROMPT,
        DEFAULT_EXTRACTION_PROMPT,
        DEFAULT_REASONING_EFFORT,
        default_prompt_path,
        detect_history_tb_columns,
        process_excel,
    )
except ImportError:
    from extract_llm_terms import (
        DEFAULT_CODEX_MODEL,
        DEFAULT_CONFLICT_PROMPT,
        DEFAULT_EXTRACTION_PROMPT,
        DEFAULT_REASONING_EFFORT,
        default_prompt_path,
        detect_history_tb_columns,
        process_excel,
    )


HISTORY_TERM_SHEET_NAME = "术语表"


class LlmTermExtractorApp(ttk.Frame):
    def __init__(self, master: tk.Misc) -> None:
        super().__init__(master, padding=16)
        self.input_file_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.source_column_var = tk.StringVar(value="A")
        self.target_column_var = tk.StringVar()
        self.start_row_var = tk.StringVar(value="2")
        self.batch_size_var = tk.StringVar(value="50")
        self.codex_model_var = tk.StringVar(value=DEFAULT_CODEX_MODEL)
        self.codex_reasoning_effort_var = tk.StringVar(value=DEFAULT_REASONING_EFFORT)
        self.extract_prompt_file_var = tk.StringVar(
            value=str(default_prompt_path(DEFAULT_EXTRACTION_PROMPT))
        )
        self.conflict_prompt_file_var = tk.StringVar(
            value=str(default_prompt_path(DEFAULT_CONFLICT_PROMPT))
        )
        self.history_tb_file_var = tk.StringVar()
        self.history_sheet_var = tk.StringVar()
        self.history_source_column_var = tk.StringVar()
        self.history_target_column_var = tk.StringVar()
        self.history_start_row_var = tk.StringVar(value="2")
        self.keep_raw_codex_output_var = tk.BooleanVar(value=False)

        self._build_ui()

    def _build_ui(self) -> None:
        ttk.Label(self, text="输入 Excel").grid(row=0, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.input_file_var, width=48).grid(
            row=0, column=1, sticky="ew", pady=(0, 8)
        )
        ttk.Button(self, text="选择", command=self.choose_input_file).grid(
            row=0, column=2, padx=(8, 0), pady=(0, 8)
        )

        ttk.Label(self, text="工作表名").grid(row=2, column=0, sticky="w", pady=(0, 8))
        self.sheet_combobox = ttk.Combobox(self, textvariable=self.sheet_var, width=20, state="readonly")
        self.sheet_combobox.grid(row=2, column=1, sticky="w", pady=(0, 8))
        self.sheet_combobox.bind("<<ComboboxSelected>>", self.handle_sheet_selected)

        ttk.Label(self, text="Source 列").grid(row=3, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.source_column_var, width=10).grid(
            row=3, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Label(self, text="Target 列（可选）").grid(row=4, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.target_column_var, width=10).grid(
            row=4, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Label(self, text="开始行").grid(row=5, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.start_row_var, width=10).grid(
            row=5, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Label(self, text="批大小").grid(row=6, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.batch_size_var, width=10).grid(
            row=6, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Label(self, text="Codex model").grid(row=7, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.codex_model_var, width=28).grid(
            row=7, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Label(self, text="Reasoning effort").grid(row=8, column=0, sticky="w", pady=(0, 8))
        ttk.Combobox(
            self,
            textvariable=self.codex_reasoning_effort_var,
            values=("low", "medium", "high", "xhigh"),
            width=12,
            state="readonly",
        ).grid(row=8, column=1, sticky="w", pady=(0, 8))

        ttk.Label(self, text="抽取 prompt").grid(row=9, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.extract_prompt_file_var, width=48).grid(
            row=9, column=1, sticky="ew", pady=(0, 8)
        )
        ttk.Button(self, text="选择", command=self.choose_extract_prompt_file).grid(
            row=9, column=2, padx=(8, 0), pady=(0, 8)
        )

        ttk.Label(self, text="冲突 prompt").grid(row=10, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.conflict_prompt_file_var, width=48).grid(
            row=10, column=1, sticky="ew", pady=(0, 8)
        )
        ttk.Button(self, text="选择", command=self.choose_conflict_prompt_file).grid(
            row=10, column=2, padx=(8, 0), pady=(0, 8)
        )

        ttk.Label(self, text="历史 TB Excel").grid(row=11, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.history_tb_file_var, width=48).grid(
            row=11, column=1, sticky="ew", pady=(0, 8)
        )
        history_buttons = ttk.Frame(self)
        history_buttons.grid(row=11, column=2, padx=(8, 0), pady=(0, 8), sticky="ew")
        ttk.Button(history_buttons, text="选择", command=self.choose_history_tb_file).grid(
            row=0, column=0, sticky="ew"
        )
        ttk.Button(history_buttons, text="清空", command=self.clear_history_tb_file).grid(
            row=0, column=1, sticky="ew", padx=(6, 0)
        )

        ttk.Label(self, text="历史 TB 工作表").grid(row=12, column=0, sticky="w", pady=(0, 8))
        self.history_sheet_combobox = ttk.Combobox(
            self,
            textvariable=self.history_sheet_var,
            width=20,
            state="readonly",
        )
        self.history_sheet_combobox.grid(row=12, column=1, sticky="w", pady=(0, 8))
        self.history_sheet_combobox.bind("<<ComboboxSelected>>", self.handle_history_sheet_selected)

        ttk.Label(self, text="历史 TB Source 列").grid(row=13, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.history_source_column_var, width=10).grid(
            row=13, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Label(self, text="历史 TB Target 列").grid(row=14, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.history_target_column_var, width=10).grid(
            row=14, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Label(self, text="历史 TB 开始行").grid(row=15, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(self, textvariable=self.history_start_row_var, width=10).grid(
            row=15, column=1, sticky="w", pady=(0, 8)
        )

        ttk.Checkbutton(
            self,
            text="保留 Codex 原始输出",
            variable=self.keep_raw_codex_output_var,
        ).grid(row=16, column=0, columnspan=3, sticky="w", pady=(0, 12))

        ttk.Button(self, text="开始 LLM 术语提取", command=self.run_extraction).grid(
            row=17, column=0, columnspan=3, sticky="ew"
        )

        ttk.Label(
            self,
            text="说明：target 列可留空；输出包含本批次术语汇总表和冲突汇总。",
        ).grid(row=18, column=0, columnspan=3, sticky="w", pady=(12, 0))

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

    def choose_extract_prompt_file(self) -> None:
        self._choose_prompt_file("选择抽取 prompt 文件", self.extract_prompt_file_var)

    def choose_conflict_prompt_file(self) -> None:
        self._choose_prompt_file("选择冲突复核 prompt 文件", self.conflict_prompt_file_var)

    def _choose_prompt_file(self, title: str, variable: tk.StringVar) -> None:
        current_path = variable.get().strip()
        initial_dir = str(Path(current_path).parent) if current_path else None
        file_path = filedialog.askopenfilename(
            title=title,
            filetypes=[("Markdown 文件", "*.md"), ("文本文件", "*.txt"), ("所有文件", "*.*")],
            initialdir=initial_dir,
        )
        if file_path:
            variable.set(file_path)

    def choose_history_tb_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="选择历史 TB Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if file_path:
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
        else:
            self.target_column_var.set("")

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
                HISTORY_TERM_SHEET_NAME
                if HISTORY_TERM_SHEET_NAME in sheet_choices.sheet_names
                else sheet_choices.default_sheet or (sheet_choices.sheet_names[0] if sheet_choices.sheet_names else "")
            )
            self.history_sheet_var.set(selected_sheet)

        self.handle_history_sheet_selected(show_error=show_error)

    def handle_history_sheet_selected(self, _event: object | None = None, show_error: bool = True) -> None:
        history_tb_file = self.history_tb_file_var.get().strip()
        selected_sheet = self.history_sheet_var.get().strip() or None
        if not history_tb_file or not selected_sheet:
            return

        history_start_row = self._parse_history_start_row(show_error=show_error)
        if history_start_row is None:
            return

        try:
            _sheet_title, source_column, target_column = detect_history_tb_columns(
                history_tb_file,
                sheet=selected_sheet,
                start_row=history_start_row,
            )
        except ValueError:
            self.history_source_column_var.set("")
            self.history_target_column_var.set("")
            return
        except Exception as exc:
            if show_error:
                messagebox.showerror("读取失败", str(exc))
            return

        self.history_source_column_var.set(source_column or "")
        self.history_target_column_var.set(target_column or "")

    def run_extraction(self) -> None:
        input_file = self.input_file_var.get().strip()
        sheet = self.sheet_var.get().strip() or None
        source_column = self.source_column_var.get().strip()
        target_column = self.target_column_var.get().strip() or None
        codex_model = self.codex_model_var.get().strip()
        codex_reasoning_effort = self.codex_reasoning_effort_var.get().strip()
        extract_prompt_file = self.extract_prompt_file_var.get().strip() or None
        conflict_prompt_file = self.conflict_prompt_file_var.get().strip() or None
        history_tb_file = self.history_tb_file_var.get().strip() or None
        history_sheet = self.history_sheet_var.get().strip() or None
        history_source_column = self.history_source_column_var.get().strip() or None
        history_target_column = self.history_target_column_var.get().strip() or None
        history_start_row = 2

        if not input_file:
            messagebox.showerror("缺少文件", "请先选择输入 Excel 文件。")
            return
        if not source_column:
            messagebox.showerror("缺少列信息", "请填写 source 列。")
            return
        if not codex_model or not codex_reasoning_effort:
            messagebox.showerror("缺少 Codex 设置", "请填写 Codex model 和 reasoning effort。")
            return

        try:
            start_row = parse_positive_int(self.start_row_var.get(), default=2, field_name="开始行")
        except ValueError as exc:
            messagebox.showerror("开始行错误", str(exc))
            return

        try:
            batch_size = parse_positive_int(self.batch_size_var.get(), default=50, field_name="批大小")
        except ValueError as exc:
            messagebox.showerror("批大小错误", str(exc))
            return

        if history_tb_file:
            parsed_history_start_row = self._parse_history_start_row()
            if parsed_history_start_row is None:
                return
            history_start_row = parsed_history_start_row
        else:
            history_sheet = None
            history_source_column = None
            history_target_column = None

        try:
            summary = process_excel(
                input_file=input_file,
                source_column=source_column,
                target_column=target_column,
                sheet=sheet,
                start_row=start_row,
                batch_size=batch_size,
                output_file=None,
                history_tb_file=history_tb_file,
                history_sheet=history_sheet,
                history_source_column=history_source_column,
                history_target_column=history_target_column,
                history_start_row=history_start_row,
                codex_model=codex_model,
                codex_reasoning_effort=codex_reasoning_effort,
                extract_prompt_file=extract_prompt_file,
                conflict_prompt_file=conflict_prompt_file,
                keep_raw_codex_output=self.keep_raw_codex_output_var.get(),
            )
        except Exception as exc:
            messagebox.showerror("处理失败", str(exc))
            return

        messagebox.showinfo(
            "处理完成",
            "\n".join(
                [
                    "LLM 术语提取已完成。",
                    f"工作表: {summary.worksheet_title}",
                    f"source 列: {summary.source_column}",
                    f"target 列: {summary.target_column or '未指定'}",
                    f"开始行: {summary.start_row}",
                    f"扫描行数: {summary.scanned_row_count}",
                    f"批次数: {summary.batch_count}",
                    f"术语数: {summary.term_count}",
                    f"冲突数: {summary.conflict_count}",
                    f"输出文件: {summary.output_path}",
                ]
            ),
        )

    def _parse_history_start_row(self, show_error: bool = True) -> int | None:
        try:
            return parse_positive_int(
                self.history_start_row_var.get(),
                default=2,
                field_name="历史 TB 开始行",
            )
        except ValueError as exc:
            if show_error:
                messagebox.showerror("历史 TB 开始行错误", str(exc))
            return None


def main() -> None:
    root = tk.Tk()
    root.title("LLM 术语提取")
    root.resizable(False, False)
    app = LlmTermExtractorApp(root)
    app.grid(row=0, column=0, sticky="nsew")
    root.mainloop()


if __name__ == "__main__":
    main()
