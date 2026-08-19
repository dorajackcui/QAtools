#!/usr/bin/env python3
"""Desktop UI for splitting and restoring Excel batches."""

from __future__ import annotations

from pathlib import Path
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from tools.excel_metadata import list_workbook_sheets
from tools.gui_common import (
    PRIMARY_BUTTON_STYLE,
    add_file_picker_row,
    add_optional_status_label,
    configure_tool_page_style,
    create_application_root,
    create_file_path_display,
    create_section,
    parse_positive_int,
)

try:
    from .excel_batcher import (
        build_default_output_dir,
        build_default_restore_path,
        restore_batches,
        split_workbook,
    )
except ImportError:
    from excel_batcher import (
        build_default_output_dir,
        build_default_restore_path,
        restore_batches,
        split_workbook,
    )


EXCEL_FILE_TYPES = (
    ("Excel 文件", "*.xlsx *.xlsm"),
    ("所有文件", "*.*"),
)


class ExcelBatcherApp(ttk.Frame):
    def __init__(self, master: tk.Misc) -> None:
        super().__init__(master, padding=16)
        self.split_input_file_var = tk.StringVar()
        self.split_sheet_var = tk.StringVar()
        self.batch_size_var = tk.StringVar(value="1000")
        self.header_rows_var = tk.StringVar(value="1")
        self.split_output_dir_var = tk.StringVar()
        self.split_output_preview_var = tk.StringVar()
        self.split_status_var = tk.StringVar()
        self.restore_batch_dir_var = tk.StringVar()
        self.restore_output_file_var = tk.StringVar()
        self.restore_status_var = tk.StringVar()
        self._build_ui()

    def _build_ui(self) -> None:
        configure_tool_page_style(self)
        notebook = ttk.Notebook(self)
        notebook.grid(row=0, column=0, sticky="nsew")

        split_page = ttk.Frame(notebook, padding=(4, 14, 4, 4))
        restore_page = ttk.Frame(notebook, padding=(4, 14, 4, 4))
        notebook.add(split_page, text="拆分 batch")
        notebook.add(restore_page, text="复原文件")
        self._build_split_page(split_page)
        self._build_restore_page(restore_page)

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

    def _build_split_page(self, parent: ttk.Frame) -> None:
        input_frame = create_section(parent, title="输入与范围", row=0)
        add_file_picker_row(
            input_frame,
            label="输入 Excel",
            variable=self.split_input_file_var,
            command=self.choose_split_input_file,
        )

        scope_frame = ttk.Frame(input_frame)
        scope_frame.grid(row=1, column=0, columnspan=3, sticky="w", pady=(12, 0))
        ttk.Label(scope_frame, text="工作表").grid(row=0, column=0, sticky="w")
        self.split_sheet_combobox = ttk.Combobox(
            scope_frame,
            textvariable=self.split_sheet_var,
            width=22,
            state="readonly",
        )
        self.split_sheet_combobox.grid(row=0, column=1, sticky="w", padx=(8, 24))

        ttk.Label(scope_frame, text="每个 batch 数据行数").grid(
            row=0,
            column=2,
            sticky="w",
        )
        ttk.Entry(scope_frame, textvariable=self.batch_size_var, width=9).grid(
            row=0,
            column=3,
            sticky="w",
            padx=(8, 24),
        )
        ttk.Label(scope_frame, text="表头行数").grid(row=0, column=4, sticky="w")
        ttk.Entry(scope_frame, textvariable=self.header_rows_var, width=7).grid(
            row=0,
            column=5,
            sticky="w",
            padx=(8, 0),
        )
        output_frame = create_section(parent, title="输出", row=1)
        ttk.Label(output_frame, text="batch 目录（可选）").grid(
            row=0,
            column=0,
            sticky="w",
        )
        self.split_output_dir_display = create_file_path_display(
            output_frame,
            variable=self.split_output_dir_var,
        )
        self.split_output_dir_display.grid(row=0, column=1, sticky="w", padx=(12, 8))
        ttk.Button(
            output_frame,
            text="选择目录",
            command=self.choose_split_output_dir,
        ).grid(row=0, column=2, sticky="ew")
        self.split_button = ttk.Button(
            parent,
            text="开始拆分",
            command=self.run_split,
            style=PRIMARY_BUTTON_STYLE,
        )
        self.split_button.grid(row=2, column=0, sticky="ew")
        add_optional_status_label(
            parent,
            variable=self.split_status_var,
            row=3,
        )
        self.split_output_preview_label = add_optional_status_label(
            parent,
            variable=self.split_output_preview_var,
            row=4,
        )
        parent.columnconfigure(0, weight=1)

    def _build_restore_page(self, parent: ttk.Frame) -> None:
        input_frame = create_section(parent, title="batch 文件", row=0)
        ttk.Label(input_frame, text="batch 目录").grid(row=0, column=0, sticky="w")
        self.restore_batch_dir_display = create_file_path_display(
            input_frame,
            variable=self.restore_batch_dir_var,
        )
        self.restore_batch_dir_display.grid(row=0, column=1, sticky="w", padx=(12, 8))
        ttk.Button(
            input_frame,
            text="选择目录",
            command=self.choose_restore_batch_dir,
        ).grid(row=0, column=2, sticky="ew")
        output_frame = create_section(parent, title="输出", row=1)
        ttk.Label(output_frame, text="复原 Excel（可选）").grid(
            row=0,
            column=0,
            sticky="w",
        )
        self.restore_output_file_display = create_file_path_display(
            output_frame,
            variable=self.restore_output_file_var,
        )
        self.restore_output_file_display.grid(row=0, column=1, sticky="w", padx=(12, 8))
        ttk.Button(
            output_frame,
            text="选择文件",
            command=self.choose_restore_output_file,
        ).grid(row=0, column=2, sticky="ew")
        self.restore_button = ttk.Button(
            parent,
            text="开始复原",
            command=self.run_restore,
            style=PRIMARY_BUTTON_STYLE,
        )
        self.restore_button.grid(row=2, column=0, sticky="ew")
        add_optional_status_label(
            parent,
            variable=self.restore_status_var,
            row=3,
        )
        parent.columnconfigure(0, weight=1)

    def choose_split_input_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="选择要拆分的 Excel 文件",
            filetypes=EXCEL_FILE_TYPES,
        )
        if not file_path:
            return
        self.split_input_file_var.set(file_path)
        self.refresh_sheet_choices()
        self.update_split_output_preview()

    def clear_sheet_choices(self) -> None:
        self.split_sheet_combobox["values"] = ()
        self.split_sheet_var.set("")

    def refresh_sheet_choices(self, show_error: bool = True) -> None:
        input_file = self.split_input_file_var.get().strip()
        if not input_file:
            self.clear_sheet_choices()
            return
        try:
            choices = list_workbook_sheets(input_file)
        except Exception as exc:
            self.clear_sheet_choices()
            if show_error:
                messagebox.showerror("读取失败", str(exc))
            return

        self.split_sheet_combobox["values"] = choices.sheet_names
        selected_sheet = self.split_sheet_var.get().strip()
        if selected_sheet not in choices.sheet_names:
            self.split_sheet_var.set(
                choices.default_sheet
                or (choices.sheet_names[0] if choices.sheet_names else "")
            )

    def choose_split_output_dir(self) -> None:
        input_file = self.split_input_file_var.get().strip()
        initial_dir = str(Path(input_file).parent) if input_file else None
        directory = filedialog.askdirectory(
            title="选择空的 batch 输出目录",
            initialdir=initial_dir,
        )
        if directory:
            self.split_output_dir_var.set(directory)
            self.update_split_output_preview()

    def update_split_output_preview(self) -> None:
        output_dir = self.split_output_dir_var.get().strip()
        input_file = self.split_input_file_var.get().strip()
        if output_dir:
            self.split_output_preview_var.set(f"输出目录：{output_dir}")
        elif input_file:
            self.split_output_preview_var.set(
                f"输出目录：{build_default_output_dir(input_file).name}"
            )
        else:
            self.split_output_preview_var.set("")

    def _parse_header_rows(self) -> int:
        value = self.header_rows_var.get().strip() or "1"
        try:
            header_rows = int(value)
        except ValueError as exc:
            raise ValueError("表头行数必须是整数。") from exc
        if header_rows < 0:
            raise ValueError("表头行数不能小于 0。")
        return header_rows

    def run_split(self) -> None:
        input_file = self.split_input_file_var.get().strip()
        if not input_file:
            messagebox.showerror("缺少文件", "请先选择要拆分的 Excel 文件。")
            return
        try:
            options = {
                "input_file": input_file,
                "sheet": self.split_sheet_var.get().strip() or None,
                "batch_size": parse_positive_int(
                    self.batch_size_var.get(),
                    default=1000,
                    field_name="每个 batch 的行数",
                ),
                "header_rows": self._parse_header_rows(),
                "output_dir": self.split_output_dir_var.get().strip() or None,
            }
        except Exception as exc:
            messagebox.showerror("拆分失败", str(exc))
            return
        self._start_split_worker(options)

    def _start_split_worker(self, options: dict[str, object]) -> None:
        self.split_button.state(["disabled"])
        self.split_status_var.set("正在读取 Excel…")
        threading.Thread(
            target=self._run_split_worker,
            args=(options,),
            daemon=True,
        ).start()

    def _run_split_worker(self, options: dict[str, object]) -> None:
        def report_progress(completed: int, total: int) -> None:
            self.after(
                0,
                self.split_status_var.set,
                f"正在拆分 {completed}/{total}",
            )

        try:
            summary = split_workbook(
                **options,
                progress_callback=report_progress,
            )
        except Exception as exc:
            self.after(0, self._finish_split_error, str(exc))
            return
        self.after(0, self._finish_split_success, summary)

    def _finish_split_error(self, error_message: str) -> None:
        self.split_button.state(["!disabled"])
        self.split_status_var.set("")
        messagebox.showerror("拆分失败", error_message)

    def _finish_split_success(self, summary) -> None:
        self.split_button.state(["!disabled"])
        self.split_status_var.set("")
        self.split_output_dir_var.set(str(summary.output_dir))
        self.update_split_output_preview()
        messagebox.showinfo(
            "拆分完成",
            "\n".join(
                [
                    "Excel batch 已生成。",
                    f"工作表: {summary.worksheet_title}",
                    f"数据行数: {summary.data_row_count}",
                    f"batch 数: {summary.batch_count}",
                    f"输出目录: {summary.output_dir}",
                ]
            ),
        )

    def choose_restore_batch_dir(self) -> None:
        directory = filedialog.askdirectory(title="选择包含 batch manifest 的目录")
        if directory:
            self.restore_batch_dir_var.set(directory)

    def choose_restore_output_file(self) -> None:
        batch_dir = self.restore_batch_dir_var.get().strip()
        suggested_path = None
        if batch_dir:
            try:
                suggested_path = build_default_restore_path(batch_dir)
            except Exception:
                suggested_path = None
        initial_dir = (
            str(suggested_path.parent)
            if suggested_path
            else (str(Path(batch_dir).parent) if batch_dir else None)
        )
        default_extension = suggested_path.suffix if suggested_path else ".xlsx"
        file_path = filedialog.asksaveasfilename(
            title="选择复原 Excel 文件",
            defaultextension=default_extension,
            filetypes=EXCEL_FILE_TYPES,
            initialdir=initial_dir,
            initialfile=suggested_path.name if suggested_path else None,
        )
        if file_path:
            self.restore_output_file_var.set(file_path)

    def run_restore(self) -> None:
        batch_dir = self.restore_batch_dir_var.get().strip()
        if not batch_dir:
            messagebox.showerror("缺少目录", "请先选择 batch 目录。")
            return
        options = {
            "manifest_or_directory": batch_dir,
            "output_file": self.restore_output_file_var.get().strip() or None,
        }
        self._start_restore_worker(options)

    def _start_restore_worker(self, options: dict[str, object]) -> None:
        self.restore_button.state(["disabled"])
        self.restore_status_var.set("正在读取 batch…")
        threading.Thread(
            target=self._run_restore_worker,
            args=(options,),
            daemon=True,
        ).start()

    def _run_restore_worker(self, options: dict[str, object]) -> None:
        def report_progress(completed: int, total: int) -> None:
            self.after(
                0,
                self.restore_status_var.set,
                f"正在复原 {completed}/{total}",
            )

        try:
            summary = restore_batches(
                **options,
                progress_callback=report_progress,
            )
        except Exception as exc:
            self.after(0, self._finish_restore_error, str(exc))
            return
        self.after(0, self._finish_restore_success, summary)

    def _finish_restore_error(self, error_message: str) -> None:
        self.restore_button.state(["!disabled"])
        self.restore_status_var.set("")
        messagebox.showerror("复原失败", error_message)

    def _finish_restore_success(self, summary) -> None:
        self.restore_button.state(["!disabled"])
        self.restore_status_var.set("")
        self.restore_output_file_var.set(str(summary.output_path))
        messagebox.showinfo(
            "复原完成",
            "\n".join(
                [
                    "batch 已复原为完整 Excel。",
                    f"工作表: {summary.worksheet_title}",
                    f"batch 数: {summary.batch_count}",
                    f"复原数据行数: {summary.restored_row_count}",
                    f"输出文件: {summary.output_path}",
                ]
            ),
        )


def main() -> None:
    root = create_application_root()
    root.title("Batch 拆分")
    root.resizable(True, True)
    app = ExcelBatcherApp(root)
    app.grid(row=0, column=0, sticky="nsew")
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    root.mainloop()


if __name__ == "__main__":
    main()
