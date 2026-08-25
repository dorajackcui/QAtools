#!/usr/bin/env python3
"""Tkinter page for merging active worksheets."""

from __future__ import annotations

from pathlib import Path
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from tools.gui_common import (
    MUTED_LABEL_STYLE,
    PRIMARY_BUTTON_STYLE,
    add_optional_status_label,
    configure_tool_page_style,
    create_application_root,
    create_file_path_display,
    create_section,
)

from .merge_active_sheets import (
    MergeSummary,
    build_default_output_path,
    merge_active_sheets,
)


class MergeActiveSheetsApp(ttk.Frame):
    """Merge the active sheet of each workbook under a selected folder."""

    def __init__(self, master: tk.Misc) -> None:
        super().__init__(master, padding=16)
        self.input_folder_var = tk.StringVar()
        self.keep_all_headers_var = tk.BooleanVar(value=False)
        self.output_preview_var = tk.StringVar()
        self.status_var = tk.StringVar()
        self._build_ui()

    def _build_ui(self) -> None:
        configure_tool_page_style(self)
        input_frame = create_section(self, title="输入目录与表头", row=0)
        ttk.Label(input_frame, text="Excel 所在目录").grid(
            row=0,
            column=0,
            sticky="w",
        )
        self.input_folder_display = create_file_path_display(
            input_frame,
            variable=self.input_folder_var,
        )
        self.input_folder_display.grid(
            row=0,
            column=1,
            sticky="w",
            padx=(12, 8),
        )
        ttk.Button(
            input_frame,
            text="选择目录",
            command=self.choose_input_folder,
        ).grid(row=0, column=2, sticky="ew")

        ttk.Checkbutton(
            input_frame,
            text="保留每个文件的表头",
            variable=self.keep_all_headers_var,
        ).grid(
            row=1,
            column=0,
            columnspan=3,
            sticky="w",
            pady=(14, 0),
        )
        ttk.Label(
            input_frame,
            text=(
                "递归读取目录中的 .xlsx/.xlsm，合并每个文件当前活动的工作表，"
                "并在首列写入 SourceFile。默认只保留第一份表头。"
            ),
            style=MUTED_LABEL_STYLE,
            wraplength=760,
        ).grid(
            row=2,
            column=0,
            columnspan=3,
            sticky="w",
            pady=(12, 0),
        )

        self.merge_button = ttk.Button(
            self,
            text="开始合并",
            command=self.run_merge,
            style=PRIMARY_BUTTON_STYLE,
        )
        self.merge_button.grid(row=1, column=0, sticky="ew")
        self.status_label = add_optional_status_label(
            self,
            variable=self.status_var,
            row=2,
        )
        self.output_preview_label = add_optional_status_label(
            self,
            variable=self.output_preview_var,
            row=3,
        )
        self.columnconfigure(0, weight=1)

    def choose_input_folder(self) -> None:
        current_folder = self.input_folder_var.get().strip()
        directory = filedialog.askdirectory(
            title="选择包含待合并 Excel 的目录",
            initialdir=current_folder or None,
        )
        if not directory:
            return
        self.input_folder_var.set(directory)
        self.update_output_preview()

    def update_output_preview(self) -> None:
        folder_text = self.input_folder_var.get().strip()
        if not folder_text:
            self.output_preview_var.set("")
            return
        output_path = build_default_output_path(Path(folder_text))
        self.output_preview_var.set(f"输出文件：{output_path}")

    def run_merge(self) -> None:
        folder_text = self.input_folder_var.get().strip()
        if not folder_text:
            messagebox.showerror("缺少目录", "请先选择包含 Excel 的目录。")
            return
        output_path = build_default_output_path(Path(folder_text))
        self.output_preview_var.set(f"输出文件：{output_path}")
        options = {
            "folder_path": folder_text,
            "output_path": output_path,
            "keep_all_headers": self.keep_all_headers_var.get(),
        }
        self._start_merge_worker(options)

    def _start_merge_worker(self, options: dict[str, object]) -> None:
        self.merge_button.state(["disabled"])
        self.status_var.set("正在读取 Excel…")
        threading.Thread(
            target=self._run_merge_worker,
            args=(options,),
            daemon=True,
        ).start()

    def _run_merge_worker(self, options: dict[str, object]) -> None:
        def report_progress(completed: int, total: int) -> None:
            self.after(
                0,
                self.status_var.set,
                f"正在合并 {completed}/{total}",
            )

        try:
            summary = merge_active_sheets(
                **options,
                progress_callback=report_progress,
            )
        except Exception as exc:  # noqa: BLE001
            self.after(0, self._finish_merge_error, str(exc))
            return
        self.after(0, self._finish_merge_success, summary)

    def _finish_merge_error(self, error_message: str) -> None:
        self.merge_button.state(["!disabled"])
        self.status_var.set("")
        messagebox.showerror("合并失败", error_message)

    def _finish_merge_success(self, summary: MergeSummary) -> None:
        self.merge_button.state(["!disabled"])
        self.status_var.set("")
        self.output_preview_var.set(f"输出文件：{summary.output_path}")
        lines = [
            "活动工作表已合并。",
            f"输入文件数: {summary.supported_file_count}",
            f"输出行数: {summary.merged_row_count}",
            f"跳过 .xls/.xlsb: {summary.skipped_file_count}",
            f"读取失败: {summary.failed_file_count}",
            f"输出文件: {summary.output_path}",
        ]
        if summary.error_log_path is not None:
            lines.append(f"错误日志: {summary.error_log_path}")
        if summary.failed_file_count:
            messagebox.showwarning("合并完成（有文件失败）", "\n".join(lines))
        else:
            messagebox.showinfo("合并完成", "\n".join(lines))


def main() -> None:
    root = create_application_root()
    root.title("合并表格")
    root.resizable(True, True)
    app = MergeActiveSheetsApp(root)
    app.grid(row=0, column=0, sticky="nsew")
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    root.mainloop()


if __name__ == "__main__":
    main()
