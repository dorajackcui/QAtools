#!/usr/bin/env python3
"""Desktop UI for mark-based and history-TB terminology checking."""

from __future__ import annotations

from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from tools.excel_metadata import detect_source_target_columns, list_workbook_sheets
from tools.gui_common import (
    APP_MAIN_BACKGROUND,
    MUTED_LABEL_STYLE,
    PRIMARY_BUTTON_STYLE,
    SECTION_FRAME_STYLE,
    add_optional_status_label,
    configure_tool_page_style,
    create_application_root,
    create_file_path_display,
    parse_positive_int,
)
from tools.tb_project_ui import TbProjectControls
from tools.tb_projects import TbProject

try:
    from .extract_terms_from_excel import (
        TERM_SHEET_NAME,
        build_default_output_path,
        detect_history_tb_columns,
        process_excel,
    )
except ImportError:
    from extract_terms_from_excel import (
        TERM_SHEET_NAME,
        build_default_output_path,
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
        self.output_preview_var = tk.StringVar()
        self.history_details_button_text_var = tk.StringVar(value="展开详情")
        self.history_details_expanded = False
        self.mark_style_vars = {
            "【】": tk.BooleanVar(value=True),
            "[]": tk.BooleanVar(value=True),
        }
        self._build_ui()

    def _build_ui(self) -> None:
        configure_tool_page_style(self)
        self.scroll_canvas = tk.Canvas(
            self,
            width=900,
            height=430,
            background=APP_MAIN_BACKGROUND,
            borderwidth=0,
            highlightthickness=0,
            yscrollincrement=16,
        )
        self.scroll_canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(
            self,
            orient="vertical",
            command=self.scroll_canvas.yview,
        )
        scrollbar.grid(row=0, column=1, sticky="ns", padx=(8, 0))
        self.scroll_canvas.configure(yscrollcommand=scrollbar.set)

        self.scroll_content = ttk.Frame(self.scroll_canvas)
        self.scroll_content.columnconfigure(0, weight=1)
        self.scroll_content_window = self.scroll_canvas.create_window(
            (0, 0),
            window=self.scroll_content,
            anchor="nw",
        )
        self.scroll_content.bind("<Configure>", self.handle_scroll_content_configure)
        self.scroll_canvas.bind("<Configure>", self.handle_scroll_canvas_configure)
        self.scroll_canvas.bind("<Enter>", self.bind_term_mousewheel)
        self.scroll_canvas.bind("<Leave>", self.unbind_term_mousewheel)

        input_frame = ttk.LabelFrame(
            self.scroll_content,
            text="输入与范围",
            padding=12,
            style=SECTION_FRAME_STYLE,
        )
        input_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        input_frame.columnconfigure(1, weight=0)

        ttk.Label(input_frame, text="输入 Excel").grid(row=0, column=0, sticky="w")
        self.input_file_display = create_file_path_display(
            input_frame,
            variable=self.input_file_var,
        )
        self.input_file_display.grid(row=0, column=1, sticky="w", padx=(12, 8))
        ttk.Button(input_frame, text="选择文件", command=self.choose_input_file).grid(
            row=0, column=2, sticky="ew"
        )

        scope_frame = ttk.Frame(input_frame)
        scope_frame.grid(row=1, column=0, columnspan=3, sticky="w", pady=(12, 0))
        ttk.Label(scope_frame, text="检查工作表").grid(row=0, column=0, sticky="w")
        self.sheet_combobox = ttk.Combobox(
            scope_frame,
            textvariable=self.sheet_var,
            width=18,
            state="readonly",
        )
        self.sheet_combobox.grid(row=0, column=1, sticky="w", padx=(8, 18))
        self.sheet_combobox.bind("<<ComboboxSelected>>", self.handle_sheet_selected)
        ttk.Label(scope_frame, text="Source 列").grid(row=0, column=2, sticky="w")
        ttk.Entry(scope_frame, textvariable=self.source_column_var, width=7).grid(
            row=0, column=3, sticky="w", padx=(8, 18)
        )
        ttk.Label(scope_frame, text="Target 列").grid(row=0, column=4, sticky="w")
        ttk.Entry(scope_frame, textvariable=self.target_column_var, width=7).grid(
            row=0, column=5, sticky="w", padx=(8, 18)
        )
        ttk.Label(scope_frame, text="开始行").grid(row=0, column=6, sticky="w")
        ttk.Spinbox(
            scope_frame,
            textvariable=self.start_row_var,
            width=7,
            from_=1,
            to=1_000_000,
        ).grid(row=0, column=7, sticky="w", padx=(8, 0))

        source_frame = ttk.LabelFrame(
            self.scroll_content,
            text="术语来源",
            padding=12,
            style=SECTION_FRAME_STYLE,
        )
        source_frame.grid(row=1, column=0, sticky="ew")
        source_frame.columnconfigure(1, weight=0)
        ttk.Label(source_frame, text="术语标记").grid(row=0, column=0, sticky="w")
        mark_frame = ttk.Frame(source_frame)
        mark_frame.grid(row=0, column=1, columnspan=2, sticky="w", padx=(12, 0))
        ttk.Checkbutton(
            mark_frame,
            text="中文方括号【】",
            variable=self.mark_style_vars["【】"],
        ).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(
            mark_frame,
            text="半角方括号 []",
            variable=self.mark_style_vars["[]"],
        ).grid(row=0, column=1, sticky="w", padx=(16, 0))

        self.tb_project_controls = TbProjectControls(
            source_frame,
            capture_project=self.capture_tb_project,
            apply_project=self.apply_tb_project,
        )
        self.tb_project_controls.grid(
            row=1,
            column=0,
            columnspan=3,
            sticky="ew",
            pady=(12, 0),
        )

        ttk.Label(source_frame, text="历史 TB（可选）").grid(
            row=2, column=0, sticky="w", pady=(12, 0)
        )
        history_entry = create_file_path_display(
            source_frame,
            variable=self.history_tb_file_var,
        )
        history_entry.grid(row=2, column=1, sticky="w", padx=(12, 8), pady=(12, 0))
        history_buttons = ttk.Frame(source_frame)
        history_buttons.grid(row=2, column=2, sticky="e", pady=(12, 0))
        ttk.Button(
            history_buttons,
            text="选择文件",
            command=self.choose_history_tb_file,
        ).grid(row=0, column=0)
        ttk.Button(
            history_buttons,
            text="清空",
            command=self.clear_history_tb_file,
        ).grid(row=0, column=1, padx=(6, 0))
        self.history_details_button = ttk.Button(
            history_buttons,
            textvariable=self.history_details_button_text_var,
            command=self.toggle_history_details,
            state="disabled",
            width=8,
        )
        self.history_details_button.grid(row=0, column=2, padx=(6, 0))

        self.history_details_frame = ttk.LabelFrame(
            source_frame,
            text="历史 TB 详情",
            padding=10,
            style=SECTION_FRAME_STYLE,
        )
        self.history_details_frame.grid(
            row=3,
            column=0,
            columnspan=3,
            sticky="ew",
            pady=(12, 0),
        )
        ttk.Label(self.history_details_frame, text="工作表").grid(
            row=0, column=0, sticky="w"
        )
        self.history_sheet_combobox = ttk.Combobox(
            self.history_details_frame,
            textvariable=self.history_sheet_var,
            width=16,
            state="readonly",
        )
        self.history_sheet_combobox.grid(row=0, column=1, sticky="w", padx=(8, 18))
        self.history_sheet_combobox.bind(
            "<<ComboboxSelected>>",
            self.handle_history_sheet_selected,
        )
        ttk.Label(self.history_details_frame, text="Source 列").grid(
            row=0, column=2, sticky="w"
        )
        ttk.Entry(
            self.history_details_frame,
            textvariable=self.history_source_column_var,
            width=7,
        ).grid(row=0, column=3, sticky="w", padx=(8, 18))
        ttk.Label(self.history_details_frame, text="Target 列").grid(
            row=0, column=4, sticky="w"
        )
        ttk.Entry(
            self.history_details_frame,
            textvariable=self.history_target_column_var,
            width=7,
        ).grid(row=0, column=5, sticky="w", padx=(8, 18))
        ttk.Label(self.history_details_frame, text="开始行").grid(
            row=0, column=6, sticky="w"
        )
        ttk.Spinbox(
            self.history_details_frame,
            textvariable=self.history_start_row_var,
            width=7,
            from_=1,
            to=1_000_000,
        ).grid(row=0, column=7, sticky="w", padx=(8, 0))
        self.history_details_frame.grid_remove()

        ttk.Label(
            source_frame,
            text=(
                "有标记时提取并检查新术语；不选标记时必须提供历史 TB；"
                "历史 TB 命中时优先使用历史 target。"
            ),
            style=MUTED_LABEL_STYLE,
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(12, 0))

        ttk.Button(
            self,
            text="开始检查",
            command=self.run_extraction,
            style=PRIMARY_BUTTON_STYLE,
        ).grid(row=1, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        self.output_preview_label = add_optional_status_label(
            self,
            variable=self.output_preview_var,
            row=2,
            columnspan=2,
        )

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

    def handle_scroll_content_configure(self, _event: object | None = None) -> None:
        self.refresh_scroll_region()

    def handle_scroll_canvas_configure(self, event: object) -> None:
        width = getattr(event, "width", self.scroll_canvas.winfo_width())
        self.scroll_canvas.itemconfigure(self.scroll_content_window, width=width)
        self.refresh_scroll_region()

    def refresh_scroll_region(self) -> None:
        content_bounds = self.scroll_canvas.bbox("all")
        if content_bounds:
            self.scroll_canvas.configure(scrollregion=content_bounds)

    def bind_term_mousewheel(self, _event: object | None = None) -> None:
        self.bind_all("<MouseWheel>", self.handle_term_mousewheel)

    def unbind_term_mousewheel(self, _event: object | None = None) -> None:
        self.unbind_all("<MouseWheel>")

    def handle_term_mousewheel(self, event: object) -> None:
        delta = getattr(event, "delta", 0)
        if delta:
            self.scroll_canvas.yview_scroll(-1 if delta > 0 else 1, "units")

    def reveal_history_details(self) -> None:
        self.update_idletasks()
        self.refresh_scroll_region()
        content_height = max(self.scroll_content.winfo_reqheight(), 1)
        target_fraction = max(
            0.0,
            min(self.history_details_frame.winfo_y() / content_height, 1.0),
        )
        self.scroll_canvas.yview_moveto(target_fraction)

    def set_history_details_expanded(self, expanded: bool) -> None:
        self.history_details_expanded = expanded
        self.history_details_button_text_var.set("收起详情" if expanded else "展开详情")
        if expanded:
            self.history_details_frame.grid()
            if hasattr(self, "scroll_canvas"):
                self.after_idle(self.reveal_history_details)
        else:
            self.history_details_frame.grid_remove()

    def toggle_history_details(self) -> None:
        if self.history_tb_file_var.get().strip():
            self.set_history_details_expanded(not self.history_details_expanded)

    def update_output_preview(self) -> None:
        input_file = self.input_file_var.get().strip()
        if not input_file:
            self.output_preview_var.set("")
            return
        output_name = build_default_output_path(input_file).name
        self.output_preview_var.set(f"输出文件：{output_name}")

    def handle_input_file_focus_out(self, _event: object | None = None) -> None:
        self.update_output_preview()
        self.refresh_sheet_choices(show_error=False)

    def handle_history_file_focus_out(self, _event: object | None = None) -> None:
        if hasattr(self, "tb_project_controls"):
            self.tb_project_controls.mark_current_settings_modified()
        if not self.history_tb_file_var.get().strip():
            self.clear_history_tb_file()
            return
        self.history_details_button.configure(state="normal")
        self.refresh_history_sheet_choices(show_error=False)
        self.set_history_details_expanded(True)

    def choose_input_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if not file_path:
            return
        self.input_file_var.set(file_path)
        self.update_output_preview()
        self.refresh_sheet_choices()

    def choose_history_tb_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="选择历史 TB Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if not file_path:
            return
        if hasattr(self, "tb_project_controls"):
            self.tb_project_controls.mark_current_settings_modified()
        self.history_tb_file_var.set(file_path)
        self.history_details_button.configure(state="normal")
        self.refresh_history_sheet_choices()
        self.set_history_details_expanded(True)

    def clear_history_tb_file(self) -> None:
        if hasattr(self, "tb_project_controls"):
            self.tb_project_controls.clear_selection()
        self.history_tb_file_var.set("")
        self.history_source_column_var.set("")
        self.history_target_column_var.set("")
        self.history_start_row_var.set("2")
        self.clear_history_sheet_choices()
        if hasattr(self, "history_details_button"):
            self.history_details_button.configure(state="disabled")
        if hasattr(self, "history_details_frame"):
            self.set_history_details_expanded(False)

    def capture_tb_project(self, project_name: str) -> TbProject:
        file_path = self.history_tb_file_var.get().strip()
        if not file_path:
            raise ValueError("请先选择历史 TB 文件。")
        if not Path(file_path).expanduser().is_file():
            raise ValueError(f"历史 TB 文件不存在：{file_path}")
        sheet = self.history_sheet_var.get().strip()
        source_column = self.history_source_column_var.get().strip()
        target_column = self.history_target_column_var.get().strip()
        if not sheet or not source_column or not target_column:
            raise ValueError("请先确认历史 TB 的工作表及 Source / Target 列。")
        start_row = parse_positive_int(
            self.history_start_row_var.get(),
            default=2,
            field_name="术语历史开始行",
        )
        return TbProject(
            name=project_name,
            file_path=str(Path(file_path).expanduser().absolute()),
            sheet=sheet,
            source_column=source_column,
            target_column=target_column,
            start_row=start_row,
        )

    def apply_tb_project(self, project: TbProject) -> None:
        self.history_tb_file_var.set(project.file_path)
        self.history_sheet_var.set(project.sheet)
        self.refresh_history_sheet_choices(show_error=False)
        self.history_sheet_var.set(project.sheet)
        self.history_source_column_var.set(project.source_column)
        self.history_target_column_var.set(project.target_column)
        self.history_start_row_var.set(str(project.start_row))
        if hasattr(self, "history_details_button"):
            self.history_details_button.configure(state="normal")
        if hasattr(self, "history_details_frame"):
            self.set_history_details_expanded(True)

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
        if not selected_mark_styles and not history_tb_file:
            messagebox.showerror(
                "缺少术语来源",
                "请至少选择一种术语 mark，或提供历史 TB。",
            )
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
            )
        except Exception as exc:
            messagebox.showerror("处理失败", str(exc))
            return

        message_lines = [
            "术语检查已完成。",
            f"工作表: {worksheet_title}",
            f"source 列: {source_col}",
            f"target 列: {target_col}",
            f"术语 mark: {'、'.join(selected_mark_styles) if selected_mark_styles else '未选择（仅历史 TB）'}",
        ]
        if history_tb_file:
            message_lines.append(f"历史 TB: {history_tb_file}")
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
    root = create_application_root()
    root.title("Excel 术语检查")
    root.resizable(True, True)
    app = ExtractTermsApp(root)
    app.grid(row=0, column=0, sticky="nsew")
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    root.mainloop()


if __name__ == "__main__":
    main()
