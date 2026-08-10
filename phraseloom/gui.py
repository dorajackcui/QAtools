from __future__ import annotations

import contextlib
import io
import queue
import sys
import threading
from dataclasses import dataclass
from pathlib import Path
from typing import Mapping

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    from phraseloom.strings_workflow import (
        default_restored_output_path,
        default_strings_output_path,
    )
else:
    from .strings_workflow import (
        default_restored_output_path,
        default_strings_output_path,
    )

from tools.gui_common import (
    MUTED_LABEL_STYLE,
    PRIMARY_BUTTON_STYLE,
    configure_tool_page_style,
    create_section,
)


EXCEL_FILE_TYPES = (("Excel 文件", "*.xlsx"), ("所有文件", "*.*"))
TOML_FILE_TYPES = (("TOML 配置", "*.toml"), ("所有文件", "*.*"))


@dataclass(frozen=True)
class FieldSpec:
    key: str
    label: str
    kind: str = "text"
    flag: str | None = None
    required: bool = False
    default: str | bool = ""
    file_types: tuple[tuple[str, str], ...] = EXCEL_FILE_TYPES


@dataclass(frozen=True)
class TaskSpec:
    key: str
    label: str
    description: str
    command: str
    fields: tuple[FieldSpec, ...]


def _input(label: str) -> FieldSpec:
    return FieldSpec("input", label, kind="open", required=True)


EXPORT_STRINGS_TASK = TaskSpec(
    "export_strings",
    "导出 Strings",
    "清洗并合并重复文本；可选将相似句分组排列。",
    "export",
    (
        _input("原始 Excel"),
        FieldSpec(
            "group_similar",
            "启用相似句分组（未聚类在前，聚类内容在后）",
            kind="bool",
            flag="--group-similar",
            default=False,
        ),
        FieldSpec(
            "source_col",
            "Source 列",
            flag="--source-col",
            default="source",
        ),
        FieldSpec(
            "target_col",
            "Target 列",
            flag="--target-col",
            default="target",
        ),
        FieldSpec(
            "context_col",
            "Context 列",
            flag="--context-col",
        ),
        FieldSpec(
            "tag_config",
            "Tag 配置（可选）",
            kind="open",
            flag="--tag-config",
            file_types=TOML_FILE_TYPES,
        ),
    ),
)

RESTORE_STRINGS_TASK = TaskSpec(
    "restore_strings",
    "回填译文",
    "读取翻译完成的 Strings 工作簿，写回原始位置并恢复原表结构。",
    "restore",
    (_input("翻译完成的 Strings 工作簿"),),
)

# GUI 只展示导出页面；回填是页面右下角的一次性文件操作。
TASKS: tuple[TaskSpec, ...] = (EXPORT_STRINGS_TASK,)
TASK_BY_KEY = {
    task.key: task
    for task in (EXPORT_STRINGS_TASK, RESTORE_STRINGS_TASK)
}


def build_cli_args(task: TaskSpec, values: Mapping[str, object]) -> list[str]:
    args = [task.command]
    missing: list[str] = []
    for field_spec in task.fields:
        value = values.get(field_spec.key, field_spec.default)
        if field_spec.kind == "bool":
            enabled = (
                value
                if isinstance(value, bool)
                else str(value).strip().lower() in {"1", "true", "yes", "on"}
            )
            if field_spec.required and not enabled:
                missing.append(field_spec.label)
            if enabled and field_spec.flag:
                args.append(field_spec.flag)
            continue

        text = str(value or "").strip()
        if field_spec.required and not text:
            missing.append(field_spec.label)
            continue
        if not text:
            continue
        if field_spec.kind == "int":
            try:
                if int(text) < 1:
                    raise ValueError
            except ValueError as exc:
                raise ValueError(f"{field_spec.label}必须是大于 0 的整数") from exc
        if field_spec.flag:
            args.extend([field_spec.flag, text])
        else:
            args.append(text)

    if missing:
        raise ValueError("请填写：" + "、".join(missing))
    return args


def validate_task_specs() -> None:
    task_specs = tuple(TASK_BY_KEY.values())
    keys = [task.key for task in task_specs]
    if len(keys) != len(set(keys)):
        raise ValueError("GUI task keys must be unique")
    for task in task_specs:
        field_keys = [field.key for field in task.fields]
        if len(field_keys) != len(set(field_keys)):
            raise ValueError(f"GUI field keys must be unique: {task.key}")
        if not any(field.required for field in task.fields):
            raise ValueError(f"GUI task needs a required input: {task.key}")


class PhraseLoomApp(ttk.Frame):
    """Reusable PhraseLoom page for Toolshub and the standalone window."""

    def __init__(self, parent: tk.Misc, *, standalone: bool = False) -> None:
        super().__init__(parent, padding=16)
        _ = standalone
        self.root = self.winfo_toplevel()
        self.current_title = tk.StringVar(value=EXPORT_STRINGS_TASK.label)
        self.current_description = tk.StringVar(
            value=EXPORT_STRINGS_TASK.description
        )
        self.output_preview_var = tk.StringVar(
            value="输出文件：选择输入 Excel 后自动生成"
        )
        self.status_var = tk.StringVar(value="请选择输入文件")
        self.last_result_text = ""
        self.export_values = {
            field.key: field.default
            for field in EXPORT_STRINGS_TASK.fields
        }
        self.field_vars: dict[str, tk.Variable] = {}
        self.result_queue: queue.Queue[tuple[int, str, str]] = queue.Queue()
        self.running_task_key: str | None = None

        configure_tool_page_style(self)
        self._build_layout()
        self._render_task()

    def _build_layout(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        style = ttk.Style(self)
        canvas_background = style.lookup("TFrame", "background") or "#f0f0f0"
        self.scroll_canvas = tk.Canvas(
            self,
            width=900,
            height=480,
            background=canvas_background,
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

        self.content_frame = ttk.Frame(self.scroll_canvas)
        self.content_frame.columnconfigure(0, weight=1)
        self.scroll_content_window = self.scroll_canvas.create_window(
            (0, 0),
            window=self.content_frame,
            anchor="nw",
        )
        self.content_frame.bind(
            "<Configure>",
            self._handle_scroll_content_configure,
        )
        self.scroll_canvas.bind(
            "<Configure>",
            self._handle_scroll_canvas_configure,
        )
        self.scroll_canvas.bind("<Enter>", self._bind_mousewheel)
        self.scroll_canvas.bind("<Leave>", self._unbind_mousewheel)

    def _handle_scroll_content_configure(self, _event=None) -> None:
        self._refresh_scroll_region()

    def _handle_scroll_canvas_configure(self, event) -> None:
        width = getattr(event, "width", self.scroll_canvas.winfo_width())
        self.scroll_canvas.itemconfigure(
            self.scroll_content_window,
            width=width,
        )
        self._refresh_scroll_region()

    def _refresh_scroll_region(self) -> None:
        content_bounds = self.scroll_canvas.bbox("all")
        if content_bounds:
            self.scroll_canvas.configure(scrollregion=content_bounds)

    def _bind_mousewheel(self, _event=None) -> None:
        self.bind_all("<MouseWheel>", self._handle_mousewheel)

    def _unbind_mousewheel(self, _event=None) -> None:
        self.unbind_all("<MouseWheel>")

    def _handle_mousewheel(self, event) -> None:
        delta = getattr(event, "delta", 0)
        if delta:
            self.scroll_canvas.yview_scroll(-1 if delta > 0 else 1, "units")

    def _capture_values(self) -> None:
        if not self.field_vars:
            return
        for key, variable in self.field_vars.items():
            self.export_values[key] = variable.get()

    def _render_task(self) -> None:
        task = EXPORT_STRINGS_TASK
        self.current_title.set(task.label)
        self.current_description.set(task.description)
        for child in self.content_frame.winfo_children():
            child.destroy()
        self.field_vars = {}

        input_section = create_section(
            self.content_frame,
            title="输入与范围",
            row=0,
        )
        input_spec = next(field for field in task.fields if field.key == "input")
        input_var = self._make_variable(input_spec)
        self._add_file_picker_row(
            input_section,
            label=input_spec.label,
            variable=input_var,
            field_spec=input_spec,
            row=0,
            focus_out_command=self._handle_input_focus_out,
        )
        self._build_column_settings(input_section, task)
        ttk.Label(
            input_section,
            text=(
                "已有 Target 会视为已完成并跳过；重复 Source 只导出一次；"
                "单元格内多行 Source 会逐行拆成 Segment，并在回填时按原换行还原；"
                "Context 留空时自动识别同名列。"
            ),
            style=MUTED_LABEL_STYLE,
            wraplength=760,
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(12, 0))

        options = create_section(
            self.content_frame,
            title="导出选项",
            row=1,
        )
        self._build_export_options(options, task)

        action_frame = ttk.Frame(self)
        action_frame.grid(
            row=1,
            column=0,
            columnspan=2,
            sticky="ew",
            pady=(10, 0),
        )
        action_frame.columnconfigure(0, weight=3)
        action_frame.columnconfigure(1, weight=1)
        self.run_button = ttk.Button(
            action_frame,
            text="导出 Strings",
            command=self._run_current_task,
            style=PRIMARY_BUTTON_STYLE,
        )
        self.run_button.grid(row=0, column=0, sticky="ew")
        self.restore_button = ttk.Button(
            action_frame,
            text="回填译文…",
            command=self._choose_and_restore,
        )
        self.restore_button.grid(
            row=0,
            column=1,
            sticky="ew",
            padx=(8, 0),
        )

        self._update_output_preview()
        ttk.Label(
            self,
            textvariable=self.output_preview_var,
            style=MUTED_LABEL_STYLE,
        ).grid(
            row=2,
            column=0,
            columnspan=2,
            sticky="w",
            pady=(8, 0),
        )

        self.status_frame = ttk.Frame(self)
        self.status_frame.grid(
            row=3,
            column=0,
            columnspan=2,
            sticky="ew",
            pady=(8, 0),
        )
        self.status_frame.columnconfigure(1, weight=1)
        self.progress = ttk.Progressbar(
            self.status_frame,
            mode="indeterminate",
            length=140,
        )
        self.progress.grid(row=0, column=0, sticky="w")
        ttk.Label(
            self.status_frame,
            textvariable=self.status_var,
            style=MUTED_LABEL_STYLE,
        ).grid(row=0, column=1, sticky="w", padx=(12, 0))
        self.status_frame.grid_remove()
        self._refresh_scroll_region()

    def _build_column_settings(
        self,
        parent: ttk.LabelFrame,
        task: TaskSpec,
    ) -> None:
        fields = {field.key: field for field in task.fields}
        compact = ttk.Frame(parent)
        compact.grid(row=1, column=0, columnspan=3, sticky="w", pady=(12, 0))

        column_fields = (
            fields["source_col"],
            fields["target_col"],
            fields["context_col"],
        )
        for index, field_spec in enumerate(column_fields):
            ttk.Label(compact, text=field_spec.label).grid(
                row=0,
                column=index * 2,
                sticky="w",
            )
            variable = self._make_variable(field_spec)
            ttk.Entry(
                compact,
                textvariable=variable,
                width=14,
            ).grid(
                row=0,
                column=index * 2 + 1,
                sticky="w",
                padx=(8, 18 if index < 2 else 0),
            )

    def _build_export_options(
        self,
        parent: ttk.LabelFrame,
        task: TaskSpec,
    ) -> None:
        fields = {field.key: field for field in task.fields}
        group_spec = fields["group_similar"]
        group_var = self._make_variable(group_spec)
        ttk.Checkbutton(
            parent,
            text=group_spec.label,
            variable=group_var,
        ).grid(row=0, column=0, columnspan=3, sticky="w")

        tag_spec = fields["tag_config"]
        tag_var = self._make_variable(tag_spec)
        self._add_file_picker_row(
            parent,
            label=tag_spec.label,
            variable=tag_var,
            field_spec=tag_spec,
            row=1,
            pady=(12, 0),
        )
        ttk.Label(
            parent,
            text="相似句仅调整排列；Tag 配置留空时使用内置规则。",
            style=MUTED_LABEL_STYLE,
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(8, 0))

    def _make_variable(self, field_spec: FieldSpec) -> tk.Variable:
        value = self.export_values.get(
            field_spec.key,
            field_spec.default,
        )
        if field_spec.kind == "bool":
            enabled = (
                value
                if isinstance(value, bool)
                else str(value).strip().lower() in {"1", "true", "yes", "on"}
            )
            variable: tk.Variable = tk.BooleanVar(value=enabled)
        else:
            variable = tk.StringVar(value=str(value or ""))
        self.field_vars[field_spec.key] = variable
        return variable

    def _add_file_picker_row(
        self,
        parent: ttk.LabelFrame,
        *,
        label: str,
        variable: tk.StringVar,
        field_spec: FieldSpec,
        row: int,
        focus_out_command=None,
        pady: tuple[int, int] = (0, 0),
    ) -> None:
        ttk.Label(parent, text=label).grid(
            row=row,
            column=0,
            sticky="w",
            pady=pady,
        )
        entry = ttk.Entry(parent, textvariable=variable, width=56)
        entry.grid(
            row=row,
            column=1,
            sticky="ew",
            padx=(12, 8),
            pady=pady,
        )
        if focus_out_command is not None:
            entry.bind("<FocusOut>", focus_out_command)
        ttk.Button(
            parent,
            text="选择",
            command=lambda: self._choose_path(field_spec, variable),
        ).grid(row=row, column=2, sticky="ew", pady=pady)
        parent.columnconfigure(1, weight=1)

    def _choose_path(self, field_spec: FieldSpec, variable: tk.StringVar) -> None:
        value = filedialog.askopenfilename(
            parent=self.root,
            title=f"选择{field_spec.label}",
            filetypes=field_spec.file_types,
        )
        if value:
            variable.set(value)
            if field_spec.key == "input":
                self._update_output_preview()

    def _handle_input_focus_out(self, _event=None) -> None:
        self._capture_values()
        self._update_output_preview()

    def _update_output_preview(self) -> None:
        input_var = self.field_vars.get("input")
        input_text = str(input_var.get()).strip() if input_var else ""
        if not input_text:
            self.output_preview_var.set("输出文件：选择输入 Excel 后自动生成")
            return
        input_path = Path(input_text).expanduser()
        output_path = default_strings_output_path(input_path)
        self.output_preview_var.set(f"输出文件：{output_path.name}")

    def _run_current_task(self) -> None:
        self._capture_values()
        task = EXPORT_STRINGS_TASK
        try:
            args = build_cli_args(task, self.export_values)
        except ValueError as exc:
            messagebox.showerror("无法开始", str(exc), parent=self.root)
            return

        self._start_task(task, args)

    def _choose_and_restore(self) -> None:
        task = RESTORE_STRINGS_TASK
        value = filedialog.askopenfilename(
            parent=self.root,
            title="选择翻译完成的 Strings 工作簿",
            filetypes=EXCEL_FILE_TYPES,
        )
        if not value:
            return
        try:
            args = build_cli_args(task, {"input": value})
        except ValueError as exc:
            messagebox.showerror("无法开始", str(exc), parent=self.root)
            return

        input_path = Path(value)
        try:
            output_path = default_restored_output_path(input_path)
        except Exception:
            stem = input_path.stem
            if stem.endswith("_strings"):
                stem = stem[: -len("_strings")]
            output_path = input_path.with_name(f"{stem}_translated.xlsx")
        self.output_preview_var.set(f"回填输出：{output_path.name}")
        self._start_task(task, args)

    def _start_task(self, task: TaskSpec, args: list[str]) -> None:
        self.running_task_key = task.key
        self.run_button.configure(state="disabled")
        self.restore_button.configure(state="disabled")
        self.status_frame.grid()
        self.progress.start(10)
        status = "正在导出…" if task.key == "export_strings" else "正在回填…"
        self.status_var.set(status)
        self._set_output("")
        worker = threading.Thread(
            target=self._execute_cli,
            args=(args,),
            daemon=True,
        )
        worker.start()
        self.root.after(100, self._poll_result)

    def _execute_cli(self, args: list[str]) -> None:
        if __package__ in {None, ""}:
            from phraseloom.cli import main as cli_main
        else:
            from .cli import main as cli_main

        stdout = io.StringIO()
        stderr = io.StringIO()
        try:
            with contextlib.redirect_stdout(stdout), contextlib.redirect_stderr(stderr):
                exit_code = cli_main(args)
        except SystemExit as exc:
            exit_code = int(exc.code or 0)
        except Exception as exc:
            exit_code = 1
            stderr.write(f"{type(exc).__name__}: {exc}\n")
        self.result_queue.put((exit_code, stdout.getvalue(), stderr.getvalue()))

    def _poll_result(self) -> None:
        try:
            result = self.result_queue.get_nowait()
        except queue.Empty:
            self.root.after(100, self._poll_result)
            return
        self._finish_run(*result)

    def _finish_run(self, exit_code: int, output: str, error: str) -> None:
        self.progress.stop()
        self.status_frame.grid_remove()
        self.run_button.configure(state="normal")
        self.restore_button.configure(state="normal")
        finished_task_key = self.running_task_key
        self.running_task_key = None
        combined = "\n".join(part.strip() for part in (output, error) if part.strip())
        self._set_output(combined)
        if exit_code == 0:
            if finished_task_key == "restore_strings":
                self.status_var.set("回填完成")
                title = "回填完成"
            else:
                self.status_var.set("导出完成")
                self._update_output_preview()
                title = "导出完成"
            messagebox.showinfo(
                title,
                combined or title,
                parent=self.root,
            )
        else:
            self.status_var.set("未完成，请查看运行结果")
            messagebox.showerror(
                "任务未完成",
                error.strip() or output.strip() or "未知错误",
                parent=self.root,
            )

    def _set_output(self, text: str) -> None:
        self.last_result_text = text


class PhraseLoomGUI(PhraseLoomApp):
    """Backward-compatible standalone PhraseLoom window."""

    def __init__(self, root: tk.Tk) -> None:
        root.title("PhraseLoom")
        root.resizable(True, True)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        super().__init__(root, standalone=True)
        self.grid(row=0, column=0, sticky="nsew")


def main() -> int:
    root = tk.Tk()
    PhraseLoomGUI(root)
    root.mainloop()
    return 0


__all__ = [
    "EXCEL_FILE_TYPES",
    "FieldSpec",
    "PhraseLoomApp",
    "PhraseLoomGUI",
    "TASKS",
    "TASK_BY_KEY",
    "TaskSpec",
    "build_cli_args",
    "main",
    "validate_task_specs",
]


if __name__ == "__main__":
    raise SystemExit(main())
