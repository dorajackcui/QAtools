from __future__ import annotations

import contextlib
import io
import queue
import threading
from dataclasses import dataclass, field
from typing import Mapping

import tkinter as tk
from tkinter import filedialog, messagebox, ttk


EXCEL_FILE_TYPES = (("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*"))
TOML_FILE_TYPES = (("TOML 配置", "*.toml"), ("所有文件", "*.*"))


@dataclass(frozen=True)
class FieldSpec:
    key: str
    label: str
    kind: str = "text"
    flag: str | None = None
    required: bool = False
    default: str | bool = ""
    advanced: bool = False
    choices: tuple[str, ...] = ()
    file_types: tuple[tuple[str, str], ...] = EXCEL_FILE_TYPES
    help_text: str = ""


@dataclass(frozen=True)
class TaskSpec:
    key: str
    label: str
    description: str
    command: str
    fields: tuple[FieldSpec, ...]
    group: str = "daily"
    exclusive_groups: tuple[tuple[str, ...], ...] = field(default_factory=tuple)


def _input(label: str = "输入文件", *, help_text: str = "") -> FieldSpec:
    return FieldSpec("input", label, kind="open", required=True, help_text=help_text)


def _output(label: str = "输出文件（留空则自动命名）") -> FieldSpec:
    return FieldSpec("output", label, kind="save", flag="--output", advanced=True)


def _source_target_fields() -> tuple[FieldSpec, FieldSpec]:
    return (
        FieldSpec(
            "source_col",
            "Source 列",
            flag="--source-col",
            default="source",
            advanced=True,
        ),
        FieldSpec(
            "target_col",
            "Target 列",
            flag="--target-col",
            default="target",
            advanced=True,
        ),
    )


def _tag_config_field() -> FieldSpec:
    return FieldSpec(
        "tag_config",
        "Tag 配置（可选）",
        kind="open",
        flag="--tag-config",
        advanced=True,
        file_types=TOML_FILE_TYPES,
    )


SOURCE_COL_FIELD, TARGET_COL_FIELD = _source_target_fields()
CONTEXT_COL_FIELD = FieldSpec(
    "context_col",
    "Context 列（留空自动识别 context）",
    flag="--context-col",
    advanced=True,
)


TASKS: tuple[TaskSpec, ...] = (
    TaskSpec(
        "prepare",
        "准备翻译",
        "生成包含原表备份、to_translate 和 prefilled_units 的翻译工作簿。",
        "prepare",
        (
            _input("本轮源文件"),
            FieldSpec(
                "tm",
                "TM 文件（留空表示不用）",
                kind="open",
                flag="--tm",
                help_text="每次都由你明确选择，不会自动沿用上一次的 TM。",
            ),
            FieldSpec(
                "use_existing_targets",
                "使用当前 target 内容作为预填",
                kind="bool",
                flag="--use-existing-targets",
                default=True,
            ),
            _output("翻译工作簿（留空则自动命名）"),
            SOURCE_COL_FIELD,
            TARGET_COL_FIELD,
            CONTEXT_COL_FIELD,
            _tag_config_field(),
        ),
    ),
    TaskSpec(
        "fill",
        "回填译文",
        "只选择翻译工作簿，自动恢复原表格式并写入译文。",
        "fill",
        (
            _input("翻译完成的工作簿"),
            _output("回填结果（留空则自动命名）"),
            FieldSpec(
                "audit_output",
                "检查文件（留空则仅在有问题时生成）",
                kind="save",
                flag="--audit-output",
                advanced=True,
            ),
            _tag_config_field(),
        ),
    ),
    TaskSpec(
        "tm_extract",
        "生成 TM",
        "从已有 source/target 完成稿提取可复用翻译记忆。",
        "tm-extract",
        (
            _input("历史完成稿"),
            _output("TM 输出（留空则自动命名）"),
            SOURCE_COL_FIELD,
            TARGET_COL_FIELD,
            CONTEXT_COL_FIELD,
            FieldSpec(
                "min_group_size",
                "最小可复用变体数",
                kind="int",
                flag="--min-group-size",
                default="2",
                advanced=True,
            ),
            _tag_config_field(),
        ),
    ),
    TaskSpec(
        "extract_report",
        "诊断 · 生成过程包",
        "生成 summary、translation_units、source_map 和 QA 等诊断工作表。",
        "extract",
        (
            _input("源文件"),
            _output("过程包输出"),
            FieldSpec("tm", "TM 文件（可选）", kind="open", flag="--tm"),
            SOURCE_COL_FIELD,
            TARGET_COL_FIELD,
            CONTEXT_COL_FIELD,
            FieldSpec(
                "examples",
                "额外示例（每行 SOURCE=TARGET）",
                kind="lines",
                flag="--example",
            ),
            FieldSpec(
                "min_group_size",
                "最小可复用变体数",
                kind="int",
                flag="--min-group-size",
                default="2",
            ),
            FieldSpec(
                "no_existing_targets",
                "不使用当前 target 作为建议",
                kind="bool",
                flag="--no-existing-targets",
            ),
            _tag_config_field(),
        ),
        group="advanced",
    ),
    TaskSpec(
        "legacy_fill",
        "诊断 · 旧版回填 / Report",
        "为旧版非自包含 todo 回填原表，或生成 report 工作簿。",
        "fill",
        (
            _input("原始源文件"),
            FieldSpec(
                "templates",
                "旧版翻译文件",
                kind="open",
                flag="--templates",
                required=True,
            ),
            FieldSpec(
                "mode",
                "模式",
                kind="choice",
                flag="--mode",
                default="target-column",
                choices=("target-column", "report"),
            ),
            _output(),
            FieldSpec(
                "audit_output",
                "检查文件（可选）",
                kind="save",
                flag="--audit-output",
            ),
            SOURCE_COL_FIELD,
            TARGET_COL_FIELD,
            CONTEXT_COL_FIELD,
            FieldSpec(
                "min_group_size",
                "最小可复用变体数",
                kind="int",
                flag="--min-group-size",
                default="2",
            ),
            _tag_config_field(),
        ),
        group="advanced",
    ),
    TaskSpec(
        "entity_tm",
        "Entity · 生成记忆库",
        "从 TM reusable units 生成 Entity memory。",
        "entity-tm",
        (
            _input("TM reusable units"),
            _output("Entity memory 输出"),
            FieldSpec(
                "min_group_size",
                "最小结构变体数",
                kind="int",
                flag="--min-group-size",
                default="3",
            ),
        ),
        group="advanced",
    ),
    TaskSpec(
        "entity_prepare",
        "Entity · 准备处理包",
        "把 translator workbook 分成 Entity 相关与非相关单元。",
        "entity-prepare",
        (
            _input("Translator workbook"),
            FieldSpec(
                "tm",
                "Entity memory（可选）",
                kind="open",
                flag="--tm",
            ),
            _output("Entity pack 输出"),
            FieldSpec(
                "min_group_size",
                "最小结构变体数",
                kind="int",
                flag="--min-group-size",
                default="3",
            ),
        ),
        group="advanced",
    ),
    TaskSpec(
        "entity_fill_pack",
        "Entity · 组合译文",
        "把已确认的结构和实体词组合回 related_units。",
        "entity-fill-pack",
        (
            _input("Entity pack"),
            _output("Filled entity pack 输出"),
            FieldSpec(
                "in_place",
                "直接更新输入文件",
                kind="bool",
                flag="--in-place",
            ),
        ),
        group="advanced",
        exclusive_groups=(("output", "in_place"),),
    ),
    TaskSpec(
        "entity_merge_pack",
        "Entity · 合并翻译包",
        "把 related_units 和 non_related_units 合并回 translator workbook。",
        "entity-merge-pack",
        (_input("Filled entity pack"), _output("Merged translator workbook")),
        group="advanced",
    ),
    TaskSpec(
        "entity_split",
        "Entity 调试 · 拆分",
        "拆分 Entity 相关和非相关工作簿。",
        "entity-split",
        (
            _input("Translator workbook"),
            FieldSpec(
                "entity_output",
                "Entity 相关输出（可选）",
                kind="save",
                flag="--entity-output",
            ),
            FieldSpec(
                "non_entity_output",
                "非 Entity 输出（可选）",
                kind="save",
                flag="--non-entity-output",
            ),
            FieldSpec(
                "min_group_size",
                "最小结构变体数",
                kind="int",
                flag="--min-group-size",
                default="3",
            ),
        ),
        group="advanced",
    ),
    TaskSpec(
        "entity_prefill",
        "Entity 调试 · 预填",
        "用 Entity TM 预填结构和实体词。",
        "entity-prefill",
        (
            _input("Entity 相关工作簿"),
            FieldSpec("tm", "Entity TM", kind="open", flag="--tm", required=True),
            _output("Prefilled entity 输出"),
        ),
        group="advanced",
    ),
    TaskSpec(
        "entity_extract_tm",
        "Entity 调试 · 提取 TM",
        "从 TM pairs 提取 Entity 结构和词条。",
        "entity-extract-tm",
        (
            _input("TM reusable units"),
            _output("Entity TM 输出"),
            FieldSpec(
                "min_group_size",
                "最小结构变体数",
                kind="int",
                flag="--min-group-size",
                default="3",
            ),
        ),
        group="advanced",
    ),
    TaskSpec(
        "entity_fill",
        "Entity 调试 · 回填",
        "把可用 Entity 译文写回 Entity 相关工作簿。",
        "entity-fill",
        (_input("Entity 相关工作簿"), _output("Filled entity 输出")),
        group="advanced",
    ),
    TaskSpec(
        "entity_merge",
        "Entity 调试 · 合并",
        "合并 Entity 相关与非相关工作簿。",
        "entity-merge",
        (
            FieldSpec(
                "entity",
                "Filled entity 工作簿",
                kind="open",
                flag="--entity",
                required=True,
            ),
            FieldSpec(
                "non_entity",
                "非 Entity 工作簿",
                kind="open",
                flag="--non-entity",
                required=True,
            ),
            _output("Merged todo 输出"),
        ),
        group="advanced",
    ),
)


TASK_BY_KEY = {task.key: task for task in TASKS}
DAILY_TASKS = tuple(task for task in TASKS if task.group == "daily")
ADVANCED_TASKS = tuple(task for task in TASKS if task.group == "advanced")


def build_cli_args(task: TaskSpec, values: Mapping[str, object]) -> list[str]:
    args = [task.command]
    missing: list[str] = []
    for field_spec in task.fields:
        value = values.get(field_spec.key, field_spec.default)
        if field_spec.kind == "bool":
            enabled = bool(value)
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
        if field_spec.kind == "choice" and text not in field_spec.choices:
            raise ValueError(f"{field_spec.label}的值无效：{text}")
        if field_spec.kind == "lines":
            for line in (line.strip() for line in text.splitlines()):
                if line:
                    if field_spec.flag:
                        args.extend([field_spec.flag, line])
            continue
        if field_spec.flag:
            args.extend([field_spec.flag, text])
        else:
            args.append(text)

    if missing:
        raise ValueError("请填写：" + "、".join(missing))
    for exclusive_group in task.exclusive_groups:
        selected = [key for key in exclusive_group if values.get(key)]
        if len(selected) > 1:
            labels = [
                next(field.label for field in task.fields if field.key == key)
                for key in selected
            ]
            raise ValueError("不能同时使用：" + "、".join(labels))
    return args


def validate_task_specs() -> None:
    keys = [task.key for task in TASKS]
    if len(keys) != len(set(keys)):
        raise ValueError("GUI task keys must be unique")
    for task in TASKS:
        field_keys = [field.key for field in task.fields]
        if len(field_keys) != len(set(field_keys)):
            raise ValueError(f"GUI field keys must be unique: {task.key}")
        if not any(field.required for field in task.fields):
            raise ValueError(f"GUI task needs a required input: {task.key}")


class PhraseLoomGUI:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("PhraseLoom")
        self.root.geometry("820x680")
        self.root.minsize(700, 580)
        self.root.option_add("*Font", ("Arial", 12))

        self.current_task_key = DAILY_TASKS[0].key
        self.task_values = {
            task.key: {field.key: field.default for field in task.fields}
            for task in TASKS
        }
        self.field_vars: dict[str, tk.Variable | tk.Text] = {}
        self.result_queue: queue.Queue[tuple[int, str, str]] = queue.Queue()
        self.show_more_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="请选择任务和文件")
        self.advanced_choice_var = tk.StringVar(value=ADVANCED_TASKS[0].label)

        self._configure_styles()
        self._build_layout()
        self._render_task()

    def _configure_styles(self) -> None:
        style = ttk.Style(self.root)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("Title.TLabel", font=("Arial", 24, "bold"))
        style.configure("Subtitle.TLabel", foreground="#5B6472")
        style.configure("Task.TButton", padding=(16, 10))
        style.configure("Primary.TButton", padding=(18, 11), font=("Arial", 12, "bold"))
        style.configure(
            "Disclosure.TButton",
            padding=(0, 8),
            foreground="#475467",
            anchor="w",
        )
        style.configure("Hint.TLabel", foreground="#667085", font=("Arial", 10))
        style.configure("Status.TLabel", foreground="#344054")

    def _build_layout(self) -> None:
        outer = ttk.Frame(self.root, padding=24)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="PhraseLoom", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            outer,
            text="Excel 本地化工作流",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(2, 18))

        self.notebook = ttk.Notebook(outer)
        self.notebook.pack(fill="both", expand=True)

        self.daily_tab = ttk.Frame(self.notebook, padding=16)
        self.advanced_tab = ttk.Frame(self.notebook, padding=16)
        self.notebook.add(self.daily_tab, text="日常流程")
        self.notebook.add(self.advanced_tab, text="高级工具")

        daily_nav = ttk.Frame(self.daily_tab)
        daily_nav.pack(fill="x", pady=(0, 14))
        for task in DAILY_TASKS:
            ttk.Button(
                daily_nav,
                text=task.label,
                style="Task.TButton",
                command=lambda key=task.key: self._select_task(key),
            ).pack(side="left", padx=(0, 8))

        advanced_nav = ttk.Frame(self.advanced_tab)
        advanced_nav.pack(fill="x", pady=(0, 14))
        ttk.Label(advanced_nav, text="任务").pack(side="left", padx=(0, 10))
        advanced_combo = ttk.Combobox(
            advanced_nav,
            textvariable=self.advanced_choice_var,
            values=[task.label for task in ADVANCED_TASKS],
            state="readonly",
            width=34,
        )
        advanced_combo.pack(side="left", fill="x", expand=True)
        advanced_combo.bind("<<ComboboxSelected>>", self._on_advanced_selected)

        self.daily_form_host = ttk.Frame(self.daily_tab)
        self.daily_form_host.pack(fill="both", expand=True)
        self.advanced_form_host = ttk.Frame(self.advanced_tab)
        self.advanced_form_host.pack(fill="both", expand=True)
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        footer = ttk.Frame(outer)
        footer.pack(fill="x", pady=(16, 0))
        self.progress = ttk.Progressbar(footer, mode="indeterminate", length=150)
        self.progress.pack(side="left")
        ttk.Label(footer, textvariable=self.status_var, style="Status.TLabel").pack(
            side="left", padx=12
        )
        self.run_button = ttk.Button(
            footer,
            text="开始",
            style="Primary.TButton",
            command=self._run_current_task,
        )
        self.run_button.pack(side="right")

        self.output = tk.Text(
            outer,
            height=7,
            wrap="word",
            borderwidth=0,
            background="#F4F6F8",
            foreground="#263238",
            padx=12,
            pady=10,
        )
        self.output.pack(fill="x", pady=(14, 0))
        self.output.configure(state="disabled")

    def _on_tab_changed(self, _event=None) -> None:
        self._capture_values()
        selected = self.notebook.index(self.notebook.select())
        if selected == 0:
            if TASK_BY_KEY[self.current_task_key].group != "daily":
                self.current_task_key = DAILY_TASKS[0].key
        else:
            chosen = next(
                task for task in ADVANCED_TASKS if task.label == self.advanced_choice_var.get()
            )
            self.current_task_key = chosen.key
        self.show_more_var.set(False)
        self._render_task()

    def _on_advanced_selected(self, _event=None) -> None:
        chosen = next(
            task for task in ADVANCED_TASKS if task.label == self.advanced_choice_var.get()
        )
        self._select_task(chosen.key)

    def _select_task(self, task_key: str) -> None:
        self._capture_values()
        self.current_task_key = task_key
        self.show_more_var.set(False)
        self._render_task()

    def _capture_values(self) -> None:
        if not self.field_vars:
            return
        values = self.task_values[self.current_task_key]
        for key, variable in self.field_vars.items():
            if isinstance(variable, tk.Text):
                values[key] = variable.get("1.0", "end").strip()
            else:
                values[key] = variable.get()

    def _render_task(self) -> None:
        task = TASK_BY_KEY[self.current_task_key]
        host = self.daily_form_host if task.group == "daily" else self.advanced_form_host
        for child in host.winfo_children():
            child.destroy()
        self.field_vars = {}

        ttk.Label(host, text=task.label, font=("Arial", 18, "bold")).pack(anchor="w")
        ttk.Label(host, text=task.description, style="Subtitle.TLabel").pack(
            anchor="w", pady=(3, 15)
        )

        form = ttk.Frame(host)
        form.pack(fill="x")
        visible_fields = [
            item
            for item in task.fields
            if not item.advanced or task.group == "advanced" or self.show_more_var.get()
        ]
        row_index = 0
        for field_spec in visible_fields:
            self._render_field(form, row_index, field_spec)
            row_index += 2 if field_spec.help_text and field_spec.kind != "bool" else 1
        form.columnconfigure(1, weight=1)

        if task.group == "daily" and any(item.advanced for item in task.fields):
            expanded = self.show_more_var.get()
            ttk.Button(
                host,
                text="▾ 收起更多选项" if expanded else "▸ 更多选项",
                style="Disclosure.TButton",
                command=self._toggle_more,
            ).pack(anchor="w", pady=(12, 0))

    def _render_field(self, parent: ttk.Frame, row: int, field_spec: FieldSpec) -> None:
        value = self.task_values[self.current_task_key].get(
            field_spec.key, field_spec.default
        )
        if field_spec.kind == "bool":
            variable = tk.BooleanVar(value=bool(value))
            ttk.Checkbutton(parent, text=field_spec.label, variable=variable).grid(
                row=row, column=0, columnspan=3, sticky="w", pady=7
            )
            self.field_vars[field_spec.key] = variable
            return

        ttk.Label(parent, text=field_spec.label).grid(
            row=row, column=0, sticky="w", padx=(0, 12), pady=6
        )
        if field_spec.kind == "lines":
            widget = tk.Text(parent, height=3, wrap="word")
            widget.insert("1.0", str(value or ""))
            widget.grid(row=row, column=1, columnspan=2, sticky="ew", pady=6)
            self.field_vars[field_spec.key] = widget
            return

        variable = tk.StringVar(value=str(value or ""))
        if field_spec.kind == "choice":
            widget = ttk.Combobox(
                parent,
                textvariable=variable,
                values=field_spec.choices,
                state="readonly",
            )
        else:
            widget = ttk.Entry(parent, textvariable=variable)
        widget.grid(row=row, column=1, sticky="ew", pady=6)
        self.field_vars[field_spec.key] = variable

        if field_spec.kind in {"open", "save"}:
            ttk.Button(
                parent,
                text="选择…",
                command=lambda spec=field_spec, var=variable: self._choose_path(spec, var),
            ).grid(row=row, column=2, padx=(8, 0), pady=6)
        if field_spec.help_text:
            ttk.Label(parent, text=field_spec.help_text, style="Hint.TLabel").grid(
                row=row + 1,
                column=1,
                columnspan=2,
                sticky="w",
                pady=(0, 4),
            )

    def _toggle_more(self) -> None:
        self._capture_values()
        self.show_more_var.set(not self.show_more_var.get())
        self._render_task()

    def _choose_path(self, field_spec: FieldSpec, variable: tk.StringVar) -> None:
        if field_spec.kind == "save":
            value = filedialog.asksaveasfilename(
                parent=self.root,
                filetypes=field_spec.file_types,
                defaultextension=".xlsx",
            )
        else:
            value = filedialog.askopenfilename(
                parent=self.root,
                filetypes=field_spec.file_types,
            )
        if value:
            variable.set(value)

    def _run_current_task(self) -> None:
        self._capture_values()
        task = TASK_BY_KEY[self.current_task_key]
        try:
            args = build_cli_args(task, self.task_values[task.key])
        except ValueError as exc:
            messagebox.showerror("无法开始", str(exc), parent=self.root)
            return

        self.run_button.configure(state="disabled")
        self.progress.start(10)
        self.status_var.set("处理中…")
        self._set_output("")
        worker = threading.Thread(
            target=self._execute_cli,
            args=(args,),
            daemon=True,
        )
        worker.start()
        self.root.after(100, self._poll_result)

    def _execute_cli(self, args: list[str]) -> None:
        from .cli import main as cli_main

        stdout = io.StringIO()
        stderr = io.StringIO()
        try:
            with contextlib.redirect_stdout(stdout), contextlib.redirect_stderr(stderr):
                exit_code = cli_main(args)
        except SystemExit as exc:
            exit_code = int(exc.code or 0)
        except Exception as exc:  # GUI boundary: surface unexpected failures cleanly.
            exit_code = 1
            stderr.write(f"{type(exc).__name__}: {exc}\n")
        output = stdout.getvalue()
        error = stderr.getvalue()
        self.result_queue.put((exit_code, output, error))

    def _poll_result(self) -> None:
        try:
            result = self.result_queue.get_nowait()
        except queue.Empty:
            self.root.after(100, self._poll_result)
            return
        self._finish_run(*result)

    def _finish_run(self, exit_code: int, output: str, error: str) -> None:
        self.progress.stop()
        self.run_button.configure(state="normal")
        combined = "\n".join(part.strip() for part in (output, error) if part.strip())
        self._set_output(combined)
        if exit_code == 0:
            self.status_var.set("完成")
        else:
            self.status_var.set("未完成，请查看下方信息")
            messagebox.showerror(
                "任务未完成",
                error.strip() or output.strip() or "未知错误",
                parent=self.root,
            )

    def _set_output(self, text: str) -> None:
        self.output.configure(state="normal")
        self.output.delete("1.0", "end")
        if text:
            self.output.insert("1.0", text)
        self.output.configure(state="disabled")


def main() -> int:
    validate_task_specs()
    try:
        root = tk.Tk()
    except tk.TclError as exc:
        print(f"Could not start PhraseLoom GUI: {exc}")
        return 1
    PhraseLoomGUI(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
