"""Reusable Tk controls for shared history-TB project presets."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk

from tools.tb_projects import TbProject, TbProjectStore


CaptureProject = Callable[[str], TbProject]
ApplyProject = Callable[[TbProject], None]


class TbProjectControls(ttk.Frame):
    def __init__(
        self,
        master: ttk.Frame,
        *,
        capture_project: CaptureProject,
        apply_project: ApplyProject,
        store: TbProjectStore | None = None,
    ) -> None:
        super().__init__(master)
        self.capture_project = capture_project
        self.apply_project = apply_project
        self.store = store or TbProjectStore()
        self.project_name_var = tk.StringVar()
        self.status_var = tk.StringVar()

        ttk.Label(self, text="TB 项目").grid(row=0, column=0, sticky="w")
        self.project_combobox = ttk.Combobox(
            self,
            textvariable=self.project_name_var,
            width=24,
            state="readonly",
        )
        self.project_combobox.grid(row=0, column=1, sticky="w", padx=(12, 8))
        self.project_combobox.bind(
            "<<ComboboxSelected>>",
            self.handle_project_selected,
        )
        self.project_combobox.bind(
            "<Button-1>",
            lambda _event: self.refresh_projects(show_error=False),
        )
        ttk.Button(
            self,
            text="保存当前",
            command=self.save_current_project,
        ).grid(row=0, column=2, sticky="w")
        ttk.Button(
            self,
            text="删除",
            command=self.delete_selected_project,
        ).grid(row=0, column=3, sticky="w", padx=(6, 0))
        ttk.Label(
            self,
            textvariable=self.status_var,
        ).grid(row=0, column=4, sticky="w", padx=(10, 0))

        self.refresh_projects(show_error=False)

    def refresh_projects(
        self,
        *,
        selected_name: str | None = None,
        show_error: bool = True,
    ) -> None:
        try:
            projects = self.store.list_projects()
        except ValueError as exc:
            self.project_combobox["values"] = ()
            if show_error:
                messagebox.showerror("TB 项目读取失败", str(exc))
            return

        project_names = tuple(project.name for project in projects)
        self.project_combobox["values"] = project_names
        requested_name = (
            selected_name
            if selected_name is not None
            else self.project_name_var.get().strip()
        )
        if requested_name in project_names:
            self.project_name_var.set(requested_name)
        else:
            self.project_name_var.set("")

    def clear_selection(self) -> None:
        self.project_name_var.set("")
        self.status_var.set("")

    def mark_current_settings_modified(self) -> None:
        if self.project_name_var.get().strip():
            self.status_var.set("当前设置已修改")

    def handle_project_selected(self, _event: object | None = None) -> None:
        project_name = self.project_name_var.get().strip()
        if not project_name:
            return
        try:
            project = self.store.find_project(project_name)
        except ValueError as exc:
            messagebox.showerror("TB 项目读取失败", str(exc))
            return
        if project is None:
            self.refresh_projects()
            return

        self.apply_project(project)
        self.status_var.set(f"已载入：{project.name}")
        if not Path(project.file_path).is_file():
            messagebox.showwarning(
                "TB 文件不存在",
                (
                    f"项目“{project.name}”对应的 TB 文件已移动或不存在。\n"
                    "请重新选择历史 TB，然后用相同项目名保存更新。"
                ),
            )

    def save_current_project(self) -> None:
        selected_name = self.project_name_var.get().strip()
        project_name = simpledialog.askstring(
            "保存 TB 项目",
            "项目名称：",
            initialvalue=selected_name,
            parent=self.winfo_toplevel(),
        )
        if project_name is None:
            return
        project_name = project_name.strip()
        if not project_name:
            messagebox.showerror("项目名称为空", "请输入项目名称。")
            return

        try:
            project = self.capture_project(project_name)
            existing_project = self.store.find_project(project_name)
        except (OSError, ValueError) as exc:
            messagebox.showerror("无法保存 TB 项目", str(exc))
            return

        if existing_project is not None and not messagebox.askyesno(
            "更新 TB 项目",
            f"项目“{existing_project.name}”已存在，是否用当前设置更新？",
        ):
            return

        try:
            self.store.save_project(project)
        except (OSError, ValueError) as exc:
            messagebox.showerror("无法保存 TB 项目", str(exc))
            return
        self.refresh_projects(selected_name=project.name)
        self.status_var.set(f"已保存：{project.name}")

    def delete_selected_project(self) -> None:
        project_name = self.project_name_var.get().strip()
        if not project_name:
            messagebox.showinfo("未选择项目", "请先选择要删除的 TB 项目。")
            return
        if not messagebox.askyesno(
            "删除 TB 项目",
            f"确定删除项目“{project_name}”吗？\n不会删除原始 TB 文件。",
        ):
            return
        try:
            self.store.delete_project(project_name)
        except (OSError, ValueError) as exc:
            messagebox.showerror("无法删除 TB 项目", str(exc))
            return
        self.refresh_projects(selected_name="")
        self.status_var.set(f"已删除：{project_name}")
