#!/usr/bin/env python3
"""Unified GUI launcher for all Excel tools in this repository."""

from __future__ import annotations

import argparse
from collections.abc import Callable
from dataclasses import dataclass
import sys
import tkinter as tk
from tkinter import messagebox, ttk

from phraseloom.gui import PhraseLoomApp
from tools.gui_common import configure_tool_page_style
from tools.chinese_target_checker.check_chinese_target_gui import ChineseTargetCheckerApp
from tools.french_nbsp_restorer.restore_french_nbsp_gui import FrenchNbspRestorerApp
from tools.line_break_checker.check_line_breaks_gui import LineBreakCheckerApp
from tools.source_consistency_checker.check_source_consistency_gui import SourceConsistencyCheckerApp
from tools.tag_placeholder_checker.check_tags_and_placeholders_gui import TagPlaceholderCheckerApp
from tools.term_pair_checker.extract_terms_gui import ExtractTermsApp
from tools.workflow.file_receiver import (
    FRENCH_NBSP_RESTORE_ACTION,
    QA_WORKFLOW_ACTION,
    ToolFileRequest,
    WorkflowFileReceiver,
    normalize_excel_input_file,
    send_tool_input_file,
)
from tools.workflow.workflow_gui import WorkflowRunnerApp
from tools.xbench_report_transformer.transform_xbench_report_gui import XbenchReportTransformerApp


ToolFactory = Callable[[tk.Misc], ttk.Frame]


@dataclass(frozen=True)
class ToolItem:
    key: str
    title: str
    description: str
    factory: ToolFactory


@dataclass(frozen=True)
class ToolGroup:
    title: str
    tools: tuple[ToolItem, ...]


TOOL_GROUPS = (
    ToolGroup(
        title="常用流程",
        tools=(
            ToolItem(
                key="workflow",
                title="一键质量检查",
                description="按顺序执行质量检查板块的全部项目，统一写入输出 Excel。",
                factory=WorkflowRunnerApp,
            ),
            ToolItem(
                key="phraseloom",
                title="PhraseLoom",
                description="导出干净、去重的 Strings 工作簿，并在翻译后回填原始 Excel。",
                factory=PhraseLoomApp,
            ),
        ),
    ),
    ToolGroup(
        title="质量检查",
        tools=(
            ToolItem(
                key="term_pair",
                title="术语检查",
                description="从可选 mark 提取新术语对，或仅用历史 TB 检查 source / target 命中。",
                factory=ExtractTermsApp,
            ),
            ToolItem(
                key="tag_checker",
                title="Tag 检查",
                description="检查 tag、placeholder、换行标记和数字 tag 在 source / target 中是否一致。",
                factory=TagPlaceholderCheckerApp,
            ),
            ToolItem(
                key="line_break_checker",
                title="换行数量检查",
                description="逐行比较 source / target 单元格中的真实换行数量是否一致。",
                factory=LineBreakCheckerApp,
            ),
            ToolItem(
                key="source_consistency_checker",
                title="同源译文一致性",
                description="检查完全相同的 source 是否对应多个不同 target。",
                factory=SourceConsistencyCheckerApp,
            ),
            ToolItem(
                key="chinese_target",
                title="Target 中文检查",
                description="扫描 target 文本中的中文字符，定位未翻译或混入中文的问题。",
                factory=ChineseTargetCheckerApp,
            ),
        ),
    ),
    ToolGroup(
        title="文本修复",
        tools=(
            ToolItem(
                key="french_nbsp",
                title="法语 NBSP 恢复",
                description="恢复法语标点前的 NBSP，修正常见空格丢失问题。",
                factory=FrenchNbspRestorerApp,
            ),
        ),
    ),
    ToolGroup(
        title="其他",
        tools=(
            ToolItem(
                key="xbench_report",
                title="Xbench QA 转换",
                description="把 Xbench QA Report 整理为按 key/source 聚合的行级 Excel。",
                factory=XbenchReportTransformerApp,
            ),
        ),
    ),
)


class ToolshubApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.withdraw()
        self.root.title("Toolshub")
        self.root.resizable(True, True)
        self.tool_groups = TOOL_GROUPS
        self.tools_by_key = {
            tool.key: tool
            for group in self.tool_groups
            for tool in group.tools
        }
        self.selected_tool_key = tk.StringVar()
        self.current_tool_title = tk.StringVar()
        self.current_tool_description = tk.StringVar()
        self.tool_frames: dict[str, ttk.Frame] = {}
        self.nav_buttons: dict[str, ttk.Radiobutton] = {}
        self.current_tool_frame: ttk.Frame | None = None
        self._build_ui()
        self._fit_window_to_content()

    def _build_ui(self) -> None:
        self._configure_style()
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        shell = ttk.Frame(self.root, padding=16)
        shell.grid(row=0, column=0, sticky="nsew")
        shell.columnconfigure(1, weight=1)
        shell.rowconfigure(0, weight=1)

        sidebar = ttk.Frame(shell, padding=(0, 0, 16, 0))
        sidebar.grid(row=0, column=0, sticky="nsw")
        self._build_sidebar(sidebar)

        workspace = ttk.Frame(shell)
        workspace.grid(row=0, column=1, sticky="nsew")
        workspace.columnconfigure(0, weight=1)
        workspace.rowconfigure(3, weight=1)

        ttk.Label(
            workspace,
            textvariable=self.current_tool_title,
            style="Toolshub.Title.TLabel",
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            workspace,
            textvariable=self.current_tool_description,
            style="Toolshub.Description.TLabel",
            wraplength=760,
        ).grid(row=1, column=0, sticky="ew", pady=(4, 12))
        ttk.Separator(workspace).grid(row=2, column=0, sticky="ew", pady=(0, 12))

        self.content_frame = ttk.Frame(workspace)
        self.content_frame.grid(row=3, column=0, sticky="nsew")
        self.content_frame.columnconfigure(0, weight=1)
        self.content_frame.rowconfigure(0, weight=1)

        first_tool = self.tool_groups[0].tools[0]
        self.select_tool(first_tool.key)

    def _configure_style(self) -> None:
        configure_tool_page_style(self.root)
        style = ttk.Style(self.root)
        style.configure("Toolshub.AppTitle.TLabel", font=("TkDefaultFont", 15, "bold"))
        style.configure("Toolshub.Category.TLabel", font=("TkDefaultFont", 10, "bold"))
        style.configure("Toolshub.Title.TLabel", font=("TkDefaultFont", 16, "bold"))
        style.configure("Toolshub.Description.TLabel", foreground="#555555")
        style.configure("Toolshub.Nav.TRadiobutton", padding=(10, 6))

    def _build_sidebar(self, parent: ttk.Frame) -> None:
        ttk.Label(
            parent,
            text="QA 工具箱",
            style="Toolshub.AppTitle.TLabel",
        ).grid(row=0, column=0, sticky="w", pady=(0, 4))
        ttk.Label(
            parent,
            text="Toolshub",
            style="Toolshub.Description.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(0, 16))

        row = 2
        for group in self.tool_groups:
            ttk.Label(
                parent,
                text=group.title,
                style="Toolshub.Category.TLabel",
            ).grid(row=row, column=0, sticky="w", pady=(12 if row > 2 else 0, 4))
            row += 1
            for tool in group.tools:
                button = ttk.Radiobutton(
                    parent,
                    text=tool.title,
                    value=tool.key,
                    variable=self.selected_tool_key,
                    command=lambda key=tool.key: self.select_tool(key),
                    style="Toolshub.Nav.TRadiobutton",
                )
                button.grid(row=row, column=0, sticky="ew", pady=1)
                self.nav_buttons[tool.key] = button
                row += 1

        parent.columnconfigure(0, minsize=170)

    def _get_or_create_tool_frame(self, key: str) -> ttk.Frame:
        frame = self.tool_frames.get(key)
        if frame is None:
            tool = self.tools_by_key[key]
            frame = tool.factory(self.content_frame)
            self.tool_frames[key] = frame
        return frame

    def select_tool(self, key: str) -> None:
        tool = self.tools_by_key[key]
        frame = self._get_or_create_tool_frame(key)
        if self.current_tool_frame is not None and self.current_tool_frame is not frame:
            self.current_tool_frame.grid_remove()

        frame.grid(row=0, column=0, sticky="nsew")
        self.selected_tool_key.set(key)
        self.current_tool_title.set(tool.title)
        self.current_tool_description.set(tool.description)
        self.current_tool_frame = frame
        frame.tkraise()

    def open_qa_workflow_file(self, file_path: str) -> None:
        """Select the workflow page and load an Excel file from Finder."""

        normalized_path = normalize_excel_input_file(
            file_path,
            action_name="QA workflow",
        )
        workflow_frame = self._get_or_create_tool_frame("workflow")
        if not isinstance(workflow_frame, WorkflowRunnerApp):
            raise RuntimeError("一键质量检查页面未正确加载。")

        self.select_tool("workflow")
        workflow_frame.load_input_file(str(normalized_path))
        self._bring_window_to_front()

    def open_french_nbsp_restore_file(
        self,
        file_path: str,
        *,
        run_immediately: bool = True,
    ) -> None:
        """Load a Finder file into French NBSP restore and optionally run it."""

        normalized_path = normalize_excel_input_file(
            file_path,
            action_name="NBSP restore",
        )
        restorer_frame = self._get_or_create_tool_frame("french_nbsp")
        if not isinstance(restorer_frame, FrenchNbspRestorerApp):
            raise RuntimeError("法语 NBSP 恢复页面未正确加载。")

        self.select_tool("french_nbsp")
        restorer_frame.load_input_file(
            str(normalized_path),
            reset_options=True,
        )
        self._bring_window_to_front()
        if run_immediately:
            self.root.after_idle(restorer_frame.run_restore)

    def _bring_window_to_front(self) -> None:
        self.root.deiconify()
        self.root.lift()
        try:
            self.root.attributes("-topmost", True)
            self.root.after(250, lambda: self.root.attributes("-topmost", False))
            self.root.focus_force()
        except tk.TclError:
            pass

    def _fit_window_to_content(self) -> None:
        self.root.update_idletasks()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        width = min(max(self.root.winfo_reqwidth(), 1080), max(screen_width - 80, 900))
        height = min(max(self.root.winfo_reqheight(), 720), max(screen_height - 80, 640))
        x = max((self.root.winfo_screenwidth() - width) // 2, 0)
        y = max((self.root.winfo_screenheight() - height) // 2, 30)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        self.root.minsize(min(width, 980), min(height, 680))
        self.root.deiconify()


def build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="打开 Toolshub Excel 工具箱。")
    finder_action_group = parser.add_mutually_exclusive_group()
    finder_action_group.add_argument(
        "--qa-workflow",
        metavar="EXCEL_FILE",
        help="把 Excel 文件载入一键质量检查页面。",
    )
    finder_action_group.add_argument(
        "--nbsp-restore",
        metavar="EXCEL_FILE",
        help="对 Excel 自动执行法语 NBSP 恢复。",
    )
    parser.add_argument(
        "--smoke-test",
        action="store_true",
        help=argparse.SUPPRESS,
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_argument_parser().parse_args(argv)
    if args.smoke_test:
        root = tk.Tk()
        try:
            root.withdraw()
            root.update_idletasks()
        finally:
            root.destroy()
        return 0

    initial_request: ToolFileRequest | None = None
    if args.qa_workflow:
        try:
            initial_request = ToolFileRequest(
                action=QA_WORKFLOW_ACTION,
                file_path=str(
                    normalize_excel_input_file(
                        args.qa_workflow,
                        action_name="QA workflow",
                    )
                ),
            )
        except ValueError as exc:
            print(str(exc), file=sys.stderr)
            return 2
    elif args.nbsp_restore:
        try:
            initial_request = ToolFileRequest(
                action=FRENCH_NBSP_RESTORE_ACTION,
                file_path=str(
                    normalize_excel_input_file(
                        args.nbsp_restore,
                        action_name="NBSP restore",
                    )
                ),
            )
        except ValueError as exc:
            print(str(exc), file=sys.stderr)
            return 2

    if initial_request and send_tool_input_file(
        initial_request.action,
        initial_request.file_path,
    ):
        return 0

    receiver = WorkflowFileReceiver()
    receiver_started = receiver.start()
    if initial_request and not receiver_started:
        if send_tool_input_file(
            initial_request.action,
            initial_request.file_path,
        ):
            return 0

    root = tk.Tk()
    app = ToolshubApp(root)

    def handle_file_request(request: ToolFileRequest) -> None:
        if request.action == FRENCH_NBSP_RESTORE_ACTION:
            app.open_french_nbsp_restore_file(request.file_path)
        else:
            app.open_qa_workflow_file(request.file_path)

    if initial_request:
        try:
            handle_file_request(initial_request)
        except Exception as exc:
            messagebox.showerror("无法载入 Excel", str(exc))

    poll_job: str | None = None

    def poll_forwarded_files() -> None:
        nonlocal poll_job
        for forwarded_request in receiver.pop_pending_requests():
            try:
                handle_file_request(forwarded_request)
            except Exception as exc:
                messagebox.showerror("无法载入 Excel", str(exc))
        poll_job = root.after(150, poll_forwarded_files)

    def close_app() -> None:
        if poll_job is not None:
            root.after_cancel(poll_job)
        receiver.close()
        root.destroy()

    if receiver_started:
        poll_forwarded_files()
    root.protocol("WM_DELETE_WINDOW", close_app)
    try:
        root.mainloop()
    finally:
        receiver.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
