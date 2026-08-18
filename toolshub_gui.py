#!/usr/bin/env python3
"""Unified GUI launcher for all Excel tools in this repository."""

from __future__ import annotations

import argparse
from collections.abc import Callable
import ctypes
from dataclasses import dataclass
import sys
import tkinter as tk
from tkinter import font as tkfont
from tkinter import messagebox, ttk

from phraseloom.gui import PhraseLoomApp
from tools.gui_common import (
    APP_MAIN_BACKGROUND,
    APP_MUTED_TEXT,
    APP_SIDEBAR_BACKGROUND,
    APP_TEXT,
    configure_tool_page_style,
)
from tools.french_nbsp_restorer.restore_french_nbsp_gui import FrenchNbspRestorerApp
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

DEFAULT_WINDOW_WIDTH = 1200
DEFAULT_WINDOW_HEIGHT = 800
WINDOW_HORIZONTAL_BREATHING_ROOM = 96
WINDOW_VERTICAL_BREATHING_ROOM = 72


def enable_high_dpi_awareness() -> None:
    """Let Tk render at the monitor's real DPI instead of bitmap scaling."""

    if sys.platform != "win32":
        return
    try:
        if ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4)):
            return
    except (AttributeError, OSError):
        pass
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except (AttributeError, OSError):
        pass


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
                description="集中执行全部质量检查项目，并统一写入输出 Excel。",
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

        shell = ttk.Frame(self.root, style="Toolshub.Shell.TFrame")
        shell.grid(row=0, column=0, sticky="nsew")
        shell.columnconfigure(1, weight=1)
        shell.rowconfigure(0, weight=1)

        sidebar = ttk.Frame(
            shell,
            padding=(16, 20, 14, 16),
            style="Toolshub.Sidebar.TFrame",
        )
        sidebar.grid(row=0, column=0, sticky="nsw")
        self._build_sidebar(sidebar)
        ttk.Separator(shell, orient="vertical").grid(
            row=0,
            column=0,
            sticky="nse",
        )

        workspace = ttk.Frame(
            shell,
            padding=(24, 22, 28, 18),
            style="Toolshub.Workspace.TFrame",
        )
        workspace.grid(row=0, column=1, sticky="nsew")
        workspace.columnconfigure(0, weight=1)
        workspace.rowconfigure(3, weight=1)

        ttk.Label(
            workspace,
            textvariable=self.current_tool_title,
            style="Toolshub.Title.TLabel",
        ).grid(row=0, column=0, sticky="w", padx=16)
        ttk.Label(
            workspace,
            textvariable=self.current_tool_description,
            style="Toolshub.Description.TLabel",
            wraplength=760,
        ).grid(row=1, column=0, sticky="ew", padx=16, pady=(5, 16))
        ttk.Separator(workspace).grid(
            row=2,
            column=0,
            sticky="ew",
            padx=16,
            pady=(0, 2),
        )

        self.content_frame = ttk.Frame(workspace)
        self.content_frame.grid(row=3, column=0, sticky="nsew")
        self.content_frame.columnconfigure(0, weight=1)
        self.content_frame.rowconfigure(0, weight=1)

        self._build_tool_pages()
        first_tool = self.tool_groups[0].tools[0]
        self.select_tool(first_tool.key)

    def _configure_style(self) -> None:
        configure_tool_page_style(self.root)
        style = ttk.Style(self.root)
        default_font = tkfont.nametofont("TkDefaultFont", root=self.root)
        family = str(default_font.actual("family"))
        native_size = abs(int(default_font.actual("size"))) or 10
        body_size = max(native_size, 10) if sys.platform == "win32" else native_size

        style.configure("Toolshub.Shell.TFrame", background=APP_MAIN_BACKGROUND)
        style.configure("Toolshub.Sidebar.TFrame", background=APP_SIDEBAR_BACKGROUND)
        style.configure("Toolshub.Workspace.TFrame", background=APP_MAIN_BACKGROUND)
        style.configure(
            "Toolshub.AppTitle.TLabel",
            background=APP_SIDEBAR_BACKGROUND,
            foreground=APP_TEXT,
            font=(family, body_size + 3, "bold"),
        )
        style.configure(
            "Toolshub.AppSubtitle.TLabel",
            background=APP_SIDEBAR_BACKGROUND,
            foreground=APP_MUTED_TEXT,
            font=(family, max(body_size - 1, 8)),
        )
        style.configure(
            "Toolshub.Category.TLabel",
            background=APP_SIDEBAR_BACKGROUND,
            foreground=APP_MUTED_TEXT,
            font=(family, max(body_size - 1, 8), "bold"),
        )
        style.configure(
            "Toolshub.Title.TLabel",
            background=APP_MAIN_BACKGROUND,
            foreground=APP_TEXT,
            font=(family, body_size + 8, "bold"),
        )
        style.configure(
            "Toolshub.Description.TLabel",
            background=APP_MAIN_BACKGROUND,
            foreground=APP_MUTED_TEXT,
            font=(family, body_size),
        )
        style.layout(
            "Toolshub.Nav.TRadiobutton",
            style.layout("Toggle.TButton"),
        )
        style.configure(
            "Toolshub.Nav.TRadiobutton",
            anchor="w",
            font=(family, body_size),
            padding=(11, 8),
        )
        style.map(
            "Toolshub.Nav.TRadiobutton",
            foreground=style.map("Toggle.TButton", "foreground"),
        )

    def _build_sidebar(self, parent: ttk.Frame) -> None:
        ttk.Label(
            parent,
            text="QAtools",
            style="Toolshub.AppTitle.TLabel",
        ).grid(row=0, column=0, sticky="w", padx=10, pady=(0, 3))
        ttk.Label(
            parent,
            text="本地化 QA 工作台",
            style="Toolshub.AppSubtitle.TLabel",
        ).grid(row=1, column=0, sticky="w", padx=10, pady=(0, 20))

        row = 2
        for group in self.tool_groups:
            ttk.Label(
                parent,
                text=group.title,
                style="Toolshub.Category.TLabel",
            ).grid(
                row=row,
                column=0,
                sticky="w",
                padx=11,
                pady=(14 if row > 2 else 0, 5),
            )
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
                button.grid(row=row, column=0, sticky="ew", pady=2)
                self.nav_buttons[tool.key] = button
                row += 1

        parent.columnconfigure(0, minsize=204)

    def _build_tool_pages(self) -> None:
        for group in self.tool_groups:
            for tool in group.tools:
                frame = tool.factory(self.content_frame)
                frame.grid(row=0, column=0, sticky="nsew")
                self.tool_frames[tool.key] = frame

    def select_tool(self, key: str) -> None:
        tool = self.tools_by_key[key]
        frame = self.tool_frames[key]
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
        workflow_frame = self.tool_frames["workflow"]
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
        restorer_frame = self.tool_frames["french_nbsp"]
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
        width, height = calculate_initial_window_size(
            requested_width=self.root.winfo_reqwidth(),
            requested_height=self.root.winfo_reqheight(),
            screen_width=screen_width,
            screen_height=screen_height,
        )
        x = max((self.root.winfo_screenwidth() - width) // 2, 0)
        y = max((self.root.winfo_screenheight() - height) // 2, 30)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        self.root.minsize(min(width, 980), min(height, 680))
        self.root.deiconify()


def calculate_initial_window_size(
    *,
    requested_width: int,
    requested_height: int,
    screen_width: int,
    screen_height: int,
) -> tuple[int, int]:
    """Add breathing room to the requested layout while respecting the screen."""

    available_width = max(screen_width - 80, 900)
    available_height = max(screen_height - 80, 640)
    width = min(
        max(
            requested_width + WINDOW_HORIZONTAL_BREATHING_ROOM,
            DEFAULT_WINDOW_WIDTH,
        ),
        available_width,
    )
    height = min(
        max(
            requested_height + WINDOW_VERTICAL_BREATHING_ROOM,
            DEFAULT_WINDOW_HEIGHT,
        ),
        available_height,
    )
    return width, height


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
        enable_high_dpi_awareness()
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

    enable_high_dpi_awareness()
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
