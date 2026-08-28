#!/usr/bin/env python3
"""Unified PySide6 desktop launcher for all QAtools Excel workflows."""

from __future__ import annotations

import argparse
import ctypes
from dataclasses import dataclass
import os
import sys

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QCloseEvent
from PySide6.QtWidgets import (
    QApplication,
    QButtonGroup,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QPushButton,
    QSizePolicy,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from tools.qt_gui_common import (
    AsyncPage,
    BORDER_COLOR,
    SIDEBAR_BACKGROUND,
    create_qt_application,
    show_error,
    show_warning,
)
from tools.header_aliases import HeaderAliasStore
from tools.qt_pages import FrenchNbspPage, PAGE_FACTORIES, SettingsPage, WorkflowPage
from tools.workflow.file_receiver import (
    FRENCH_NBSP_RESTORE_ACTION,
    QA_WORKFLOW_ACTION,
    ToolFileRequest,
    WorkflowFileReceiver,
    normalize_excel_input_file,
    send_tool_input_file,
)


DEFAULT_WINDOW_WIDTH = 1000
DEFAULT_WINDOW_HEIGHT = 660
MINIMUM_WINDOW_WIDTH = 840
MINIMUM_WINDOW_HEIGHT = 540
SIDEBAR_WIDTH = 184
WINDOW_HORIZONTAL_BREATHING_ROOM = 16
WINDOW_VERTICAL_BREATHING_ROOM = 16


@dataclass(frozen=True)
class ToolItem:
    key: str
    title: str


@dataclass(frozen=True)
class ToolGroup:
    title: str
    tools: tuple[ToolItem, ...]


TOOL_GROUPS = (
    ToolGroup(
        title="常用流程",
        tools=(
            ToolItem(key="workflow", title="一键质量检查"),
            ToolItem(key="phraseloom", title="PhraseLoom"),
        ),
    ),
    ToolGroup(
        title="文本修复",
        tools=(ToolItem(key="french_nbsp", title="法语 NBSP 恢复"),),
    ),
    ToolGroup(
        title="其他",
        tools=(
            ToolItem(key="excel_batcher", title="Batch 拆分"),
            ToolItem(key="excel_merger", title="合并表格"),
            ToolItem(key="xbench_report", title="Xbench QA 转换"),
        ),
    ),
)
SETTINGS_ITEM = ToolItem(key="settings", title="设置")


class ToolshubApp(QMainWindow):
    """One native Qt window with persistent pages in a QStackedWidget."""

    def __init__(
        self,
        *,
        show_window: bool = True,
        header_alias_store: HeaderAliasStore | None = None,
    ) -> None:
        super().__init__()
        self.setWindowTitle("Toolshub")
        self.setMinimumSize(MINIMUM_WINDOW_WIDTH, MINIMUM_WINDOW_HEIGHT)
        self.setAutoFillBackground(True)
        self.header_alias_store = header_alias_store or HeaderAliasStore()
        self.tool_groups = TOOL_GROUPS
        self.tools_by_key = {
            tool.key: tool
            for group in self.tool_groups
            for tool in group.tools
        }
        self.tools_by_key[SETTINGS_ITEM.key] = SETTINGS_ITEM
        self.tool_frames: dict[str, QWidget] = {}
        self.nav_buttons: dict[str, QPushButton] = {}
        self.current_tool_key = ""
        self.current_tool_frame: QWidget | None = None
        self._receiver: WorkflowFileReceiver | None = None
        self._poll_timer: QTimer | None = None
        self._build_ui()
        self._fit_window_to_screen()
        self.select_tool(self.tool_groups[0].tools[0].key)
        if show_window:
            self.show()

    def _build_ui(self) -> None:
        shell = QWidget()
        shell.setObjectName("toolshubShell")
        shell_layout = QHBoxLayout(shell)
        shell_layout.setContentsMargins(0, 0, 0, 0)
        shell_layout.setSpacing(0)
        self.sidebar = self._build_sidebar()
        shell_layout.addWidget(self.sidebar)

        separator = QFrame()
        separator.setObjectName("sidebarSeparator")
        separator.setFrameShape(QFrame.Shape.VLine)
        separator.setStyleSheet(f"color: {BORDER_COLOR}; background: {BORDER_COLOR}; max-width: 1px;")
        shell_layout.addWidget(separator)

        workspace = QWidget()
        workspace.setObjectName("toolshubWorkspace")
        workspace_layout = QVBoxLayout(workspace)
        workspace_layout.setContentsMargins(16, 14, 16, 12)
        workspace_layout.setSpacing(7)
        self.title_label = QLabel()
        self.title_label.setObjectName("pageTitle")
        workspace_layout.addWidget(self.title_label)
        self.page_stack = QStackedWidget()
        self.page_stack.setObjectName("toolPageStack")
        workspace_layout.addWidget(self.page_stack, 1)
        shell_layout.addWidget(workspace, 1)
        self.setCentralWidget(shell)

        for group in self.tool_groups:
            for tool in group.tools:
                page_factory = PAGE_FACTORIES[tool.key]
                if tool.key in {"workflow", "french_nbsp"}:
                    page = page_factory(header_alias_store=self.header_alias_store)
                else:
                    page = page_factory()
                page.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
                self.page_stack.addWidget(page)
                self.tool_frames[tool.key] = page
        settings_page = PAGE_FACTORIES[SETTINGS_ITEM.key](
            header_alias_store=self.header_alias_store
        )
        settings_page.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.page_stack.addWidget(settings_page)
        self.tool_frames[SETTINGS_ITEM.key] = settings_page
        if isinstance(settings_page, SettingsPage):
            settings_page.settings_saved.connect(self._refresh_header_detection)

    def _build_sidebar(self) -> QWidget:
        sidebar = QWidget()
        sidebar.setObjectName("toolshubSidebar")
        sidebar.setFixedWidth(SIDEBAR_WIDTH)
        sidebar.setStyleSheet(f"#toolshubSidebar {{ background: {SIDEBAR_BACKGROUND}; }}")
        layout = QVBoxLayout(sidebar)
        layout.setContentsMargins(12, 16, 10, 12)
        layout.setSpacing(2)
        brand = QLabel("QAtools")
        brand.setObjectName("brandLabel")
        layout.addWidget(brand)
        self.nav_group = QButtonGroup(self)
        self.nav_group.setExclusive(True)
        for group_index, group in enumerate(self.tool_groups):
            category = QLabel(group.title)
            category.setProperty("role", "navSection")
            category.setContentsMargins(8, 8 if group_index else 2, 8, 3)
            layout.addWidget(category)
            for tool in group.tools:
                button = QPushButton(tool.title)
                button.setProperty("navItem", True)
                button.setCheckable(True)
                button.setCursor(Qt.CursorShape.PointingHandCursor)
                button.clicked.connect(lambda _checked=False, key=tool.key: self.select_tool(key))
                self.nav_group.addButton(button)
                self.nav_buttons[tool.key] = button
                layout.addWidget(button)
        layout.addStretch(1)
        settings_button = QPushButton("⚙  设置")
        settings_button.setObjectName("settingsNavButton")
        settings_button.setProperty("navItem", True)
        settings_button.setCheckable(True)
        settings_button.setCursor(Qt.CursorShape.PointingHandCursor)
        settings_button.clicked.connect(
            lambda _checked=False: self.select_tool(SETTINGS_ITEM.key)
        )
        self.nav_group.addButton(settings_button)
        self.nav_buttons[SETTINGS_ITEM.key] = settings_button
        layout.addWidget(settings_button)
        return sidebar

    def _refresh_header_detection(self) -> None:
        workflow_page = self.tool_frames.get("workflow")
        if isinstance(workflow_page, WorkflowPage):
            workflow_page.detect_main_columns()
        french_page = self.tool_frames.get("french_nbsp")
        if isinstance(french_page, FrenchNbspPage):
            french_page.detect_columns()

    def select_tool(self, key: str) -> None:
        tool = self.tools_by_key[key]
        page = self.tool_frames[key]
        self.current_tool_key = key
        self.current_tool_frame = page
        self.title_label.setText(tool.title)
        self.nav_buttons[key].setChecked(True)
        self.page_stack.setCurrentWidget(page)

    def open_qa_workflow_file(self, file_path: str) -> None:
        normalized = normalize_excel_input_file(file_path, action_name="QA workflow")
        page = self.tool_frames["workflow"]
        if not isinstance(page, WorkflowPage):
            raise RuntimeError("一键质量检查页面未正确加载。")
        self.select_tool("workflow")
        page.load_input_file(str(normalized))
        self._bring_window_to_front()

    def open_french_nbsp_restore_file(self, file_path: str, *, run_immediately: bool = True) -> None:
        normalized = normalize_excel_input_file(file_path, action_name="NBSP restore")
        page = self.tool_frames["french_nbsp"]
        if not isinstance(page, FrenchNbspPage):
            raise RuntimeError("法语 NBSP 恢复页面未正确加载。")
        self.select_tool("french_nbsp")
        page.load_input_file(str(normalized), reset_options=True)
        self._bring_window_to_front()
        if run_immediately:
            QTimer.singleShot(0, page.run_restore)

    def _bring_window_to_front(self) -> None:
        if self.isMinimized():
            self.showNormal()
        else:
            self.show()
        self.raise_()
        self.activateWindow()
        if os.name == "nt":
            try:
                window_handle = int(self.winId())
                user32 = ctypes.windll.user32
                user32.ShowWindow(window_handle, 9)  # SW_RESTORE
                user32.SetForegroundWindow(window_handle)
            except (AttributeError, OSError, TypeError, ValueError):
                pass

    def _fit_window_to_screen(self) -> None:
        screen = QApplication.primaryScreen()
        if screen is None:
            self.resize(DEFAULT_WINDOW_WIDTH, DEFAULT_WINDOW_HEIGHT)
            return
        available = screen.availableGeometry()
        requested = self.sizeHint()
        width, height = calculate_initial_window_size(
            requested_width=requested.width(),
            requested_height=requested.height(),
            screen_width=available.width(),
            screen_height=available.height(),
        )
        self.resize(width, height)
        self.move(
            available.x() + max((available.width() - width) // 2, 0),
            available.y() + max((available.height() - height) // 2, 0),
        )

    def attach_receiver(self, receiver: WorkflowFileReceiver) -> None:
        self._receiver = receiver
        self._poll_timer = QTimer(self)
        self._poll_timer.setInterval(150)
        self._poll_timer.timeout.connect(self._poll_forwarded_files)
        self._poll_timer.start()

    def _poll_forwarded_files(self) -> None:
        if self._receiver is None:
            return
        for request in self._receiver.pop_pending_requests():
            try:
                self.handle_file_request(request)
            except Exception as exc:  # noqa: BLE001
                show_error(self, "无法载入 Excel", str(exc))

    def handle_file_request(self, request: ToolFileRequest) -> None:
        if request.action == FRENCH_NBSP_RESTORE_ACTION:
            self.open_french_nbsp_restore_file(request.file_path)
        else:
            self.open_qa_workflow_file(request.file_path)

    def closeEvent(self, event: QCloseEvent) -> None:  # noqa: N802 - Qt API
        running_tools = [
            self.tools_by_key[key].title
            for key, page in self.tool_frames.items()
            if isinstance(page, AsyncPage) and page.has_running_tasks()
        ]
        if running_tools:
            event.ignore()
            show_warning(
                self,
                "任务仍在执行",
                (
                    "以下任务仍在处理 Excel，完成前不能关闭程序：\n"
                    + "、".join(running_tools)
                    + "\n\n请等待任务完成后再关闭，以免输出文件不完整。"
                ),
            )
            return
        if self._poll_timer is not None:
            self._poll_timer.stop()
        if self._receiver is not None:
            self._receiver.close()
        super().closeEvent(event)


def calculate_initial_window_size(
    *,
    requested_width: int,
    requested_height: int,
    screen_width: int,
    screen_height: int,
) -> tuple[int, int]:
    available_width = max(screen_width - 80, 900)
    available_height = max(screen_height - 80, 640)
    width = min(
        max(requested_width + WINDOW_HORIZONTAL_BREATHING_ROOM, DEFAULT_WINDOW_WIDTH),
        available_width,
    )
    height = min(
        max(requested_height + WINDOW_VERTICAL_BREATHING_ROOM, DEFAULT_WINDOW_HEIGHT),
        available_height,
    )
    return width, height


def build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="打开 Toolshub Excel 工具箱。")
    action_group = parser.add_mutually_exclusive_group()
    action_group.add_argument("--qa-workflow", metavar="EXCEL_FILE", help="把 Excel 文件载入一键质量检查页面。")
    action_group.add_argument("--nbsp-restore", metavar="EXCEL_FILE", help="对 Excel 自动执行法语 NBSP 恢复。")
    parser.add_argument("--smoke-test", action="store_true", help=argparse.SUPPRESS)
    return parser


def _initial_request(args: argparse.Namespace) -> ToolFileRequest | None:
    if args.qa_workflow:
        return ToolFileRequest(
            action=QA_WORKFLOW_ACTION,
            file_path=str(normalize_excel_input_file(args.qa_workflow, action_name="QA workflow")),
        )
    if args.nbsp_restore:
        return ToolFileRequest(
            action=FRENCH_NBSP_RESTORE_ACTION,
            file_path=str(normalize_excel_input_file(args.nbsp_restore, action_name="NBSP restore")),
        )
    return None


def main(argv: list[str] | None = None) -> int:
    args = build_argument_parser().parse_args(argv)
    if args.smoke_test:
        os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    try:
        initial_request = _initial_request(args)
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 2

    if initial_request and send_tool_input_file(initial_request.action, initial_request.file_path):
        return 0

    receiver = WorkflowFileReceiver()
    receiver_started = receiver.start()
    if initial_request and not receiver_started and send_tool_input_file(initial_request.action, initial_request.file_path):
        return 0

    app, owns_app = create_qt_application([sys.argv[0]])
    window = ToolshubApp(show_window=not args.smoke_test)
    if receiver_started:
        window.attach_receiver(receiver)
    if initial_request:
        try:
            window.handle_file_request(initial_request)
        except Exception as exc:  # noqa: BLE001
            if args.smoke_test:
                print(str(exc), file=sys.stderr)
                receiver.close()
                window.close()
                return 2
            show_error(window, "无法载入 Excel", str(exc))

    if args.smoke_test:
        app.processEvents()
        receiver.close()
        window.close()
        return 0
    # pythonw can create the native window before Qt starts dispatching events,
    # which lets Windows leave it behind other applications. Activate once as
    # the event loop starts and once more after the native window settles.
    QTimer.singleShot(0, window._bring_window_to_front)
    QTimer.singleShot(180, window._bring_window_to_front)
    if not owns_app:
        return 0
    try:
        return app.exec()
    finally:
        receiver.close()


if __name__ == "__main__":
    raise SystemExit(main())
