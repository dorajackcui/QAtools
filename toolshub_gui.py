#!/usr/bin/env python3
"""Unified PySide6 desktop launcher for all QAtools Excel workflows."""

from __future__ import annotations

import argparse
import ctypes
from dataclasses import dataclass
from html import escape
import os
from pathlib import Path
import sys

from PySide6.QtCore import QEventLoop, QLockFile, QStandardPaths, Qt, QTimer
from PySide6.QtGui import QCloseEvent
from PySide6.QtWidgets import (
    QApplication,
    QButtonGroup,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from tools.header_aliases import HeaderAliasStore
from tools.qt_navigation import ToolNavigationServer, send_path_request, send_tool_selection
from tools.qt_gui_common import (
    ACCENT_COLOR,
    AsyncPage,
    BORDER_COLOR,
    NavigationButton,
    SIDEBAR_BACKGROUND,
    create_qt_application,
    show_error,
    show_warning,
)
from tools.qt_pages import PAGE_FACTORIES
from tools.workflow.file_receiver import (
    FRENCH_NBSP_RESTORE_ACTION,
    QA_WORKFLOW_ACTION,
    ToolFileRequest,
    WorkflowFileReceiver,
    normalize_excel_input_file,
)


DEFAULT_WINDOW_WIDTH = 1000
DEFAULT_WINDOW_HEIGHT = 660
MINIMUM_WINDOW_WIDTH = 840
MINIMUM_WINDOW_HEIGHT = 540
SIDEBAR_WIDTH = 184
WINDOW_HORIZONTAL_BREATHING_ROOM = 16
WINDOW_VERTICAL_BREATHING_ROOM = 16
GUI_INSTANCE_LOCK_NAME = "qatools-toolshub-gui.lock"
DIRECTORY_ACTIONS = {
    "compatibility": "compatibility_dir",
    "excel_merger": "merge_dir",
    "untranslated_stats": "untranslated_dir",
}
QA_CHECK_WIDGETS = {
    "term": "term_check",
    "tag": "tag_check",
    "line-break": "line_break_check",
    "consistency": "consistency_check",
    "target-consistency": "target_consistency_check",
    "substring-consistency": "substring_consistency_check",
    "number": "number_check",
    "url": "url_check",
    "chinese": "chinese_check",
    "text": "target_text_check",
}


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
            ToolItem(key="content_sync", title="内容同步"),
        ),
    ),
    ToolGroup(
        title="文件与表格",
        tools=(
            ToolItem(key="compatibility", title="兼容性重存"),
            ToolItem(key="column_tools", title="列操作"),
            ToolItem(key="excel_batcher", title="Batch 拆分"),
            ToolItem(key="excel_merger", title="合并表格"),
            ToolItem(key="deep_replace", title="同名文件替换"),
            ToolItem(key="file_collector", title="文件提取"),
        ),
    ),
    ToolGroup(
        title="翻译辅助",
        tools=(
            ToolItem(key="french_nbsp", title="法语 NBSP 恢复"),
            ToolItem(key="untranslated_stats", title="未翻译统计"),
            ToolItem(key="xbench_report", title="Xbench QA 转换"),
        ),
    ),
)
SETTINGS_ITEM = ToolItem(key="settings", title="设置")


class _PageCache(dict[str, QWidget]):
    """Indexing creates a page once; get/items inspect only existing pages."""

    def __init__(self, create_page) -> None:
        super().__init__()
        self._create_page = create_page

    def __missing__(self, key: str) -> QWidget:
        page = self._create_page(key)
        self[key] = page
        return page


class ToolshubApp(QMainWindow):
    """One native Qt window with persistent pages in a QStackedWidget."""

    def __init__(
        self,
        *,
        show_window: bool = True,
        header_alias_store: HeaderAliasStore | None = None,
        initial_tool: str = "workflow",
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
        self.tool_frames: dict[str, QWidget] = _PageCache(self._create_page)
        self.nav_buttons: dict[str, QPushButton] = {}
        self.current_tool_key = ""
        self.current_tool_frame: QWidget | None = None
        self._receiver: WorkflowFileReceiver | None = None
        self._poll_timer: QTimer | None = None
        self._navigation: ToolNavigationServer | None = None
        self._opening_path = False
        self._build_ui()
        self._fit_window_to_screen()
        if show_window:
            self.show()
            # Paint confirmation of the click before importing workbook code.
            QApplication.processEvents(QEventLoop.ProcessEventsFlag.ExcludeUserInputEvents)
        self.select_tool(initial_tool)

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
        workspace_layout.setContentsMargins(16, 16, 16, 12)
        workspace_layout.setSpacing(12)
        self.title_label = QLabel()
        self.title_label.setObjectName("pageTitle")
        self.title_label.setTextFormat(Qt.TextFormat.RichText)
        self.title_label.setText(f'QAtools<span style="color: {ACCENT_COLOR}">.</span>')
        workspace_layout.addWidget(self.title_label)
        self.page_stack = QStackedWidget()
        self.page_stack.setObjectName("toolPageStack")
        self.loading_page = QLabel("正在启动 QAtools…")
        self.loading_page.setObjectName("startupStatus")
        self.loading_page.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.page_stack.addWidget(self.loading_page)
        workspace_layout.addWidget(self.page_stack, 1)
        shell_layout.addWidget(workspace, 1)
        self.setCentralWidget(shell)

    def _create_page(self, key: str) -> QWidget:
        factory = PAGE_FACTORIES[key]
        if key in {"workflow", "french_nbsp", "settings"}:
            page = factory(header_alias_store=self.header_alias_store)
        else:
            page = factory()
        page.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.page_stack.addWidget(page)
        if key == "settings":
            page.settings_saved.connect(self._refresh_header_detection)
        return page

    def preload_pages(self) -> None:
        """Exercise all lazy factories in packaging smoke tests."""
        for key in PAGE_FACTORIES:
            self.tool_frames[key]

    def _build_sidebar(self) -> QWidget:
        sidebar = QWidget()
        sidebar.setObjectName("toolshubSidebar")
        sidebar.setFixedWidth(SIDEBAR_WIDTH)
        sidebar.setStyleSheet(f"#toolshubSidebar {{ background: {SIDEBAR_BACKGROUND}; }}")
        layout = QVBoxLayout(sidebar)
        layout.setContentsMargins(12, 16, 10, 12)
        layout.setSpacing(2)
        brand = QLabel(f'QAtools<span style="color: {ACCENT_COLOR}">.</span>')
        brand.setAccessibleName("QAtools")
        brand.setObjectName("brandLabel")
        layout.addWidget(brand)
        self.nav_group = QButtonGroup(self)
        self.nav_group.setExclusive(True)
        navigation = QWidget()
        navigation.setObjectName("sidebarNavigation")
        navigation.setStyleSheet(f"#sidebarNavigation {{ background: {SIDEBAR_BACKGROUND}; }}")
        nav_layout = QVBoxLayout(navigation)
        nav_layout.setContentsMargins(0, 0, 0, 0)
        nav_layout.setSpacing(2)
        for group_index, group in enumerate(self.tool_groups):
            if group_index:
                nav_layout.addSpacing(10)
                divider = QFrame()
                divider.setObjectName("navSectionDivider")
                divider.setFixedHeight(1)
                divider_row = QHBoxLayout()
                divider_row.setContentsMargins(8, 0, 8, 0)
                divider_row.addWidget(divider)
                nav_layout.addLayout(divider_row)
                nav_layout.addSpacing(6)
            category = QLabel(group.title)
            category.setProperty("role", "navSection")
            category.setContentsMargins(8, 4, 8, 5)
            nav_layout.addWidget(category)
            for tool in group.tools:
                button = NavigationButton(tool.title)
                button.clicked.connect(lambda _checked=False, key=tool.key: self.select_tool(key))
                self.nav_group.addButton(button)
                self.nav_buttons[tool.key] = button
                nav_layout.addWidget(button)
        nav_layout.addStretch(1)
        self.nav_scroll = QScrollArea()
        self.nav_scroll.setWidgetResizable(True)
        self.nav_scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.nav_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.nav_scroll.setStyleSheet(f"QScrollArea {{ background: {SIDEBAR_BACKGROUND}; }}")
        self.nav_scroll.setWidget(navigation)
        layout.addWidget(self.nav_scroll, 1)
        settings_button = NavigationButton("⚙  设置")
        settings_button.setObjectName("settingsNavButton")
        settings_button.clicked.connect(
            lambda _checked=False: self.select_tool(SETTINGS_ITEM.key)
        )
        self.nav_group.addButton(settings_button)
        self.nav_buttons[SETTINGS_ITEM.key] = settings_button
        layout.addWidget(settings_button)
        return sidebar

    def _refresh_header_detection(self) -> None:
        workflow_page = self.tool_frames.get("workflow")
        if workflow_page is not None:
            workflow_page.detect_main_columns()
        french_page = self.tool_frames.get("french_nbsp")
        if french_page is not None:
            french_page.detect_columns()

    def select_tool(self, key: str) -> None:
        tool = self.tools_by_key[key]
        previous_key = self.current_tool_key
        if key not in self.tool_frames:
            self.loading_page.setText(f"正在打开{tool.title}…")
            self.page_stack.addWidget(self.loading_page)
            self.page_stack.setCurrentWidget(self.loading_page)
            if self.isVisible():
                self.repaint()
        try:
            page = self.tool_frames[key]
        except Exception:
            if previous_key:
                self.page_stack.setCurrentWidget(self.tool_frames[previous_key])
                self.nav_buttons[previous_key].setChecked(True)
            raise
        finally:
            if previous_key or key in self.tool_frames:
                self.page_stack.removeWidget(self.loading_page)
        self.current_tool_key = key
        self.current_tool_frame = page
        self.title_label.setText(f'{escape(tool.title)}<span style="color: {ACCENT_COLOR}">.</span>')
        self.title_label.setAccessibleName(tool.title)
        self.nav_buttons[key].setChecked(True)
        self.page_stack.setCurrentWidget(page)
        if key != SETTINGS_ITEM.key:
            self.nav_scroll.ensureWidgetVisible(self.nav_buttons[key])

    def open_qa_workflow_file(self, file_path: str) -> None:
        normalized = normalize_excel_input_file(file_path, action_name="QA workflow")
        page = self.tool_frames["workflow"]
        self._ensure_page_idle("workflow")
        self.select_tool("workflow")
        page.load_input_file(str(normalized))
        self._bring_window_to_front()

    def open_tool(self, key: str, checks: list[str] | None = None) -> None:
        """Select a shared Qt page, optionally presetting the QA checks."""
        if key not in self.tools_by_key or any(check not in QA_CHECK_WIDGETS for check in (checks or [])):
            return
        if checks and key != "workflow":
            return
        self.select_tool(key)
        if checks:
            page = self.tool_frames["workflow"]
            for check, widget_name in QA_CHECK_WIDGETS.items():
                getattr(page, widget_name).setChecked(check in checks)

    def _open_forwarded_tool(self, key: str, checks: list[str]) -> None:
        self.open_tool(key, checks)
        self._bring_window_to_front()

    def open_french_nbsp_restore_file(self, file_path: str, *, run_immediately: bool = True) -> None:
        normalized = normalize_excel_input_file(file_path, action_name="NBSP restore")
        page = self.tool_frames["french_nbsp"]
        self._ensure_page_idle("french_nbsp")
        self.select_tool("french_nbsp")
        self._bring_window_to_front()
        loaded = page.load_input_file(str(normalized), reset_options=True)
        if run_immediately and loaded:
            page.run_restore()

    def _ensure_page_idle(self, key: str) -> None:
        page = self.tool_frames[key]
        if isinstance(page, AsyncPage) and (page.has_running_tasks() or not page.run_button.isEnabled()):
            raise ValueError(f"{self.tools_by_key[key].title}正在处理文件，请等待完成后再从右键菜单打开。")

    def open_tool_directory(self, key: str, directory: str) -> None:
        if key not in DIRECTORY_ACTIONS:
            raise ValueError("不支持的目录操作。")
        path = Path(directory).expanduser().absolute()
        if not path.is_dir():
            raise ValueError(f"文件目录不存在：{path}")
        self._ensure_page_idle(key)
        page = self.tool_frames[key]
        picker = page.input_dir if key == "excel_merger" else page.input_picker
        picker.set_path(str(path))
        self.select_tool(key)
        self._bring_window_to_front()

    def _open_forwarded_path(self, action: str, path: str) -> None:
        try:
            self.handle_file_request(ToolFileRequest(action=action, file_path=path))
        except Exception as exc:  # noqa: BLE001
            self._bring_window_to_front()
            show_error(self, "无法打开右键操作", str(exc))

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
        if self._opening_path:
            raise ValueError("正在载入文件，请完成当前操作后重试。")
        self._opening_path = True
        try:
            if request.action == FRENCH_NBSP_RESTORE_ACTION:
                self.open_french_nbsp_restore_file(request.file_path)
            elif request.action == QA_WORKFLOW_ACTION:
                self.open_qa_workflow_file(request.file_path)
            elif request.action in DIRECTORY_ACTIONS:
                self.open_tool_directory(request.action, request.file_path)
            else:
                raise ValueError("不支持的右键操作。")
        finally:
            self._opening_path = False

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
        if self._navigation is not None:
            self._navigation.close()
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
    action_group.add_argument("--tool", choices=tuple(PAGE_FACTORIES), help="启动时打开指定的 PySide6 工具页面。")
    action_group.add_argument("--compatibility-dir", metavar="DIRECTORY", help="打开兼容性重存并填入目录。")
    action_group.add_argument("--merge-dir", metavar="DIRECTORY", help="打开合并表格并填入目录。")
    action_group.add_argument("--untranslated-dir", metavar="DIRECTORY", help="打开未翻译统计并填入目录。")
    parser.add_argument("--check", action="append", choices=tuple(QA_CHECK_WIDGETS), help="配合 --tool workflow 预选检查项，可重复；不会自动运行。")
    parser.add_argument("--smoke-test", action="store_true", help=argparse.SUPPRESS)
    return parser


def _initial_request(args: argparse.Namespace) -> ToolFileRequest | None:
    for action, argument in DIRECTORY_ACTIONS.items():
        directory = getattr(args, argument)
        if directory:
            path = Path(directory).expanduser().absolute()
            if not path.is_dir():
                raise ValueError(f"文件目录不存在：{path}")
            return ToolFileRequest(action=action, file_path=str(path))
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


def _acquire_gui_instance_lock(lock_path: str | None = None) -> QLockFile | None:
    if lock_path is None:
        temp_directory = QStandardPaths.writableLocation(
            QStandardPaths.StandardLocation.TempLocation
        )
        lock_path = os.path.join(temp_directory, GUI_INSTANCE_LOCK_NAME)
    instance_lock = QLockFile(lock_path)
    return instance_lock if instance_lock.tryLock(0) else None


def main(argv: list[str] | None = None) -> int:
    parser = build_argument_parser()
    args = parser.parse_args(argv)
    if args.check and args.tool != "workflow":
        parser.error("--check 必须配合 --tool workflow 使用。")
    if args.smoke_test:
        os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    app, owns_app = create_qt_application([sys.argv[0]])
    try:
        initial_request = _initial_request(args)
    except ValueError as exc:
        if args.smoke_test:
            print(str(exc), file=sys.stderr)
        else:
            show_error(None, "无法打开右键操作", str(exc))
        return 2
    instance_lock: QLockFile | None = None
    if not args.smoke_test:
        instance_lock = _acquire_gui_instance_lock()
        if instance_lock is None:
            if initial_request:
                if send_path_request(initial_request.action, initial_request.file_path):
                    return 0
                show_error(None, "无法打开右键操作", "工具箱未能接收路径，请等待其启动完成后重试；如果刚升级，请重启工具箱。")
                return 1
            if args.tool and not send_tool_selection(args.tool, args.check):
                print("工具箱正在启动或当前实例不支持页面切换，请稍后重试或重启工具箱。", file=sys.stderr)
                return 1
            return 0

    receiver = WorkflowFileReceiver()
    receiver_started = receiver.start()

    requested_tool = args.tool or "workflow"
    if initial_request:
        requested_tool = {
            QA_WORKFLOW_ACTION: "workflow", FRENCH_NBSP_RESTORE_ACTION: "french_nbsp",
        }.get(initial_request.action, initial_request.action)
    window = ToolshubApp(show_window=not args.smoke_test, initial_tool=requested_tool)
    if not args.smoke_test:
        window._navigation = ToolNavigationServer(window)
        window._navigation.requested.connect(window._open_forwarded_tool)
        window._navigation.path_requested.connect(window._open_forwarded_path)
        window._navigation.start()
    if args.tool:
        window.open_tool(args.tool, args.check)
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
        window.preload_pages()
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
        if instance_lock is not None:
            instance_lock.unlock()


if __name__ == "__main__":
    raise SystemExit(main())
