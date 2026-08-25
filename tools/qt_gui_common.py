"""Shared PySide6 widgets and background-task helpers for Toolshub."""

from __future__ import annotations

from collections.abc import Callable
import os
import sys
from typing import Any

from PySide6.QtCore import QObject, QRunnable, Qt, QThreadPool, Signal, Slot
from PySide6.QtGui import QColor, QPalette
from PySide6.QtWidgets import (
    QApplication,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)


APP_BACKGROUND = "#15181d"
SIDEBAR_BACKGROUND = "#101217"
CARD_BACKGROUND = "#1c2027"
INPUT_BACKGROUND = "#222730"
BORDER_COLOR = "#343b47"
TEXT_COLOR = "#f0f2f5"
MUTED_TEXT_COLOR = "#9ca5b4"
ACCENT_COLOR = "#4f8cff"
ACCENT_HOVER_COLOR = "#6a9dff"
ERROR_COLOR = "#ff7373"


def configure_qt_application(app: QApplication) -> None:
    """Apply one opaque, cross-platform palette to avoid theme repaint gaps."""

    app.setApplicationName("QAtools")
    app.setOrganizationName("QAtools")
    app.setStyle("Fusion")
    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window, QColor(APP_BACKGROUND))
    palette.setColor(QPalette.ColorRole.WindowText, QColor(TEXT_COLOR))
    palette.setColor(QPalette.ColorRole.Base, QColor(INPUT_BACKGROUND))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor(CARD_BACKGROUND))
    palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(CARD_BACKGROUND))
    palette.setColor(QPalette.ColorRole.ToolTipText, QColor(TEXT_COLOR))
    palette.setColor(QPalette.ColorRole.Text, QColor(TEXT_COLOR))
    palette.setColor(QPalette.ColorRole.Button, QColor(CARD_BACKGROUND))
    palette.setColor(QPalette.ColorRole.ButtonText, QColor(TEXT_COLOR))
    palette.setColor(QPalette.ColorRole.Highlight, QColor(ACCENT_COLOR))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor("#ffffff"))
    palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Text, QColor("#6f7784"))
    palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.ButtonText, QColor("#6f7784"))
    app.setPalette(palette)
    app.setStyleSheet(
        f"""
        QWidget {{
            color: {TEXT_COLOR};
            background: {APP_BACKGROUND};
            font-size: 13px;
        }}
        QMainWindow, QScrollArea, QScrollArea > QWidget > QWidget {{
            background: {APP_BACKGROUND};
        }}
        QGroupBox {{
            background: {CARD_BACKGROUND};
            border: 1px solid {BORDER_COLOR};
            border-radius: 8px;
            margin-top: 12px;
            padding: 14px 12px 12px 12px;
            font-weight: 600;
        }}
        QGroupBox::title {{
            subcontrol-origin: margin;
            left: 12px;
            padding: 0 5px;
            color: {TEXT_COLOR};
            background: {CARD_BACKGROUND};
        }}
        QLineEdit, QComboBox, QSpinBox, QPlainTextEdit {{
            background: {INPUT_BACKGROUND};
            border: 1px solid {BORDER_COLOR};
            border-radius: 5px;
            padding: 6px 8px;
            selection-background-color: {ACCENT_COLOR};
        }}
        QLineEdit:focus, QComboBox:focus, QSpinBox:focus, QPlainTextEdit:focus {{
            border-color: {ACCENT_COLOR};
        }}
        QComboBox::drop-down {{ border: none; width: 24px; }}
        QComboBox QAbstractItemView {{
            background: {INPUT_BACKGROUND};
            border: 1px solid {BORDER_COLOR};
            selection-background-color: {ACCENT_COLOR};
        }}
        QPushButton {{
            background: #2a303a;
            border: 1px solid #414958;
            border-radius: 5px;
            padding: 7px 13px;
        }}
        QPushButton:hover {{ background: #343c48; }}
        QPushButton:pressed {{ background: #222832; }}
        QPushButton:disabled {{ color: #6f7784; background: #20242b; border-color: #2a3038; }}
        QPushButton[primary="true"] {{
            background: {ACCENT_COLOR};
            border-color: {ACCENT_COLOR};
            color: white;
            font-weight: 600;
            padding: 9px 16px;
        }}
        QPushButton[primary="true"]:hover {{ background: {ACCENT_HOVER_COLOR}; }}
        QCheckBox, QRadioButton {{ spacing: 7px; background: transparent; }}
        QCheckBox::indicator, QRadioButton::indicator {{ width: 16px; height: 16px; }}
        QTabWidget::pane {{ border: 1px solid {BORDER_COLOR}; border-radius: 6px; }}
        QTabBar::tab {{
            background: {CARD_BACKGROUND};
            border: 1px solid {BORDER_COLOR};
            padding: 8px 18px;
        }}
        QTabBar::tab:selected {{ background: {ACCENT_COLOR}; color: white; }}
        QScrollBar:vertical {{ background: {APP_BACKGROUND}; width: 12px; margin: 0; }}
        QScrollBar::handle:vertical {{ background: #48505d; min-height: 30px; border-radius: 5px; }}
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
        QProgressBar {{
            border: 1px solid {BORDER_COLOR}; border-radius: 4px;
            background: {INPUT_BACKGROUND}; text-align: center;
        }}
        QProgressBar::chunk {{ background: {ACCENT_COLOR}; }}
        QToolTip {{ color: {TEXT_COLOR}; background: {CARD_BACKGROUND}; border: 1px solid {BORDER_COLOR}; }}
        """
    )


def create_qt_application(argv: list[str] | None = None) -> tuple[QApplication, bool]:
    existing = QApplication.instance()
    if isinstance(existing, QApplication):
        return existing, False
    app = QApplication(list(sys.argv if argv is None else argv))
    configure_qt_application(app)
    return app, True


def section(title: str) -> tuple[QGroupBox, QVBoxLayout]:
    box = QGroupBox(title)
    layout = QVBoxLayout(box)
    layout.setContentsMargins(12, 15, 12, 12)
    layout.setSpacing(10)
    return box, layout


def muted_label(text: str = "", *, word_wrap: bool = False) -> QLabel:
    label = QLabel(text)
    label.setStyleSheet(f"color: {MUTED_TEXT_COLOR}; background: transparent;")
    label.setWordWrap(word_wrap)
    return label


def primary_button(text: str) -> QPushButton:
    button = QPushButton(text)
    button.setProperty("primary", True)
    button.setCursor(Qt.CursorShape.PointingHandCursor)
    return button


def horizontal_rule() -> QFrame:
    rule = QFrame()
    rule.setFrameShape(QFrame.Shape.HLine)
    rule.setStyleSheet(f"color: {BORDER_COLOR};")
    return rule


class PathPicker(QWidget):
    path_changed = Signal(str)

    def __init__(
        self,
        label: str,
        *,
        choose_text: str = "选择文件",
        allow_clear: bool = False,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        title = QLabel(label)
        title.setMinimumWidth(112)
        self.line_edit = QLineEdit()
        self.line_edit.setReadOnly(True)
        self.line_edit.setPlaceholderText("尚未选择")
        self.line_edit.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.choose_button = QPushButton(choose_text)
        self.clear_button: QPushButton | None = None
        layout.addWidget(title)
        layout.addWidget(self.line_edit, 1)
        layout.addWidget(self.choose_button)
        if allow_clear:
            self.clear_button = QPushButton("清空")
            self.clear_button.clicked.connect(self.clear)
            layout.addWidget(self.clear_button)

    def path(self) -> str:
        return self.line_edit.text().strip()

    def set_path(self, path: str) -> None:
        normalized = str(path or "")
        self.line_edit.setText(normalized)
        self.line_edit.setToolTip(normalized)
        self.path_changed.emit(normalized)

    @Slot()
    def clear(self) -> None:
        self.set_path("")


class WorkerSignals(QObject):
    succeeded = Signal(object)
    failed = Signal(str)
    progress = Signal(int, int)


class BackgroundWorker(QRunnable):
    def __init__(self, function: Callable[..., Any], *args: Any, **kwargs: Any) -> None:
        super().__init__()
        self.function = function
        self.args = args
        self.kwargs = kwargs
        self.signals = WorkerSignals()

    @Slot()
    def run(self) -> None:
        try:
            result = self.function(*self.args, **self.kwargs)
        except Exception as exc:  # noqa: BLE001 - errors are shown in the GUI
            self.signals.failed.emit(str(exc))
            return
        self.signals.succeeded.emit(result)


class AsyncPage(QWidget):
    """Base page that keeps Excel work away from Qt's GUI thread."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._thread_pool = QThreadPool.globalInstance()
        self._workers: set[BackgroundWorker] = set()

    def has_running_tasks(self) -> bool:
        """Return whether this page still owns an unfinished worker."""

        return bool(self._workers)

    def run_in_background(
        self,
        function: Callable[..., Any],
        *,
        on_success: Callable[[Any], None],
        on_error: Callable[[str], None],
        on_progress: Callable[[int, int], None] | None = None,
        args: tuple[Any, ...] = (),
        kwargs: dict[str, Any] | None = None,
    ) -> None:
        worker = BackgroundWorker(function, *args, **(kwargs or {}))
        self._workers.add(worker)

        def cleanup() -> None:
            self._workers.discard(worker)

        worker.signals.succeeded.connect(on_success)
        worker.signals.succeeded.connect(lambda _result: cleanup())
        worker.signals.failed.connect(on_error)
        worker.signals.failed.connect(lambda _message: cleanup())
        if on_progress is not None:
            worker.signals.progress.connect(on_progress)
        self._thread_pool.start(worker)


def show_error(parent: QWidget, title: str, message: str) -> None:
    QMessageBox.critical(parent, title, message)


def show_info(parent: QWidget, title: str, message: str) -> None:
    QMessageBox.information(parent, title, message)


def show_warning(parent: QWidget, title: str, message: str) -> None:
    QMessageBox.warning(parent, title, message)


def parse_positive_int(raw_value: str, *, default: int, field_name: str) -> int:
    value = raw_value.strip() or str(default)
    try:
        parsed = int(value)
    except ValueError as exc:
        raise ValueError(f"{field_name}必须是整数。") from exc
    if parsed < 1:
        raise ValueError(f"{field_name}必须大于等于 1。")
    return parsed


def reveal_in_file_manager(path: str) -> None:
    """Open the containing directory without depending on a GUI toolkit."""

    target = os.path.abspath(path)
    if sys.platform == "darwin":
        os.spawnlp(os.P_NOWAIT, "open", "open", "-R", target)
    elif os.name == "nt":
        os.startfile(os.path.dirname(target))  # type: ignore[attr-defined]


__all__ = [
    "APP_BACKGROUND",
    "AsyncPage",
    "BackgroundWorker",
    "BORDER_COLOR",
    "CARD_BACKGROUND",
    "MUTED_TEXT_COLOR",
    "PathPicker",
    "SIDEBAR_BACKGROUND",
    "TEXT_COLOR",
    "configure_qt_application",
    "create_qt_application",
    "horizontal_rule",
    "muted_label",
    "parse_positive_int",
    "primary_button",
    "section",
    "show_error",
    "show_info",
    "show_warning",
]
