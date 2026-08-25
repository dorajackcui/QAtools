"""Shared PySide6 widgets and background-task helpers for Toolshub."""

from __future__ import annotations

from collections.abc import Callable
import os
import sys
from typing import Any

from PySide6.QtCore import QObject, QPointF, QRunnable, QRectF, Qt, QThreadPool, Signal, Slot
from PySide6.QtGui import QColor, QFontDatabase, QPainter, QPainterPath, QPalette, QPen
from PySide6.QtWidgets import (
    QApplication,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProxyStyle,
    QPushButton,
    QSizePolicy,
    QStyle,
    QStyleOption,
    QVBoxLayout,
    QWidget,
)


# Warm, opaque light tokens derived from the user's Codex theme reference. The
# supplied surface, ink, and accent remain the three anchors; nearby elevations
# are deliberately subtle so structure is felt without adding visual noise.
APP_BACKGROUND = "#f9f9f7"
SIDEBAR_BACKGROUND = "#f1f1ee"
CARD_BACKGROUND = "#ffffff"
INPUT_BACKGROUND = "#f4f4f1"
BORDER_COLOR = "#deded8"
BORDER_STRONG_COLOR = "#c6c6bf"
TEXT_COLOR = "#2d2d2b"
MUTED_TEXT_COLOR = "#6f6f6a"
SUBTLE_TEXT_COLOR = "#989892"
ACCENT_COLOR = "#cc7d5e"
ACCENT_HOVER_COLOR = "#b96e51"
ACCENT_PRESSED_COLOR = "#a86149"
ACCENT_FOREGROUND_COLOR = "#ffffff"
ERROR_COLOR = "#ff5f38"


class _ToolshubStyle(QProxyStyle):
    """Draw accessible checkbox/radio indicators consistently on Win and macOS."""

    def drawPrimitive(
        self,
        element: QStyle.PrimitiveElement,
        option: QStyleOption,
        painter: QPainter,
        widget: QWidget | None = None,
    ) -> None:
        if element == QStyle.PrimitiveElement.PE_IndicatorCheckBox:
            self._draw_checkbox(option, painter)
            return
        if element == QStyle.PrimitiveElement.PE_IndicatorRadioButton:
            self._draw_radio(option, painter)
            return
        super().drawPrimitive(element, option, painter, widget)

    @staticmethod
    def _state(option: QStyleOption, flag: QStyle.StateFlag) -> bool:
        return bool(option.state & flag)

    def _draw_checkbox(self, option: QStyleOption, painter: QPainter) -> None:
        enabled = self._state(option, QStyle.StateFlag.State_Enabled)
        checked = self._state(option, QStyle.StateFlag.State_On)
        partial = self._state(option, QStyle.StateFlag.State_NoChange)
        hovered = self._state(option, QStyle.StateFlag.State_MouseOver)
        rect = QRectF(option.rect).adjusted(1.0, 1.0, -1.0, -1.0)

        if not enabled:
            border, fill, mark = "#c8c8c2", "#ecece8", "#a1a19a"
        elif checked or partial:
            border, fill, mark = ACCENT_COLOR, ACCENT_COLOR, ACCENT_FOREGROUND_COLOR
        else:
            border = ACCENT_COLOR if hovered else "#aaa9a2"
            fill, mark = INPUT_BACKGROUND, TEXT_COLOR

        painter.save()
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(QPen(QColor(border), 1.2))
        painter.setBrush(QColor(fill))
        painter.drawRoundedRect(rect, 3.0, 3.0)

        mark_pen = QPen(QColor(mark), 1.7)
        mark_pen.setCapStyle(Qt.PenCapStyle.RoundCap)
        mark_pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
        painter.setPen(mark_pen)
        if partial:
            painter.drawLine(
                QPointF(rect.left() + 3.2, rect.center().y()),
                QPointF(rect.right() - 3.2, rect.center().y()),
            )
        elif checked:
            path = QPainterPath()
            path.moveTo(rect.left() + 3.0, rect.center().y())
            path.lineTo(rect.left() + 5.8, rect.bottom() - 3.1)
            path.lineTo(rect.right() - 2.6, rect.top() + 3.0)
            painter.drawPath(path)
        painter.restore()

    def _draw_radio(self, option: QStyleOption, painter: QPainter) -> None:
        enabled = self._state(option, QStyle.StateFlag.State_Enabled)
        checked = self._state(option, QStyle.StateFlag.State_On)
        hovered = self._state(option, QStyle.StateFlag.State_MouseOver)
        rect = QRectF(option.rect).adjusted(1.0, 1.0, -1.0, -1.0)

        if not enabled:
            border, fill, dot = "#c8c8c2", "#ecece8", "#a1a19a"
        elif checked:
            border, fill, dot = ACCENT_COLOR, ACCENT_COLOR, ACCENT_FOREGROUND_COLOR
        else:
            border = ACCENT_COLOR if hovered else "#aaa9a2"
            fill, dot = INPUT_BACKGROUND, TEXT_COLOR

        painter.save()
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(QPen(QColor(border), 1.2))
        painter.setBrush(QColor(fill))
        painter.drawEllipse(rect)
        if checked:
            dot_rect = rect.adjusted(4.0, 4.0, -4.0, -4.0)
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(QColor(dot))
            painter.drawEllipse(dot_rect)
        painter.restore()


def configure_qt_application(app: QApplication) -> None:
    """Apply one opaque, cross-platform palette to avoid theme repaint gaps."""

    app.setApplicationName("QAtools")
    app.setOrganizationName("QAtools")
    app.setStyle(_ToolshubStyle("Fusion"))
    available_fonts = set(QFontDatabase.families())
    for preferred_font in ("Geist", "Inter"):
        if preferred_font in available_fonts:
            ui_font = app.font()
            ui_font.setFamily(preferred_font)
            app.setFont(ui_font)
            break
    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window, QColor(APP_BACKGROUND))
    palette.setColor(QPalette.ColorRole.WindowText, QColor(TEXT_COLOR))
    palette.setColor(QPalette.ColorRole.Base, QColor(INPUT_BACKGROUND))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor(CARD_BACKGROUND))
    palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(TEXT_COLOR))
    palette.setColor(QPalette.ColorRole.ToolTipText, QColor(APP_BACKGROUND))
    palette.setColor(QPalette.ColorRole.Text, QColor(TEXT_COLOR))
    palette.setColor(QPalette.ColorRole.Button, QColor(CARD_BACKGROUND))
    palette.setColor(QPalette.ColorRole.ButtonText, QColor(TEXT_COLOR))
    palette.setColor(QPalette.ColorRole.Highlight, QColor(ACCENT_COLOR))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor(ACCENT_FOREGROUND_COLOR))
    palette.setColor(QPalette.ColorRole.PlaceholderText, QColor(SUBTLE_TEXT_COLOR))
    palette.setColor(QPalette.ColorRole.Link, QColor(ACCENT_HOVER_COLOR))
    palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Text, QColor("#aaa9a3"))
    palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.ButtonText, QColor("#aaa9a3"))
    app.setPalette(palette)
    app.setStyleSheet(
        f"""
        QWidget {{
            color: {TEXT_COLOR};
            font-size: 13px;
        }}
        QMainWindow,
        QWidget#toolshubShell,
        QWidget#toolshubWorkspace,
        QStackedWidget#toolPageStack,
        QScrollArea,
        QScrollArea > QWidget > QWidget {{
            background: {APP_BACKGROUND};
        }}
        QScrollArea {{ border: none; }}
        QLabel {{ background: transparent; }}
        QFrame#sectionCard {{
            background: {CARD_BACKGROUND};
            border: 1px solid {BORDER_COLOR};
            border-radius: 8px;
        }}
        QFrame#pageActionBar {{
            background: {APP_BACKGROUND};
            border: none;
            border-top: 1px solid {BORDER_COLOR};
        }}
        QLabel#sectionTitle {{
            color: #464642;
            background: transparent;
            font-size: 12px;
            font-weight: 600;
        }}
        QLabel#pageTitle {{
            color: {TEXT_COLOR};
            background: transparent;
            font-size: 18px;
            font-weight: 600;
        }}
        QLabel#brandLabel {{
            color: #373734;
            background: transparent;
            font-size: 16px;
            font-weight: 650;
            padding: 0 8px 8px 8px;
        }}
        QLabel[role="navSection"] {{
            color: #92928c;
            background: transparent;
            font-size: 11px;
            font-weight: 600;
        }}
        QLineEdit, QComboBox, QSpinBox, QPlainTextEdit {{
            background: {INPUT_BACKGROUND};
            border: 1px solid #d6d6d0;
            border-radius: 6px;
            padding: 4px 7px;
            selection-background-color: {ACCENT_COLOR};
            selection-color: {ACCENT_FOREGROUND_COLOR};
        }}
        QLineEdit:hover, QComboBox:hover, QSpinBox:hover, QPlainTextEdit:hover {{
            border-color: {BORDER_STRONG_COLOR};
        }}
        QLineEdit:focus, QComboBox:focus, QSpinBox:focus, QPlainTextEdit:focus {{
            border-color: {ACCENT_COLOR};
        }}
        QLineEdit:read-only {{
            color: #767671;
            background: #f1f1ee;
        }}
        QComboBox::drop-down {{ border: none; width: 24px; }}
        QComboBox QAbstractItemView {{
            background: {INPUT_BACKGROUND};
            border: 1px solid {BORDER_COLOR};
            border-radius: 6px;
            outline: none;
            selection-background-color: #eee1dc;
            selection-color: {TEXT_COLOR};
        }}
        QPushButton {{
            color: #444440;
            background: #f1f1ee;
            border: 1px solid #d8d8d2;
            border-radius: 6px;
            padding: 5px 10px;
        }}
        QPushButton:hover {{ background: #e9e9e5; border-color: {BORDER_STRONG_COLOR}; }}
        QPushButton:focus {{ border-color: {ACCENT_COLOR}; }}
        QPushButton:pressed {{ background: #e1e1dc; }}
        QPushButton:disabled {{ color: #aaa9a3; background: #f4f4f1; border-color: #e5e5df; }}
        QPushButton[primary="true"] {{
            background: {ACCENT_COLOR};
            border-color: {ACCENT_COLOR};
            color: {ACCENT_FOREGROUND_COLOR};
            font-weight: 600;
            padding: 6px 14px;
        }}
        QPushButton[primary="true"]:hover {{
            background: {ACCENT_HOVER_COLOR};
            border-color: {ACCENT_HOVER_COLOR};
        }}
        QPushButton[primary="true"]:pressed {{
            background: {ACCENT_PRESSED_COLOR};
            border-color: {ACCENT_PRESSED_COLOR};
        }}
        QPushButton[primary="true"]:disabled {{
            color: #8a5a47;
            background: #e6b9a7;
            border-color: #e6b9a7;
        }}
        QPushButton[navItem="true"] {{
            color: #73736e;
            text-align: left;
            background: transparent;
            border: none;
            border-radius: 6px;
            padding: 7px 9px;
        }}
        QPushButton[navItem="true"]:hover {{
            color: #3e3e3a;
            background: #e9e9e5;
        }}
        QPushButton[navItem="true"]:checked {{
            color: {TEXT_COLOR};
            background: #e3e3de;
            font-weight: 600;
        }}
        QCheckBox, QRadioButton {{ spacing: 6px; background: transparent; }}
        QCheckBox::indicator, QRadioButton::indicator {{ width: 15px; height: 15px; }}
        QCheckBox:disabled, QRadioButton:disabled {{ color: #9b9b95; }}
        QFrame#segmentedControl {{
            background: {INPUT_BACKGROUND};
            border: 1px solid {BORDER_COLOR};
            border-radius: 6px;
        }}
        QPushButton[segmentedMode="true"] {{
            color: {MUTED_TEXT_COLOR};
            background: transparent;
            border: none;
            border-radius: 5px;
            padding: 5px 12px;
        }}
        QPushButton[segmentedMode="true"]:hover {{
            color: {TEXT_COLOR};
            background: #eeeeea;
        }}
        QPushButton[segmentedMode="true"]:checked {{
            color: {TEXT_COLOR};
            background: #f1ddd5;
            font-weight: 600;
        }}
        QTabWidget::pane {{ border: none; background: transparent; }}
        QTabBar {{ background: transparent; }}
        QTabBar::tab {{
            color: #797974;
            background: transparent;
            border: none;
            border-radius: 6px;
            margin: 0 2px 4px 0;
            padding: 5px 10px;
        }}
        QTabBar::tab:hover {{ color: {TEXT_COLOR}; background: #eeeeea; }}
        QTabBar::tab:selected {{ color: {TEXT_COLOR}; background: #e4e4df; font-weight: 600; }}
        QScrollBar:vertical {{ background: transparent; width: 8px; margin: 0; }}
        QScrollBar::handle:vertical {{ background: #c8c8c2; min-height: 28px; border-radius: 3px; }}
        QScrollBar::handle:vertical:hover {{ background: #aaa9a2; }}
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
        QProgressBar {{
            color: #5f5f5a;
            border: 1px solid {BORDER_COLOR};
            border-radius: 4px;
            background: {INPUT_BACKGROUND};
            text-align: center;
        }}
        QProgressBar::chunk {{ background: {ACCENT_COLOR}; }}
        QToolTip {{ color: {APP_BACKGROUND}; background: {TEXT_COLOR}; border: 1px solid {TEXT_COLOR}; }}
        """
    )


def create_qt_application(argv: list[str] | None = None) -> tuple[QApplication, bool]:
    existing = QApplication.instance()
    if isinstance(existing, QApplication):
        return existing, False
    app = QApplication(list(sys.argv if argv is None else argv))
    configure_qt_application(app)
    return app, True


def section(title: str) -> tuple[QFrame, QVBoxLayout]:
    """Build a quiet surface with an internal heading instead of a fieldset."""

    box = QFrame()
    box.setObjectName("sectionCard")
    outer = QVBoxLayout(box)
    outer.setContentsMargins(12, 10, 12, 10)
    outer.setSpacing(7)
    heading = QLabel(title)
    heading.setObjectName("sectionTitle")
    outer.addWidget(heading)
    layout = QVBoxLayout()
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(7)
    outer.addLayout(layout)
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
        layout.setSpacing(6)
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
    "ACCENT_COLOR",
    "ACCENT_FOREGROUND_COLOR",
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
