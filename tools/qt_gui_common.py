"""Shared PySide6 widgets and background-task helpers for Toolshub."""

from __future__ import annotations

from collections.abc import Callable
import os
import sys
from typing import Any

from PySide6.QtCore import QEvent, QObject, QPointF, QRunnable, QRectF, Qt, QThreadPool, Signal, Slot
from PySide6.QtGui import QColor, QFontDatabase, QPainter, QPainterPath, QPalette, QPen
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProxyStyle,
    QPushButton,
    QRadioButton,
    QSizePolicy,
    QStyle,
    QStyleOption,
    QTabBar,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

from tools import qt_control_resources  # noqa: F401 - registers embedded control icons

# YIZHI visual reference 1.1 (2026-09-15), adapted to compact desktop forms.
# Brand, actions, selection and keyboard focus have separate semantic roles.
# See docs/gui-conventions.md for the desktop adaptation.
APP_BACKGROUND = "#ffffff"
SIDEBAR_BACKGROUND = "#f7f7f5"
CARD_BACKGROUND = "#ffffff"
INPUT_BACKGROUND = "#f7f7f5"
BUTTON_BACKGROUND = "#f0f0ed"
BUTTON_BORDER_COLOR = "#e3e3dd"
BORDER_COLOR = "#e9e9e5"
CONTROL_BORDER_COLOR = "#deded8"
BORDER_STRONG_COLOR = "#b6b6ac"
INDICATOR_BORDER_COLOR = "#85857c"
CONTROL_FOCUS_COLOR = "#96968c"
TEXT_COLOR = "#2f2f2c"
MUTED_TEXT_COLOR = "#707068"
SUBTLE_TEXT_COLOR = MUTED_TEXT_COLOR
ACCENT_COLOR = "#d97757"
# A deeper brand orange keeps white action labels readable (4.66:1).
PRIMARY_COLOR = "#b45b3c"
PRIMARY_FOREGROUND_COLOR = "#ffffff"
ACCENT_FOREGROUND_COLOR = PRIMARY_FOREGROUND_COLOR  # Compatibility import.
PRIMARY_HOVER_COLOR = "#a65034"
PRIMARY_PRESSED_COLOR = "#91442c"
FOCUS_COLOR = "#2383e2"
HOVER_BACKGROUND = "#f0f0ed"
SELECTION_BACKGROUND = "#e9e9e5"
PROGRESS_COLOR = "#b6b6ac"
DISABLED_BACKGROUND = "#f7f7f5"
DISABLED_TEXT_COLOR = "#85857c"
DISABLED_BORDER_COLOR = "#d5d5ce"
ERROR_COLOR = "#ff5f38"


class NavigationButton(QPushButton):
    """Keep mouse selection flat and show a focus border for keyboard navigation."""

    def __init__(self, text: str, parent: QWidget | None = None) -> None:
        super().__init__(text, parent)
        self._keyboard_focus = False
        self.setProperty("navItem", True)
        self.setCheckable(True)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

    def event(self, event: QEvent) -> bool:
        if event.type() == QEvent.Type.FocusIn:
            self._keyboard_focus = event.reason() in (
                Qt.FocusReason.TabFocusReason,
                Qt.FocusReason.BacktabFocusReason,
                Qt.FocusReason.ShortcutFocusReason,
            )
        elif event.type() == QEvent.Type.KeyPress:
            self._keyboard_focus = True
        elif event.type() == QEvent.Type.MouseButtonPress:
            # A click may keep focus on the same button, with no FocusIn event.
            self._keyboard_focus = False
        else:
            return super().event(event)
        self.update()
        return super().event(event)

    def paintEvent(self, event) -> None:
        super().paintEvent(event)
        if self.hasFocus() and self._keyboard_focus:
            painter = QPainter(self)
            painter.setRenderHint(QPainter.RenderHint.Antialiasing)
            painter.setPen(QPen(QColor(CONTROL_FOCUS_COLOR), 1.0))
            painter.setBrush(Qt.BrushStyle.NoBrush)
            painter.drawRoundedRect(QRectF(self.rect()).adjusted(1, 1, -1, -1), 3, 3)
            painter.end()


class _ToolshubStyle(QProxyStyle):
    """Draw flat, accessible control indicators consistently across platforms."""

    def drawPrimitive(
        self,
        element: QStyle.PrimitiveElement,
        option: QStyleOption,
        painter: QPainter,
        widget: QWidget | None = None,
    ) -> None:
        if element == QStyle.PrimitiveElement.PE_FrameFocusRect:
            if isinstance(widget, QPushButton) or (
                isinstance(widget, QToolButton) and widget.property("settingsButton")
            ):
                return  # The stylesheet already outlines the entire button.
            focus_rect = option.rect
            if isinstance(widget, (QCheckBox, QRadioButton)):
                focus_rect = widget.rect()
            elif isinstance(widget, QTabBar):
                focus_rect = widget.tabRect(widget.currentIndex())
            painter.save()
            painter.setRenderHint(QPainter.RenderHint.Antialiasing)
            pen = QPen(QColor(CONTROL_FOCUS_COLOR), 1.0, Qt.PenStyle.DashLine)
            painter.setPen(pen)
            painter.setBrush(Qt.BrushStyle.NoBrush)
            painter.drawRoundedRect(QRectF(focus_rect).adjusted(1, 1, -1, -1), 3, 3)
            painter.restore()
            return
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
            border, fill, mark = DISABLED_BORDER_COLOR, DISABLED_BACKGROUND, DISABLED_TEXT_COLOR
        elif checked or partial:
            border, fill, mark = PRIMARY_COLOR, PRIMARY_COLOR, PRIMARY_FOREGROUND_COLOR
        else:
            border = MUTED_TEXT_COLOR if hovered else INDICATOR_BORDER_COLOR
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
            border, fill, dot = DISABLED_BORDER_COLOR, DISABLED_BACKGROUND, DISABLED_TEXT_COLOR
        elif checked:
            border, fill, dot = PRIMARY_COLOR, PRIMARY_COLOR, PRIMARY_FOREGROUND_COLOR
        else:
            border = MUTED_TEXT_COLOR if hovered else INDICATOR_BORDER_COLOR
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
    for preferred_font in ("Inter",):
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
    palette.setColor(QPalette.ColorRole.Button, QColor(BUTTON_BACKGROUND))
    palette.setColor(QPalette.ColorRole.ButtonText, QColor(TEXT_COLOR))
    palette.setColor(QPalette.ColorRole.Highlight, QColor(PRIMARY_COLOR))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor(PRIMARY_FOREGROUND_COLOR))
    palette.setColor(QPalette.ColorRole.PlaceholderText, QColor(SUBTLE_TEXT_COLOR))
    palette.setColor(QPalette.ColorRole.Link, QColor(TEXT_COLOR))
    for role in (QPalette.ColorRole.Text, QPalette.ColorRole.WindowText, QPalette.ColorRole.ButtonText):
        palette.setColor(QPalette.ColorGroup.Disabled, role, QColor(DISABLED_TEXT_COLOR))
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
            border-radius: 6px;
        }}
        QDialog#settingsDialog {{
            background: {CARD_BACKGROUND};
        }}
        QFrame#pageActionBar {{
            background: {APP_BACKGROUND};
            border: none;
            border-top: 1px solid {BORDER_COLOR};
        }}
        QLabel#sectionTitle {{
            color: {TEXT_COLOR};
            background: transparent;
            font-size: 13px;
            font-weight: 600;
        }}
        QLabel#pageTitle {{
            color: {TEXT_COLOR};
            background: transparent;
            font-size: 24px;
            font-weight: 600;
        }}
        QLabel#brandLabel {{
            color: {TEXT_COLOR};
            background: transparent;
            font-size: 20px;
            font-weight: 600;
            padding: 0 8px 8px 8px;
        }}
        QLabel[role="navSection"] {{
            color: {MUTED_TEXT_COLOR};
            background: transparent;
            font-size: 12px;
            font-weight: 600;
        }}
        QFrame#navSectionDivider {{
            background: {BORDER_COLOR};
            border: none;
        }}
        QLineEdit, QComboBox, QSpinBox, QPlainTextEdit {{
            background: {INPUT_BACKGROUND};
            border: 1px solid {CONTROL_BORDER_COLOR};
            border-radius: 4px;
            padding: 4px 7px;
            placeholder-text-color: {MUTED_TEXT_COLOR};
            selection-background-color: {PRIMARY_COLOR};
            selection-color: {PRIMARY_FOREGROUND_COLOR};
        }}
        QLineEdit:hover, QComboBox:hover, QSpinBox:hover, QPlainTextEdit:hover {{
            border-color: {BORDER_STRONG_COLOR};
        }}
        QLineEdit:focus, QComboBox:focus, QSpinBox:focus, QPlainTextEdit:focus {{
            border: 2px solid {FOCUS_COLOR};
            padding: 3px 6px;
            background: {CARD_BACKGROUND};
        }}
        QLineEdit:read-only {{
            color: {MUTED_TEXT_COLOR};
            background: {SIDEBAR_BACKGROUND};
        }}
        QLineEdit:read-only:focus, QPlainTextEdit:read-only:focus,
        QSpinBox:read-only:focus, QComboBox:!editable:focus {{
            border: 1px solid {CONTROL_FOCUS_COLOR};
            padding: 4px 7px;
            background: {INPUT_BACKGROUND};
        }}
        QLineEdit:disabled, QComboBox:disabled, QSpinBox:disabled, QPlainTextEdit:disabled {{
            color: {DISABLED_TEXT_COLOR};
            background: {DISABLED_BACKGROUND};
            border-color: {DISABLED_BORDER_COLOR};
        }}
        QSpinBox {{ padding-right: 24px; }}
        QSpinBox:focus {{ padding-right: 23px; }}
        QSpinBox:read-only:focus {{ padding-right: 24px; }}
        QSpinBox::up-button, QSpinBox::down-button {{
            subcontrol-origin: padding;
            width: 20px;
            border: none;
            background: transparent;
            border-radius: 3px;
        }}
        QSpinBox::up-button {{ subcontrol-position: top right; }}
        QSpinBox::down-button {{ subcontrol-position: bottom right; }}
        QSpinBox::up-button:hover, QSpinBox::down-button:hover {{ background: {HOVER_BACKGROUND}; }}
        QSpinBox::up-button:pressed, QSpinBox::down-button:pressed {{ background: {SELECTION_BACKGROUND}; }}
        QSpinBox::up-button:off, QSpinBox::down-button:off,
        QSpinBox::up-button:disabled, QSpinBox::down-button:disabled {{ background: transparent; }}
        QSpinBox::up-arrow, QSpinBox::down-arrow {{ width: 10px; height: 6px; }}
        QSpinBox::up-arrow {{ image: url(:/qatools/controls/spin-up.svg); }}
        QSpinBox::down-arrow {{ image: url(:/qatools/controls/spin-down.svg); }}
        QSpinBox::up-arrow:disabled, QSpinBox::up-arrow:off {{ image: url(:/qatools/controls/spin-up-disabled.svg); }}
        QSpinBox::down-arrow:disabled, QSpinBox::down-arrow:off {{ image: url(:/qatools/controls/spin-down-disabled.svg); }}
        QComboBox::drop-down {{ border: none; width: 24px; }}
        QComboBox QAbstractItemView {{
            background: {INPUT_BACKGROUND};
            border: 1px solid {CONTROL_BORDER_COLOR};
            border-radius: 4px;
            outline: none;
            selection-background-color: {SELECTION_BACKGROUND};
            selection-color: {TEXT_COLOR};
        }}
        QPushButton {{
            color: {TEXT_COLOR};
            background: {BUTTON_BACKGROUND};
            border: 1px solid {BUTTON_BORDER_COLOR};
            border-radius: 4px;
            padding: 5px 10px;
        }}
        QPushButton:hover {{ background: {SELECTION_BACKGROUND}; border-color: {CONTROL_BORDER_COLOR}; }}
        QPushButton:focus {{ border-color: {CONTROL_FOCUS_COLOR}; }}
        QPushButton:pressed {{ background: {SELECTION_BACKGROUND}; }}
        QPushButton:disabled {{ color: {DISABLED_TEXT_COLOR}; background: {DISABLED_BACKGROUND}; border-color: {DISABLED_BORDER_COLOR}; }}
        QPushButton[primary="true"], QDialogButtonBox QPushButton:default,
        QMessageBox QPushButton:default, QInputDialog QPushButton:default {{
            background: {PRIMARY_COLOR};
            border-color: {PRIMARY_COLOR};
            color: {PRIMARY_FOREGROUND_COLOR};
            font-weight: 600;
            padding: 6px 14px;
        }}
        QPushButton[primary="true"]:hover, QDialogButtonBox QPushButton:default:hover,
        QMessageBox QPushButton:default:hover, QInputDialog QPushButton:default:hover {{
            background: {PRIMARY_HOVER_COLOR};
            border-color: {PRIMARY_HOVER_COLOR};
        }}
        QPushButton[primary="true"]:pressed, QDialogButtonBox QPushButton:default:pressed,
        QMessageBox QPushButton:default:pressed, QInputDialog QPushButton:default:pressed {{
            background: {PRIMARY_PRESSED_COLOR};
            border-color: {PRIMARY_PRESSED_COLOR};
        }}
        QPushButton[primary="true"]:focus, QDialogButtonBox QPushButton:default:focus,
        QMessageBox QPushButton:default:focus, QInputDialog QPushButton:default:focus {{
            border-color: {PRIMARY_PRESSED_COLOR};
        }}
        QPushButton[primary="true"]:disabled, QDialogButtonBox QPushButton:default:disabled,
        QMessageBox QPushButton:default:disabled, QInputDialog QPushButton:default:disabled {{
            color: {DISABLED_TEXT_COLOR};
            background: {DISABLED_BACKGROUND};
            border-color: {DISABLED_BORDER_COLOR};
        }}
        QToolButton[settingsButton="true"] {{
            color: {MUTED_TEXT_COLOR};
            background: transparent;
            border: 1px solid transparent;
            border-radius: 4px;
            padding: 0;
        }}
        QToolButton[settingsButton="true"]:hover {{
            color: {TEXT_COLOR};
            background: {SIDEBAR_BACKGROUND};
            border-color: {CONTROL_BORDER_COLOR};
        }}
        QToolButton[settingsButton="true"]:focus {{
            border-color: {CONTROL_FOCUS_COLOR};
        }}
        QToolButton[settingsButton="true"]:pressed {{
            color: {TEXT_COLOR};
            background: {SELECTION_BACKGROUND};
        }}
        QToolButton[settingsButton="true"]:disabled {{
            color: {DISABLED_TEXT_COLOR};
            background: transparent;
            border-color: transparent;
        }}
        QPushButton[navItem="true"] {{
            color: {MUTED_TEXT_COLOR};
            text-align: left;
            background: transparent;
            border: 1px solid transparent;
            border-radius: 4px;
            padding: 7px 9px;
        }}
        QPushButton[navItem="true"]:hover {{
            color: {TEXT_COLOR};
            background: {HOVER_BACKGROUND};
        }}
        QPushButton[navItem="true"]:checked {{
            color: {TEXT_COLOR};
            background: {SELECTION_BACKGROUND};
            font-weight: 600;
        }}
        QPushButton[navItem="true"]:checked:hover {{ background: {SELECTION_BACKGROUND}; }}
        QPushButton[navItem="true"]:focus {{ border-color: transparent; }}
        QCheckBox, QRadioButton {{ spacing: 8px; padding: 2px; background: transparent; }}
        QCheckBox::indicator, QRadioButton::indicator {{ width: 15px; height: 15px; }}
        QCheckBox:disabled, QRadioButton:disabled {{ color: {DISABLED_TEXT_COLOR}; }}
        QFrame#segmentedControl {{
            background: {SIDEBAR_BACKGROUND};
            border: 1px solid {CONTROL_BORDER_COLOR};
            border-radius: 4px;
        }}
        QPushButton[segmentedMode="true"] {{
            color: {MUTED_TEXT_COLOR};
            background: transparent;
            border: 1px solid transparent;
            border-radius: 4px;
            padding: 5px 12px;
        }}
        QPushButton[segmentedMode="true"]:hover {{
            color: {TEXT_COLOR};
            background: {SIDEBAR_BACKGROUND};
        }}
        QPushButton[segmentedMode="true"]:checked {{
            color: {TEXT_COLOR};
            background: {SELECTION_BACKGROUND};
            font-weight: 600;
        }}
        QPushButton[segmentedMode="true"]:focus {{ border: 1px dashed {CONTROL_FOCUS_COLOR}; }}
        QTabWidget::pane {{ border: none; background: transparent; }}
        QTabBar {{ background: transparent; }}
        QTabBar::tab {{
            color: {MUTED_TEXT_COLOR};
            background: transparent;
            border: 1px solid transparent;
            border-radius: 4px;
            margin: 0 2px 4px 0;
            padding: 5px 10px;
        }}
        QTabBar::tab:hover {{ color: {TEXT_COLOR}; background: {SIDEBAR_BACKGROUND}; }}
        QTabBar::tab:selected {{ color: {TEXT_COLOR}; background: {SELECTION_BACKGROUND}; font-weight: 600; }}
        QTableView {{
            background: {CARD_BACKGROUND};
            alternate-background-color: {SIDEBAR_BACKGROUND};
            border: 1px solid {BORDER_COLOR};
            gridline-color: {BORDER_COLOR};
            selection-background-color: {SELECTION_BACKGROUND};
            selection-color: {TEXT_COLOR};
        }}
        QHeaderView::section {{
            color: {MUTED_TEXT_COLOR};
            background: {SIDEBAR_BACKGROUND};
            border: none;
            border-right: 1px solid {BORDER_COLOR};
            border-bottom: 1px solid {BORDER_COLOR};
            padding: 4px 8px;
        }}
        QScrollBar:vertical {{ background: transparent; width: 8px; margin: 0; }}
        QScrollBar::handle:vertical {{ background: {BORDER_STRONG_COLOR}; min-height: 28px; border-radius: 3px; }}
        QScrollBar::handle:vertical:hover {{ background: {MUTED_TEXT_COLOR}; }}
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
        QScrollBar:horizontal {{ background: transparent; height: 8px; margin: 0; }}
        QScrollBar::handle:horizontal {{ background: {BORDER_STRONG_COLOR}; min-width: 28px; border-radius: 3px; }}
        QScrollBar::handle:horizontal:hover {{ background: {MUTED_TEXT_COLOR}; }}
        QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{ width: 0; }}
        QProgressBar {{
            color: {TEXT_COLOR};
            border: 1px solid {BORDER_COLOR};
            border-radius: 4px;
            background: {SIDEBAR_BACKGROUND};
            text-align: center;
        }}
        QProgressBar::chunk {{ background: {PROGRESS_COLOR}; }}
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


def section(title: str = "") -> tuple[QFrame, QVBoxLayout]:
    """Build a quiet surface with an internal heading instead of a fieldset."""

    box = QFrame()
    box.setObjectName("sectionCard")
    outer = QVBoxLayout(box)
    outer.setContentsMargins(12, 12, 12, 12)
    outer.setSpacing(8)
    if title:
        heading = QLabel(title)
        heading.setObjectName("sectionTitle")
        outer.addWidget(heading)
    layout = QVBoxLayout()
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(8)
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
    "ACCENT_COLOR",
    "ACCENT_FOREGROUND_COLOR",
    "APP_BACKGROUND",
    "AsyncPage",
    "BackgroundWorker",
    "BORDER_COLOR",
    "CARD_BACKGROUND",
    "MUTED_TEXT_COLOR",
    "NavigationButton",
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
