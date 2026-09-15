"""Shared layout and input helpers for Qt pages; no tool business rules."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)


EXCEL_FILTER = "Excel 文件 (*.xlsx *.xlsm);;所有文件 (*)"


def _scroll_page(content: QWidget) -> QScrollArea:
    area = QScrollArea()
    area.setWidgetResizable(True)
    area.setFrameShape(QScrollArea.Shape.NoFrame)
    area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
    area.setWidget(content)
    return area



def _add_action_bar(layout: QVBoxLayout, *buttons: QPushButton) -> QFrame:
    """Pin consistently sized page actions to a quiet bottom bar."""

    bar = QFrame()
    bar.setObjectName("pageActionBar")
    bar.setFixedHeight(46)
    row = QHBoxLayout(bar)
    row.setContentsMargins(2, 9, 6, 2)
    row.setSpacing(8)
    row.addStretch(1)
    for button in buttons:
        button.setFixedWidth(136 if button.property("primary") else 108)
        row.addWidget(button)
    layout.addWidget(bar)
    return bar



def _choose_excel(parent: QWidget, title: str, initial_dir: str = "") -> str:
    path, _ = QFileDialog.getOpenFileName(parent, title, initial_dir, EXCEL_FILTER)
    return path



def _set_combo_values(combo: QComboBox, values: tuple[str, ...], selected: str = "") -> str:
    combo.blockSignals(True)
    combo.clear()
    combo.addItems(values)
    chosen = selected if selected in values else (values[0] if values else "")
    if chosen:
        combo.setCurrentText(chosen)
    combo.blockSignals(False)
    return chosen
