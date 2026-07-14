#!/usr/bin/env python3
"""Small shared helpers for Tkinter GUI modules."""

from __future__ import annotations

from pathlib import Path
import tkinter as tk
from tkinter import ttk


PRIMARY_BUTTON_STYLE = "Tool.Primary.TButton"
SECTION_FRAME_STYLE = "Tool.Section.TLabelframe"
MUTED_LABEL_STYLE = "Tool.Muted.TLabel"


def configure_tool_page_style(widget: tk.Misc) -> None:
    """Configure the small shared visual language used by every tool page."""
    style = ttk.Style(widget)
    style.configure(
        PRIMARY_BUTTON_STYLE,
        font=("TkDefaultFont", 10, "bold"),
        padding=(12, 8),
    )
    style.configure(
        f"{SECTION_FRAME_STYLE}.Label",
        font=("TkDefaultFont", 10, "bold"),
    )
    style.configure(MUTED_LABEL_STYLE, foreground="#555555")


def create_section(
    parent: tk.Misc,
    *,
    title: str,
    row: int,
    pady: tuple[int, int] = (0, 10),
) -> ttk.LabelFrame:
    """Create a full-width section matching the workflow page."""
    section = ttk.LabelFrame(
        parent,
        text=title,
        padding=12,
        style=SECTION_FRAME_STYLE,
    )
    section.grid(row=row, column=0, sticky="ew", pady=pady)
    return section


def add_file_picker_row(
    parent: tk.Misc,
    *,
    label: str,
    variable,
    command,
    focus_out_command=None,
) -> ttk.Entry:
    """Add the common full-width file picker row to a section."""
    ttk.Label(parent, text=label).grid(row=0, column=0, sticky="w")
    entry = ttk.Entry(parent, textvariable=variable, width=56)
    entry.grid(row=0, column=1, sticky="ew", padx=(12, 8))
    if focus_out_command is not None:
        entry.bind("<FocusOut>", focus_out_command)
    ttk.Button(parent, text="选择", command=command).grid(
        row=0,
        column=2,
        sticky="ew",
    )
    parent.columnconfigure(1, weight=1)
    return entry


class OutputPreviewMixin:
    """Share typed-path refresh and automatic output-name previews."""

    output_path_builder = None

    def handle_input_file_focus_out(self, _event: object | None = None) -> None:
        self.refresh_sheet_choices(show_error=False)
        self.update_output_preview()

    def update_output_preview(self) -> None:
        input_file = self.input_file_var.get().strip()
        if not input_file:
            self.output_preview_var.set("输出文件：选择输入 Excel 后自动生成")
            return
        if self.output_path_builder is None:
            raise RuntimeError("未配置默认输出路径生成器。")
        output_name = Path(self.output_path_builder(Path(input_file))).name
        self.output_preview_var.set(f"输出文件：{output_name}")


def parse_positive_int(raw_value: str, *, default: int, field_name: str) -> int:
    value = raw_value.strip() or str(default)
    try:
        parsed = int(value)
    except ValueError as exc:
        raise ValueError(f"{field_name}必须是整数。") from exc
    if parsed < 1:
        raise ValueError(f"{field_name}必须大于 0。")
    return parsed


def set_combobox_values(combobox, values: tuple[str, ...], variable, default_value: str | None) -> str:
    combobox["values"] = values
    selected = variable.get().strip()
    if selected not in values:
        selected = default_value or (values[0] if values else "")
        variable.set(selected)
    return selected
