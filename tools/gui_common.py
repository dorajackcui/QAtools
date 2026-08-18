#!/usr/bin/env python3
"""Small shared helpers for Tkinter GUI modules."""

from __future__ import annotations

from pathlib import Path
import sys
import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk

import sv_ttk


PRIMARY_BUTTON_STYLE = "Accent.TButton"
SECTION_FRAME_STYLE = "Tool.Section.TLabelframe"
MUTED_LABEL_STYLE = "Tool.Muted.TLabel"

# Sun Valley dark theme palette. Custom colors are limited to the application
# shell; standard controls are rendered by sv-ttk itself.
APP_WINDOW_BACKGROUND = "#1c1c1c"
APP_MAIN_BACKGROUND = "#1c1c1c"
APP_SIDEBAR_BACKGROUND = "#181818"
APP_INPUT_BACKGROUND = "#2b2b2b"
APP_TEXT = "#fafafa"
APP_MUTED_TEXT = "#9e9e9e"
APP_SELECTION_BACKGROUND = "#2f60d8"
APP_SELECTION_TEXT = "#ffffff"


def _font_specs(widget: tk.Misc) -> dict[str, tuple[str, int, str]]:
    default_font = tkfont.nametofont("TkDefaultFont", root=widget)
    family = str(default_font.actual("family"))
    native_size = abs(int(default_font.actual("size"))) or 10
    body_size = max(native_size, 10) if sys.platform == "win32" else native_size
    return {
        "body": (family, body_size, "normal"),
        "small": (family, max(body_size - 1, 8), "normal"),
        "label": (family, body_size, "bold"),
        "section": (family, body_size + 1, "bold"),
        "title": (family, body_size + 8, "bold"),
        "brand": (family, body_size + 3, "bold"),
    }


def _apply_sun_valley_theme(root: tk.Tk) -> None:
    """Apply the shared theme once per Tcl/Tk interpreter."""

    if getattr(root, "_toolshub_sun_valley_theme_applied", False):
        return
    sv_ttk.set_theme("dark", root=root)
    setattr(root, "_toolshub_sun_valley_theme_applied", True)


def configure_tool_page_style(widget: tk.Misc) -> None:
    """Apply Sun Valley dark and the small set of shared app styles."""

    tk_root = widget._root()
    _apply_sun_valley_theme(tk_root)

    style = ttk.Style(widget)
    fonts = _font_specs(widget)
    window = widget.winfo_toplevel()
    window.configure(background=APP_WINDOW_BACKGROUND)
    window.option_add("*selectBackground", APP_SELECTION_BACKGROUND)
    window.option_add("*selectForeground", APP_SELECTION_TEXT)
    window.option_add("*TCombobox*Listbox.background", APP_INPUT_BACKGROUND)
    window.option_add("*TCombobox*Listbox.foreground", APP_TEXT)

    style.configure(".", font=fonts["body"])
    style.configure(
        PRIMARY_BUTTON_STYLE,
        font=fonts["label"],
        padding=(16, 8),
    )
    style.configure(
        SECTION_FRAME_STYLE,
        padding=14,
    )
    style.configure(
        f"{SECTION_FRAME_STYLE}.Label",
        font=fonts["section"],
    )
    style.configure(
        MUTED_LABEL_STYLE,
        foreground=APP_MUTED_TEXT,
        font=fonts["small"],
    )


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
        padding=16,
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
