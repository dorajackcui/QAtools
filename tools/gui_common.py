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
FILE_PATH_DISPLAY_WIDTH = 38

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


def _preferred_ui_font_family(widget: tk.Misc) -> str:
    available_families = {
        str(family).casefold(): str(family)
        for family in tkfont.families(root=widget)
    }
    if sys.platform == "win32":
        candidates = ("Microsoft YaHei UI", "Segoe UI", "Arial")
    elif sys.platform == "darwin":
        candidates = ("PingFang SC", "Helvetica Neue", "Arial")
    else:
        candidates = ("Noto Sans CJK SC", "Noto Sans", "DejaVu Sans")

    for candidate in candidates:
        matched_family = available_families.get(candidate.casefold())
        if matched_family is not None:
            return matched_family

    default_font = tkfont.nametofont("TkDefaultFont", root=widget)
    return str(default_font.actual("family"))


def _configure_ui_fonts(widget: tk.Misc) -> None:
    """Replace unavailable Sun Valley fonts with a native CJK UI family."""

    family = _preferred_ui_font_family(widget)
    font_specs = {
        "SunValleyCaptionFont": (9, "normal"),
        "SunValleyBodyFont": (10, "normal"),
        "SunValleyBodyStrongFont": (10, "bold"),
        "SunValleyBodyLargeFont": (12, "normal"),
        "SunValleySubtitleFont": (14, "bold"),
        "SunValleyTitleFont": (18, "bold"),
        "SunValleyTitleLargeFont": (24, "bold"),
        "SunValleyDisplayFont": (36, "bold"),
    }
    for font_name, (size, weight) in font_specs.items():
        theme_font = tkfont.nametofont(font_name, root=widget)
        theme_font.configure(family=family, size=size, weight=weight)

    for font_name in ("TkDefaultFont", "TkTextFont", "TkMenuFont"):
        system_font = tkfont.nametofont(font_name, root=widget)
        font_options: dict[str, object] = {"family": family}
        if sys.platform == "win32":
            font_options["size"] = 10
        system_font.configure(**font_options)


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
    _configure_ui_fonts(tk_root)

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


def _selected_file_name(value: object) -> str:
    """Return a short picker status while keeping the full path in its variable."""

    path_text = str(value or "").strip()
    if not path_text:
        return "未选择文件"
    return path_text.replace("\\", "/").rsplit("/", 1)[-1] or path_text


def create_file_path_display(
    parent: tk.Misc,
    *,
    variable: tk.Variable,
    width: int = FILE_PATH_DISPLAY_WIDTH,
) -> ttk.Label:
    """Create a muted, non-editable filename display for a file picker."""

    display_variable = tk.StringVar(master=parent)

    def sync_display(*_args: object) -> None:
        display_variable.set(_selected_file_name(variable.get()))

    sync_display()
    trace_id = variable.trace_add("write", sync_display)
    label = ttk.Label(
        parent,
        textvariable=display_variable,
        width=width,
        anchor="w",
        style=MUTED_LABEL_STYLE,
    )
    # The source variable retains the complete path used by the tool.
    label._file_path_source_variable = variable  # type: ignore[attr-defined]
    label._file_path_display_variable = display_variable  # type: ignore[attr-defined]
    label._file_path_trace_id = trace_id  # type: ignore[attr-defined]

    def remove_trace(event: tk.Event) -> None:
        if event.widget is not label:
            return
        try:
            variable.trace_remove("write", trace_id)
        except tk.TclError:
            pass

    label.bind("<Destroy>", remove_trace, add="+")
    return label


def refresh_top_aligned_scroll_region(canvas: tk.Canvas) -> None:
    """Keep short canvas content pinned to the top of its viewport."""

    content_bounds = canvas.bbox("all")
    if not content_bounds:
        return

    left, top, right, bottom = content_bounds
    viewport_width = max(canvas.winfo_width(), 1)
    viewport_height = max(canvas.winfo_height(), 1)
    content_height = bottom - top
    canvas.configure(
        scrollregion=(
            min(left, 0),
            min(top, 0),
            max(right, viewport_width),
            max(bottom, viewport_height),
        )
    )
    if content_height <= viewport_height:
        canvas.yview_moveto(0.0)


def add_optional_status_label(
    parent: tk.Misc,
    *,
    variable: tk.Variable,
    row: int,
    column: int = 0,
    columnspan: int = 1,
    pady: tuple[int, int] = (8, 0),
) -> ttk.Label:
    """Show a muted status row only while its variable contains text."""

    label = ttk.Label(
        parent,
        textvariable=variable,
        style=MUTED_LABEL_STYLE,
    )
    label.grid(
        row=row,
        column=column,
        columnspan=columnspan,
        sticky="w",
        pady=pady,
    )

    def sync_visibility(*_args: object) -> None:
        if str(variable.get()).strip():
            label.grid()
        else:
            label.grid_remove()

    sync_visibility()
    trace_id = variable.trace_add("write", sync_visibility)

    def remove_trace(event: tk.Event) -> None:
        if event.widget is not label:
            return
        try:
            variable.trace_remove("write", trace_id)
        except tk.TclError:
            pass

    label.bind("<Destroy>", remove_trace, add="+")
    return label


def add_file_picker_row(
    parent: tk.Misc,
    *,
    label: str,
    variable,
    command,
) -> ttk.Label:
    """Add a button-driven picker with a non-editable filename status."""
    ttk.Label(parent, text=label).grid(row=0, column=0, sticky="w")
    path_display = create_file_path_display(parent, variable=variable)
    path_display.grid(row=0, column=1, sticky="w", padx=(12, 8))
    ttk.Button(parent, text="选择文件", command=command).grid(
        row=0,
        column=2,
        sticky="ew",
    )
    parent.columnconfigure(1, weight=0)
    return path_display


class OutputPreviewMixin:
    """Share typed-path refresh and automatic output-name previews."""

    output_path_builder = None

    def handle_input_file_focus_out(self, _event: object | None = None) -> None:
        self.refresh_sheet_choices(show_error=False)
        self.update_output_preview()

    def update_output_preview(self) -> None:
        input_file = self.input_file_var.get().strip()
        if not input_file:
            self.output_preview_var.set("")
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
