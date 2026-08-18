#!/usr/bin/env python3
"""Small shared helpers for Tkinter GUI modules."""

from __future__ import annotations

import ctypes
from pathlib import Path
import sys
import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk

from PIL import Image, ImageDraw, ImageTk


PRIMARY_BUTTON_STYLE = "Tool.Primary.TButton"
SECTION_FRAME_STYLE = "Tool.Section.TLabelframe"
MUTED_LABEL_STYLE = "Tool.Muted.TLabel"

# Codex theme reference supplied by the user:
# accent=#cc7d5e, ink=#f9f9f7, surface=#2d2d2b, contrast=60.
# The remaining colors are opaque blends derived for native ttk widgets.
APP_WINDOW_BACKGROUND = "#2d2d2b"
APP_MAIN_BACKGROUND = "#2d2d2b"
APP_SIDEBAR_BACKGROUND = "#282826"
APP_INPUT_BACKGROUND = "#353532"
APP_HOVER_BACKGROUND = "#393936"
APP_SELECTED_BACKGROUND = "#493b35"
APP_BORDER = "#454541"
APP_BORDER_STRONG = "#595953"
APP_TEXT = "#f9f9f7"
APP_MUTED_TEXT = "#aaa9a3"
APP_PRIMARY = "#cc7d5e"
APP_PRIMARY_HOVER = "#d78b6c"
APP_ON_PRIMARY = "#241f1c"
APP_DISABLED_TEXT = "#777771"


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


def _checkbox_image(
    widget: tk.Misc,
    *,
    checked: bool,
    active: bool = False,
    disabled: bool = False,
    focused: bool = False,
    scale: float | None = None,
) -> ImageTk.PhotoImage:
    if scale is None:
        scale = _window_scale(widget)
    box_size = max(round(16 * scale), 16)
    image_width = max(round(21 * scale), 21)
    supersample = 4
    render_box = box_size * supersample
    render_width = image_width * supersample
    image = Image.new("RGBA", (render_width, render_box), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)

    inset = max(round(1.2 * scale * supersample), supersample)
    radius = max(round(4 * scale * supersample), 4)
    outline_width = max(
        round((2.0 if focused else 1.25) * scale * supersample),
        supersample,
    )
    if disabled:
        fill = APP_MAIN_BACKGROUND
        outline = APP_BORDER
    elif checked:
        fill = APP_PRIMARY_HOVER if active else APP_PRIMARY
        outline = APP_TEXT if focused else fill
    else:
        fill = APP_INPUT_BACKGROUND
        outline = APP_PRIMARY if active or focused else APP_BORDER_STRONG

    draw.rounded_rectangle(
        (inset, inset, render_box - inset - 1, render_box - inset - 1),
        radius=radius,
        fill=fill,
        outline=outline,
        width=outline_width,
    )
    if checked:
        check_color = APP_DISABLED_TEXT if disabled else APP_ON_PRIMARY
        check_width = max(round(1.8 * scale * supersample), 2 * supersample)
        points = (
            (render_box * 0.28, render_box * 0.52),
            (render_box * 0.44, render_box * 0.68),
            (render_box * 0.74, render_box * 0.34),
        )
        draw.line(points, fill=check_color, width=check_width, joint="curve")

    image = image.resize((image_width, box_size), Image.Resampling.LANCZOS)
    return ImageTk.PhotoImage(image, master=widget)


def _window_scale(widget: tk.Misc) -> float:
    """Return the physical-pixel scale for the widget's current monitor."""

    root = widget.winfo_toplevel()
    if sys.platform == "win32":
        try:
            get_dpi_for_window = ctypes.windll.user32.GetDpiForWindow
            get_dpi_for_window.argtypes = [ctypes.c_void_p]
            get_dpi_for_window.restype = ctypes.c_uint
            dpi = int(get_dpi_for_window(ctypes.c_void_p(root.winfo_id())))
            if dpi > 0:
                return max(dpi / 96.0, 1.0)
        except (AttributeError, OSError, tk.TclError, TypeError, ValueError):
            pass

    try:
        return max(float(widget.winfo_fpixels("1i")) / 96.0, 1.0)
    except (tk.TclError, TypeError, ValueError):
        return max(float(widget.tk.call("tk", "scaling")) * 72.0 / 96.0, 1.0)


def _sync_checkbox_style(widget: tk.Misc, style: ttk.Style, state: dict) -> None:
    scale_key = max(round(_window_scale(widget) * 96.0), 96)
    if state["scale_key"] == scale_key:
        return

    if scale_key not in state["assets"]:
        scale = scale_key / 96.0
        assets = {
            "unchecked": _checkbox_image(widget, checked=False, scale=scale),
            "unchecked_active": _checkbox_image(
                widget, checked=False, active=True, scale=scale
            ),
            "unchecked_focus": _checkbox_image(
                widget, checked=False, focused=True, scale=scale
            ),
            "checked": _checkbox_image(widget, checked=True, scale=scale),
            "checked_active": _checkbox_image(
                widget, checked=True, active=True, scale=scale
            ),
            "checked_focus": _checkbox_image(
                widget, checked=True, focused=True, scale=scale
            ),
            "unchecked_disabled": _checkbox_image(
                widget, checked=False, disabled=True, scale=scale
            ),
            "checked_disabled": _checkbox_image(
                widget, checked=True, disabled=True, scale=scale
            ),
        }
        element_name = f"Tool.Checkbutton.indicator.{scale_key}"
        state["assets"][scale_key] = assets
        state["elements"][scale_key] = element_name
        style.element_create(
            element_name,
            "image",
            assets["unchecked"],
            ("disabled", "selected", assets["checked_disabled"]),
            ("disabled", assets["unchecked_disabled"]),
            ("focus", "selected", assets["checked_focus"]),
            ("focus", assets["unchecked_focus"]),
            ("active", "selected", assets["checked_active"]),
            ("selected", assets["checked"]),
            ("active", assets["unchecked_active"]),
            sticky="",
        )

    element_name = state["elements"][scale_key]
    style.layout(
        "TCheckbutton",
        [
            (
                "Checkbutton.padding",
                {
                    "sticky": "nsew",
                    "children": [
                        (element_name, {"side": "left", "sticky": ""}),
                        (
                            "Checkbutton.focus",
                            {
                                "side": "left",
                                "sticky": "nsew",
                                "children": [
                                    ("Checkbutton.label", {"sticky": "nsew"})
                                ],
                            },
                        ),
                    ],
                },
            )
        ],
    )
    state["scale_key"] = scale_key


def _configure_checkbox_style(widget: tk.Misc, style: ttk.Style) -> None:
    root = widget.winfo_toplevel()
    state_name = "_toolshub_checkbox_style_state"
    state = getattr(root, state_name, None)
    if state is None:
        state = {
            "after_id": None,
            "assets": {},
            "elements": {},
            "scale_key": None,
        }
        setattr(root, state_name, state)

        def schedule_dpi_refresh(event: tk.Event) -> None:
            if event.widget is not root:
                return
            if state["after_id"] is not None:
                root.after_cancel(state["after_id"])

            def refresh() -> None:
                state["after_id"] = None
                if root.winfo_exists():
                    _sync_checkbox_style(root, ttk.Style(root), state)

            state["after_id"] = root.after(75, refresh)

        root.bind("<Configure>", schedule_dpi_refresh, add="+")

    _sync_checkbox_style(widget, style, state)


def configure_tool_page_style(widget: tk.Misc) -> None:
    """Configure the shared Codex-inspired visual language for tool pages."""
    style = ttk.Style(widget)
    if "clam" in style.theme_names() and style.theme_use() != "clam":
        style.theme_use("clam")

    fonts = _font_specs(widget)
    root = widget.winfo_toplevel()
    root.configure(background=APP_WINDOW_BACKGROUND)
    root.option_add("*selectBackground", APP_PRIMARY)
    root.option_add("*selectForeground", APP_ON_PRIMARY)
    root.option_add("*TCombobox*Listbox.background", APP_INPUT_BACKGROUND)
    root.option_add("*TCombobox*Listbox.foreground", APP_TEXT)

    style.configure(".", font=fonts["body"])
    style.configure("TFrame", background=APP_MAIN_BACKGROUND)
    style.configure(
        "TLabel",
        background=APP_MAIN_BACKGROUND,
        foreground=APP_TEXT,
        font=fonts["body"],
    )
    style.configure(
        "TButton",
        anchor="center",
        background=APP_INPUT_BACKGROUND,
        bordercolor=APP_BORDER_STRONG,
        borderwidth=1,
        darkcolor=APP_BORDER_STRONG,
        focuscolor=APP_BORDER_STRONG,
        focusthickness=1,
        foreground=APP_TEXT,
        lightcolor=APP_BORDER_STRONG,
        padding=(11, 7),
        relief="flat",
    )
    style.map(
        "TButton",
        background=[
            ("pressed", APP_SELECTED_BACKGROUND),
            ("active", APP_HOVER_BACKGROUND),
        ],
        bordercolor=[("focus", APP_PRIMARY), ("!focus", APP_BORDER_STRONG)],
        foreground=[("disabled", APP_DISABLED_TEXT), ("!disabled", APP_TEXT)],
    )
    style.configure(
        PRIMARY_BUTTON_STYLE,
        background=APP_PRIMARY,
        bordercolor=APP_PRIMARY,
        darkcolor=APP_PRIMARY,
        focuscolor=APP_PRIMARY,
        font=fonts["label"],
        foreground=APP_ON_PRIMARY,
        lightcolor=APP_PRIMARY,
        padding=(16, 9),
        relief="flat",
    )
    style.map(
        PRIMARY_BUTTON_STYLE,
        background=[
            ("pressed", APP_PRIMARY_HOVER),
            ("active", APP_PRIMARY_HOVER),
        ],
        bordercolor=[
            ("pressed", APP_PRIMARY_HOVER),
            ("active", APP_PRIMARY_HOVER),
        ],
        foreground=[("disabled", APP_DISABLED_TEXT), ("!disabled", APP_ON_PRIMARY)],
    )
    style.configure(
        "TEntry",
        background=APP_INPUT_BACKGROUND,
        bordercolor=APP_BORDER_STRONG,
        darkcolor=APP_BORDER_STRONG,
        fieldbackground=APP_INPUT_BACKGROUND,
        foreground=APP_TEXT,
        insertcolor=APP_TEXT,
        lightcolor=APP_BORDER_STRONG,
        padding=(9, 7),
    )
    style.map(
        "TEntry",
        bordercolor=[("focus", APP_PRIMARY), ("!focus", APP_BORDER_STRONG)],
        lightcolor=[("focus", APP_PRIMARY), ("!focus", APP_BORDER_STRONG)],
        darkcolor=[("focus", APP_PRIMARY), ("!focus", APP_BORDER_STRONG)],
    )
    for control_style in ("TCombobox", "TSpinbox"):
        style.configure(
            control_style,
            arrowcolor=APP_MUTED_TEXT,
            background=APP_INPUT_BACKGROUND,
            bordercolor=APP_BORDER_STRONG,
            darkcolor=APP_BORDER_STRONG,
            fieldbackground=APP_INPUT_BACKGROUND,
            foreground=APP_TEXT,
            lightcolor=APP_BORDER_STRONG,
            padding=(8, 6),
        )
        style.map(
            control_style,
            bordercolor=[("focus", APP_PRIMARY), ("!focus", APP_BORDER_STRONG)],
            fieldbackground=[
                ("readonly", APP_INPUT_BACKGROUND),
                ("disabled", APP_HOVER_BACKGROUND),
            ],
            foreground=[("disabled", APP_DISABLED_TEXT), ("!disabled", APP_TEXT)],
        )
    for selector_style in ("TCheckbutton", "TRadiobutton"):
        style.configure(
            selector_style,
            background=APP_MAIN_BACKGROUND,
            focuscolor=APP_PRIMARY,
            focusthickness=1,
            foreground=APP_TEXT,
            indicatorbackground=APP_INPUT_BACKGROUND,
            indicatorforeground=APP_TEXT,
            padding=(0, 3),
        )
        style.map(
            selector_style,
            background=[("active", APP_MAIN_BACKGROUND)],
            foreground=[("disabled", APP_DISABLED_TEXT), ("!disabled", APP_TEXT)],
            indicatorbackground=[
                ("selected", APP_PRIMARY),
                ("disabled", APP_BORDER),
                ("!selected", APP_INPUT_BACKGROUND),
            ],
            indicatorforeground=[("selected", APP_ON_PRIMARY)],
        )
    style.configure(
        "Vertical.TScrollbar",
        arrowsize=11,
        background=APP_BORDER_STRONG,
        bordercolor=APP_MAIN_BACKGROUND,
        darkcolor=APP_BORDER_STRONG,
        lightcolor=APP_BORDER_STRONG,
        troughcolor=APP_MAIN_BACKGROUND,
    )
    style.configure(
        "TSeparator",
        background=APP_BORDER,
        bordercolor=APP_BORDER,
    )
    style.configure(
        SECTION_FRAME_STYLE,
        background=APP_MAIN_BACKGROUND,
        bordercolor=APP_BORDER,
        borderwidth=1,
        darkcolor=APP_BORDER,
        lightcolor=APP_BORDER,
        padding=16,
        relief="solid",
    )
    style.configure(
        f"{SECTION_FRAME_STYLE}.Label",
        background=APP_MAIN_BACKGROUND,
        foreground=APP_TEXT,
        font=fonts["section"],
    )
    style.configure(
        MUTED_LABEL_STYLE,
        background=APP_MAIN_BACKGROUND,
        foreground=APP_MUTED_TEXT,
        font=fonts["small"],
    )
    _configure_checkbox_style(widget, style)


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
