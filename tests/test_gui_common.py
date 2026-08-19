from __future__ import annotations

import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

from tools import gui_common
from tools.gui_common import parse_positive_int, set_combobox_values


class FakeVar:
    def __init__(self, value: str = "") -> None:
        self.value = value

    def get(self) -> str:
        return self.value

    def set(self, value: str) -> None:
        self.value = value


class FakeCombobox(dict):
    def __init__(self) -> None:
        super().__init__()
        self["values"] = ()


class GuiCommonTests(unittest.TestCase):
    def test_sun_valley_theme_is_applied_once_per_root(self) -> None:
        root = SimpleNamespace()

        with patch.object(gui_common.sv_ttk, "set_theme") as set_theme:
            gui_common._apply_sun_valley_theme(root)
            gui_common._apply_sun_valley_theme(root)

        set_theme.assert_called_once_with("dark", root=root)
        self.assertTrue(root._toolshub_sun_valley_theme_applied)

    def test_typography_is_derived_from_the_native_base_size(self) -> None:
        font_names = (
            gui_common.BODY_FONT,
            gui_common.SMALL_FONT,
            gui_common.CATEGORY_FONT,
            gui_common.LABEL_FONT,
            gui_common.SECTION_FONT,
            gui_common.TITLE_FONT,
            gui_common.BRAND_FONT,
            "SunValleyCaptionFont",
            "SunValleyBodyFont",
            "SunValleyBodyStrongFont",
            "SunValleyBodyLargeFont",
            "SunValleySubtitleFont",
            "SunValleyTitleFont",
            "SunValleyTitleLargeFont",
            "SunValleyDisplayFont",
            "TkDefaultFont",
            "TkTextFont",
            "TkMenuFont",
        )
        fonts = {name: Mock() for name in font_names}
        fonts["TkDefaultFont"].actual.return_value = 13

        with (
            patch.object(
                gui_common,
                "_preferred_ui_font_family",
                return_value="PingFang SC",
            ),
            patch.object(
                gui_common.tkfont,
                "nametofont",
                side_effect=lambda name, root: fonts[name],
            ),
        ):
            gui_common._configure_ui_typography(SimpleNamespace())

        for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont"):
            fonts[name].configure.assert_called_once_with(family="PingFang SC")
        fonts[gui_common.BODY_FONT].configure.assert_called_once_with(
            family="PingFang SC",
            size=13,
            weight="normal",
        )
        fonts[gui_common.TITLE_FONT].configure.assert_called_once_with(
            family="PingFang SC",
            size=21,
            weight="bold",
        )

    def test_windows_uses_one_system_dpi_scale(self) -> None:
        dpi_api = Mock()
        fake_windll = SimpleNamespace(
            shcore=SimpleNamespace(SetProcessDpiAwareness=dpi_api)
        )

        with (
            patch.object(gui_common.sys, "platform", "win32"),
            patch.object(gui_common.ctypes, "windll", fake_windll, create=True),
        ):
            gui_common.configure_system_dpi_awareness()

        dpi_api.assert_called_once_with(1)

    def test_application_root_applies_dpi_policy_before_creating_tk(self) -> None:
        calls: list[str] = []
        root = object()
        with (
            patch.object(
                gui_common,
                "configure_system_dpi_awareness",
                side_effect=lambda: calls.append("dpi"),
            ),
            patch.object(
                gui_common.tk,
                "Tk",
                side_effect=lambda: calls.append("tk") or root,
            ),
        ):
            self.assertIs(gui_common.create_application_root(), root)

        self.assertEqual(calls, ["dpi", "tk"])

    def test_parse_positive_int_accepts_blank_default(self) -> None:
        self.assertEqual(parse_positive_int("", default=2, field_name="开始行"), 2)

    def test_parse_positive_int_rejects_non_integer(self) -> None:
        with self.assertRaisesRegex(ValueError, "开始行必须是整数"):
            parse_positive_int("x", default=2, field_name="开始行")

    def test_parse_positive_int_rejects_zero(self) -> None:
        with self.assertRaisesRegex(ValueError, "开始行必须大于 0"):
            parse_positive_int("0", default=2, field_name="开始行")

    def test_set_combobox_values_keeps_selected_value_when_present(self) -> None:
        combobox = FakeCombobox()
        variable = FakeVar("B")

        selected = set_combobox_values(combobox, ("A", "B"), variable, default_value="A")

        self.assertEqual(selected, "B")
        self.assertEqual(variable.get(), "B")
        self.assertEqual(combobox["values"], ("A", "B"))

    def test_set_combobox_values_uses_default_when_selection_missing(self) -> None:
        combobox = FakeCombobox()
        variable = FakeVar("Missing")

        selected = set_combobox_values(combobox, ("A", "B"), variable, default_value="B")

        self.assertEqual(selected, "B")
        self.assertEqual(variable.get(), "B")


if __name__ == "__main__":
    unittest.main()
