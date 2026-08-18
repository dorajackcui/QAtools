from __future__ import annotations

import unittest
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
    def test_checkbox_style_rebuilds_for_a_new_monitor_dpi_and_keeps_focus_state(
        self,
    ) -> None:
        widget = Mock()
        style = Mock()
        state = {
            "after_id": None,
            "assets": {},
            "elements": {},
            "scale_key": None,
        }
        generated_images = [f"image-{index}" for index in range(16)]

        with (
            patch.object(gui_common, "_window_scale", side_effect=(1.0, 1.0, 1.5)),
            patch.object(
                gui_common,
                "_checkbox_image",
                side_effect=generated_images,
            ) as checkbox_image,
        ):
            gui_common._sync_checkbox_style(widget, style, state)
            gui_common._sync_checkbox_style(widget, style, state)
            gui_common._sync_checkbox_style(widget, style, state)

        self.assertEqual(checkbox_image.call_count, 16)
        self.assertEqual(set(state["assets"]), {96, 144})
        self.assertEqual(style.element_create.call_count, 2)
        self.assertEqual(style.layout.call_count, 2)

        first_element_args = style.element_create.call_args_list[0].args
        self.assertIn(
            ("focus", "selected", state["assets"][96]["checked_focus"]),
            first_element_args,
        )
        self.assertIn("Checkbutton.focus", repr(style.layout.call_args_list[0]))

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
