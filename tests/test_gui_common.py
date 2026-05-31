from __future__ import annotations

import unittest

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
