from __future__ import annotations

import io
import os
import runpy
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest.mock import Mock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QScrollArea
from phraseloom.gui import main
from phraseloom.qt_page import PhraseLoomPage
from phraseloom.strings_workflow import export_strings_workbook, restore_strings_workbook


class PhraseLoomGuiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self) -> None:
        self.page = PhraseLoomPage()
        self.addCleanup(self.page.close)
        self.page.run_in_background = Mock()

    def test_gui_entry_opens_the_shared_qt_page(self) -> None:
        with patch("toolshub_gui.main", return_value=0) as launch:
            self.assertEqual(main(), 0)
        launch.assert_called_once_with(["--tool", "phraseloom"])

    def test_gui_module_remains_callable_as_a_direct_script(self) -> None:
        path = Path(__file__).parents[2] / "phraseloom" / "gui.py"
        with patch("toolshub_gui.main", return_value=0) as launch, patch("sys.argv", [str(path)]):
            with self.assertRaises(SystemExit) as result:
                runpy.run_path(str(path), run_name="__main__")
        self.assertEqual(result.exception.code, 0)
        launch.assert_called_once_with(["--tool", "phraseloom"])

    def test_cli_dispatches_gui_to_qt_entry(self) -> None:
        from phraseloom.cli import _dispatch
        with patch("toolshub_gui.main", return_value=0) as launch:
            self.assertEqual(_dispatch(["gui"]), 0)
        launch.assert_called_once_with(["--tool", "phraseloom"])

    def test_top_level_help_keeps_export_restore_and_gui(self) -> None:
        from phraseloom.cli import _dispatch
        with redirect_stdout(io.StringIO()) as output:
            self.assertEqual(_dispatch(["--help"]), 0)
        for command in ("gui", "export", "restore"):
            self.assertIn(f"phraseloom {command}", output.getvalue())
        self.assertNotIn("TM", output.getvalue())

    def test_qt_page_keeps_export_and_secondary_restore_in_one_workspace(self) -> None:
        self.assertEqual(self.page.export_button.text(), "导出 Strings")
        self.assertEqual(self.page.restore_button.text(), "回填译文…")
        self.assertIsNotNone(self.page.findChild(QScrollArea))
        self.assertTrue(self.page.split_lines.isChecked())
        self.assertFalse(self.page.group_similar.isChecked())
        self.assertEqual(self.page.source_column.text(), "source")
        self.assertEqual(self.page.target_column.text(), "target")

    def test_export_forwards_defaults_and_custom_options(self) -> None:
        self.page.input_picker.set_path("/tmp/source.xlsx")
        self.page.run_export()
        call = self.page.run_in_background.call_args
        self.assertIs(call.args[0], export_strings_workbook)
        self.assertEqual(call.kwargs["args"], ("/tmp/source.xlsx",))
        self.assertEqual(call.kwargs["kwargs"], {
            "source_col": "source", "target_col": "target", "context_col": None,
            "group_similar": False, "tag_config": None, "split_lines": True,
        })
        self.page.source_column.setText("en")
        self.page.target_column.setText("fr")
        self.page.context_column.setText("screen")
        self.page.group_similar.setChecked(True)
        self.page.split_lines.setChecked(False)
        self.page.tag_picker.set_path("/tmp/tags.toml")
        self.page.run_export()
        self.assertEqual(self.page.run_in_background.call_args.kwargs["kwargs"], {
            "source_col": "en", "target_col": "fr", "context_col": "screen",
            "group_similar": True, "tag_config": "/tmp/tags.toml", "split_lines": False,
        })

    def test_missing_input_is_rejected_before_starting_work(self) -> None:
        with patch("phraseloom.qt_page.show_error") as error:
            self.page.run_export()
        error.assert_called_once()
        self.page.run_in_background.assert_not_called()

    def test_restore_uses_only_selected_strings_file_and_cancel_does_nothing(self) -> None:
        with patch("phraseloom.qt_page._choose_excel", return_value=""):
            self.page.choose_and_restore()
        self.page.run_in_background.assert_not_called()
        with (
            patch("phraseloom.qt_page._choose_excel", return_value="/tmp/source_strings.xlsx"),
            patch("phraseloom.qt_page.default_restored_output_path", return_value=Path("/tmp/source_translated.xlsx")),
        ):
            self.page.choose_and_restore()
        call = self.page.run_in_background.call_args
        self.assertIs(call.args[0], restore_strings_workbook)
        self.assertEqual(call.kwargs["args"], ("/tmp/source_strings.xlsx",))
        self.assertIn("source_translated.xlsx", self.page.preview.text())
        self.assertFalse(self.page.export_button.isEnabled())
        self.assertFalse(self.page.restore_button.isEnabled())

    def test_error_reenables_actions(self) -> None:
        self.page._set_running(True)
        with patch("phraseloom.qt_page.show_error"):
            self.page._finish_error("失败", "test")
        self.assertTrue(self.page.export_button.isEnabled())
        self.assertTrue(self.page.restore_button.isEnabled())

    def test_optional_tag_config_can_be_cleared(self) -> None:
        self.page.tag_picker.set_path("/tmp/tags.toml")
        self.page.tag_picker.clear_button.click()
        self.assertEqual(self.page.tag_picker.path(), "")


if __name__ == "__main__":
    unittest.main()
