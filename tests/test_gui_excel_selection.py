from __future__ import annotations

import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from openpyxl import Workbook
from PySide6.QtWidgets import QApplication
from toolshub_gui import ToolshubApp
from tools.tb_projects import TbProject


class GuiSheetSelectionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self) -> None:
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.directory = Path(temporary.name)
        self.enterContext(patch("tools.workflow.gui_options.default_header_aliases_path",
                                return_value=self.directory / "aliases.json"))
        self.enterContext(patch("tools.tb_projects.default_tb_projects_path",
                                return_value=self.directory / "projects.json"))
        self.window = ToolshubApp(show_window=False)
        self.addCleanup(self.window.close)
        self.workflow = self.window.tool_frames["workflow"]
        self.workflow.run_in_background = Mock()
        self.input = self.directory / "input.xlsx"
        workbook = Workbook()
        try:
            workbook.active.title = "Data"
            workbook.active.append(["source", "target"])
            alternate = workbook.create_sheet("Alternate")
            alternate["C1"] = " SOURCE "
            alternate["E1"] = " target "
            workbook.create_sheet("NoHeader")
            workbook.active = 1
            workbook.save(self.input)
        finally:
            workbook.close()

    def test_qa_loads_active_sheet_detects_columns_and_updates_on_switch(self) -> None:
        self.workflow.load_input_file(str(self.input))
        self.assertEqual(self.workflow.sheet.currentText(), "Alternate")
        self.assertEqual(self.workflow.source_column.text(), "C")
        self.assertEqual(self.workflow.target_column.text(), "E")
        self.workflow.sheet.setCurrentText("Data")
        self.assertEqual(self.workflow.source_column.text(), "A")
        self.assertEqual(self.workflow.target_column.text(), "B")
        self.assertIn("workflow_check_input.xlsx", self.workflow.output_preview.text())
        self.workflow.sheet.setCurrentText("NoHeader")
        self.assertEqual(self.workflow.source_column.text(), "")
        self.assertEqual(self.workflow.target_column.text(), "")

    def test_finder_load_selects_qa_without_starting_work(self) -> None:
        self.window._bring_window_to_front = Mock()
        self.window.open_qa_workflow_file(str(self.input))
        self.assertEqual(self.window.current_tool_key, "workflow")
        self.assertEqual(self.workflow.input_picker.path(), str(self.input))
        self.workflow.run_in_background.assert_not_called()

    def test_history_tb_defaults_to_term_sheet_and_project_restores_fields(self) -> None:
        history = self.directory / "history.xlsx"
        workbook = Workbook()
        try:
            workbook.active.title = "Raw"
            terms = workbook.create_sheet("术语表")
            terms.append(["source术语", "target术语", "source术语（无mark）", "target术语（无mark）"])
            workbook.save(history)
        finally:
            workbook.close()
        self.workflow.history_picker.set_path(str(history))
        self.workflow.refresh_history_sheets()
        self.assertEqual(self.workflow.history_sheet.currentText(), "术语表")
        self.assertEqual(self.workflow.history_source.text(), "A")
        self.assertEqual(self.workflow.history_target.text(), "B")
        self.workflow.tb_store.save_project(TbProject("Project", str(history), "术语表", "C", "D", 4))
        self.workflow.refresh_tb_projects("Project")
        self.workflow.load_selected_tb_project("Project")
        self.assertEqual(self.workflow.history_source.text(), "C")
        self.assertEqual(self.workflow.history_target.text(), "D")
        self.assertEqual(self.workflow.history_start_row.value(), 4)

    def test_term_check_allows_history_only_and_ignores_unselected_history(self) -> None:
        self.workflow.load_input_file(str(self.input))
        self.window.open_tool("workflow", ["term"])
        self.workflow.mark_book.setChecked(False)
        self.workflow.mark_square.setChecked(False)
        self.workflow.history_picker.set_path(str(self.input))
        self.workflow.refresh_history_sheets()
        self.workflow.run_selected_tasks()
        self.assertEqual(self.workflow.run_in_background.call_args.kwargs["kwargs"]["term_mark_styles"], ())
        self.window.open_tool("workflow", ["tag"])
        self.workflow.run_selected_tasks()
        kwargs = self.workflow.run_in_background.call_args.kwargs["kwargs"]
        self.assertFalse(kwargs["run_term_pair_check"])
        self.assertTrue(kwargs["run_tag_check"])

    def test_french_nbsp_detects_target_when_switching_sheet(self) -> None:
        page = self.window.tool_frames["french_nbsp"]
        page.load_input_file(str(self.input))
        self.assertEqual(page.sheet.currentText(), "Alternate")
        self.assertEqual(page.target_column.text(), "E")
        page.sheet.setCurrentText("Data")
        self.assertEqual(page.target_column.text(), "B")

    def test_xbench_passes_selected_sheet_and_previews_output(self) -> None:
        page = self.window.tool_frames["xbench_report"]
        page.run_in_background = Mock()
        with patch("tools.xbench_report_transformer.qt_page._choose_excel", return_value=str(self.input)):
            page.choose_input()
        self.assertEqual(page.sheet.currentText(), "Alternate")
        page.sheet.setCurrentText("Data")
        page.run_transform()
        self.assertEqual(page.run_in_background.call_args.kwargs["kwargs"]["sheet"], "Data")
        self.assertIn("input", page.preview.text())

    def test_merger_uses_previewed_output_and_header_option(self) -> None:
        page = self.window.tool_frames["excel_merger"]
        page.run_in_background = Mock()
        expected = self.directory / "merged.xlsx"
        with patch("tools.excel_merger.qt_page.build_merge_output_path", return_value=expected):
            page.input_dir.set_path(str(self.directory))
            page.keep_headers.setChecked(True)
            page.run_merge()
        self.assertEqual(page.preview.text(), f"输出文件：{expected}")
        self.assertEqual(page.run_in_background.call_args.kwargs["kwargs"], {
            "folder_path": str(self.directory), "output_path": expected, "keep_all_headers": True,
        })

    def test_apply_revisions_uses_selected_report_without_loading_it_on_gui_thread(self) -> None:
        report = str(self.directory / "workflow_check_input.xlsx")
        with (
            patch("tools.workflow.qt_page.QFileDialog.getOpenFileName", return_value=(report, "")),
            patch("tools.workflow.qt_page.QFileDialog.getSaveFileName", return_value=(str(self.directory / "revised.xlsx"), "")),
        ):
            self.workflow.apply_revisions()
        call = self.workflow.run_in_background.call_args
        self.assertEqual(call.kwargs["args"], (report,))


if __name__ == "__main__":
    unittest.main()
