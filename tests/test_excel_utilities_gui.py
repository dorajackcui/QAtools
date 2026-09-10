from __future__ import annotations

import os
from pathlib import Path
import tempfile
import time
import unittest
from unittest.mock import Mock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
from PySide6.QtCore import QCoreApplication, QEvent
from PySide6.QtWidgets import QApplication

from tools.content_sync.qt_page import MasterToTargetPage, TargetToMasterPage
from tools.excel_utilities_pages import ColumnToolsPage, CompatibilityPage, DeepReplacePage, UntranslatedStatsPage
from tools.excel_file_ops import BatchSummary, OperationResult
from test_momotools_utilities import workbook, values


class UtilitiesGuiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def page(self, factory):
        page = factory()
        self.addCleanup(self.dispose_page, page)
        return page

    @staticmethod
    def dispose_page(page):
        # Closing hides a widget; explicitly dispose its timers on the GUI thread.
        page.close()
        page.deleteLater()
        QCoreApplication.sendPostedEvents(None, QEvent.Type.DeferredDelete)

    def wait(self, page):
        deadline = time.monotonic() + 15
        while page.has_running_tasks() and time.monotonic() < deadline:
            self.app.processEvents()
            time.sleep(0.005)
        self.assertFalse(page.has_running_tasks())
        self.assertTrue(page.run_button.isEnabled())

    def test_all_pages_construct_without_workbook_io_or_com(self):
        with patch("openpyxl.load_workbook", side_effect=AssertionError("eager IO")), \
             patch("tools.excel_com._load_com", side_effect=AssertionError("eager COM")):
            for factory in (TargetToMasterPage, ColumnToolsPage, CompatibilityPage, DeepReplacePage, UntranslatedStatsPage):
                page = self.page(factory)
                self.assertEqual(page.run_button.parentWidget().objectName(), "pageActionBar")

    def test_all_six_output_fields_are_optional_and_input_selection_preserves_choice(self):
        for factory in (MasterToTargetPage, TargetToMasterPage, ColumnToolsPage, CompatibilityPage, DeepReplacePage, UntranslatedStatsPage):
            with self.subTest(page=factory.__name__):
                page = self.page(factory)
                if isinstance(page, MasterToTargetPage):
                    page.master_picker.set_path("master.xlsx")
                    picker = page.target_picker
                    run = page.run_sync
                else:
                    picker = page.input_picker
                    run = page.run_operation
                if isinstance(page, DeepReplacePage):
                    page.source_picker.set_path("sources")
                picker.set_path("inputs")
                self.assertEqual(page.output_picker.path(), "")
                page.run_in_background = Mock()
                run()
                self.assertIsNone(page.run_in_background.call_args.kwargs["kwargs"]["output_dir"])
                page.output_picker.set_path("explicit-copy")
                picker.set_path("other-inputs")
                self.assertEqual(page.output_picker.path(), "explicit-copy")
                run()
                self.assertEqual(page.run_in_background.call_args.kwargs["kwargs"]["output_dir"], "explicit-copy")
                page.output_picker.line_edit.clear()
                run()
                self.assertIsNone(page.run_in_background.call_args.kwargs["kwargs"]["output_dir"])

    def test_column_options_snapshot_and_action_specific_fields(self):
        page = self.page(ColumnToolsPage)
        page.input_picker.set_path("inputs")
        self.assertTrue(page.header_rows.isEnabled())
        self.assertFalse(page.inserted_header.isEnabled())
        page.action.setCurrentIndex(1)
        self.assertFalse(page.header_rows.isEnabled())
        self.assertTrue(page.inserted_header.isEnabled())
        page.run_in_background = Mock()
        page.run_operation()
        kwargs = page.run_in_background.call_args.kwargs["kwargs"]
        page.column.setText("D")
        self.assertEqual(kwargs["column"], "C")
        self.assertEqual(kwargs["action"], "insert")
        self.assertFalse(page.content.isEnabled())

    def test_deep_replace_defaults_to_inplace_and_requires_source(self):
        page = self.page(DeepReplacePage)
        self.assertEqual(page.output_picker.path(), "")
        page.input_picker.set_path("targets")
        page.run_in_background = Mock()
        with patch("tools.excel_utilities_pages.show_error") as error:
            page.run_operation()
        error.assert_called_once()
        page.run_in_background.assert_not_called()

    def test_partial_failures_visible_and_controls_restored(self):
        page = self.page(CompatibilityPage)
        page._busy(True)
        summary = BatchSummary("excel-compatibility", Path("output"), [OperationResult("a.xlsx", "failed", message="locked")])
        with patch("tools.excel_utilities_pages.show_warning") as warning:
            page._finish(summary)
        warning.assert_called_once()
        self.assertTrue(page.content.isEnabled())
        self.assertIn("失败: 1", page.result.toPlainText())

    def test_reverse_and_stats_real_background_flows(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            targets = root / "targets"
            master = root / "master.xlsx"
            workbook(master, [["id", "k", "s", "t"], [1, "k", "你好", "old"]])
            workbook(targets / "a.xlsx", [["k", "s", "t"], ["k", "你好", "None"]])
            reverse = self.page(TargetToMasterPage)
            reverse.master_picker.set_path(str(master))
            reverse.target_picker.set_path(str(targets))
            with patch("tools.content_sync.qt_page.show_error") as error:
                reverse.run_sync()
                self.wait(reverse)
            error.assert_not_called()
            self.assertEqual(reverse.output_picker.path(), "")
            self.assertEqual(values(master, ["D2"]), ["None"])
            self.assertIn("实际更新单元格: 1", reverse.result.toPlainText())
            stats = self.page(UntranslatedStatsPage)
            stats.input_picker.set_path(str(targets))
            with patch("tools.excel_utilities_pages.show_error") as error:
                stats.run_operation()
                self.wait(stats)
            error.assert_not_called()
            self.assertEqual(stats.output_picker.path(), "")
            self.assertTrue((targets / "未翻译统计.xlsx").is_file())
            self.assertIn("统计表:", stats.result.toPlainText())

    def test_file_replacement_real_background_flow(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, target = root / "source", root / "target"
            source.mkdir()
            target.mkdir()
            (source / "a.xlsx").write_bytes(b"new")
            (target / "a.xlsx").write_bytes(b"old")
            page = self.page(DeepReplacePage)
            page.source_picker.set_path(str(source))
            page.input_picker.set_path(str(target))
            with patch("tools.excel_utilities_pages.show_error") as error:
                page.run_operation()
                self.wait(page)
            error.assert_not_called()
            self.assertEqual(page.output_picker.path(), "")
            self.assertEqual((target / "a.xlsx").read_bytes(), b"new")
            self.assertEqual(next((target / ".qatools-backups").rglob("a.xlsx")).read_bytes(), b"old")


if __name__ == "__main__":
    unittest.main()
