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
from openpyxl import Workbook, load_workbook

from tools.content_sync.master_to_target import FileResult, SyncSummary
from tools.content_sync.qt_page import ContentSyncPage, MasterToTargetPage


class ContentSyncGuiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.page = MasterToTargetPage()
        self.addCleanup(self.dispose_page, self.page)

    @staticmethod
    def dispose_page(page):
        # Destroy Qt timers on the GUI thread before a later worker can trigger GC.
        page.close()
        page.deleteLater()
        QCoreApplication.sendPostedEvents(None, QEvent.Type.DeferredDelete)

    def test_page_construction_does_not_load_workbook_or_com(self):
        with patch("openpyxl.load_workbook", side_effect=AssertionError("eager IO")), \
             patch("tools.content_sync.qt_page._master_sheets", side_effect=AssertionError("eager probe")):
            page = ContentSyncPage()
            self.dispose_page(page)
        self.assertEqual(self.page.run_button.parentWidget().objectName(), "pageActionBar")
        self.assertFalse(self.page.fill_blank_only.isChecked())
        self.assertFalse(self.page.allow_blank_write.isChecked())

    def test_direction_tabs_preserve_independent_inputs_results_and_logs(self):
        page = ContentSyncPage()
        self.addCleanup(self.dispose_page, page)
        self.assertEqual([page.tabs.tabText(i) for i in range(page.tabs.count())],
                         ["Master → 小表", "小表 → Master"])
        for index in (0, 1):
            page.tabs.setCurrentIndex(index)
            direction = page.tabs.currentWidget()
            direction.master_picker.set_path(f"master-{index}.xlsx")
            direction.result.setPlainText(f"result-{index}")
            direction.logs.start_run(f"log-{index}")
            direction.logs.finish_run("done")
        for index in (0, 1):
            page.tabs.setCurrentIndex(index)
            direction = page.tabs.currentWidget()
            self.assertEqual(direction.master_picker.path(), f"master-{index}.xlsx")
            self.assertEqual(direction.result.toPlainText(), f"result-{index}")
            direction.logs.open_logs()
            self.app.processEvents()
            self.assertIn(f"log-{index}", direction.logs.text.toPlainText())
            self.assertNotIn(f"log-{1 - index}", direction.logs.text.toPlainText())
            direction.logs.close()

    def test_run_captures_widget_values_before_dispatch(self):
        self.page.master_picker.set_path("master.xlsx")
        self.page.target_picker.set_path("targets")
        self.page.column_count.setValue(2)
        self.page.fill_blank_only.setChecked(True)
        self.page.run_in_background = Mock()
        self.page.run_sync()
        kwargs = self.page.run_in_background.call_args.kwargs["kwargs"]
        self.page.column_count.setValue(3)
        self.page.target_sheet.setText("changed afterwards")
        self.assertEqual(kwargs["column_count"], 2)
        self.assertIsNone(kwargs["target_sheet"])
        self.assertTrue(kwargs["fill_blank_only"])
        self.assertFalse(self.page.run_button.isEnabled())
        self.assertFalse(self.page.content.isEnabled())

    def test_missing_input_and_overlapping_columns_do_not_start_worker(self):
        self.page.run_in_background = Mock()
        with patch("tools.content_sync.qt_page.show_error") as error:
            self.page.run_sync()
            error.assert_called_once()
        self.page.master_picker.set_path("master.xlsx")
        self.page.target_picker.set_path("targets")
        self.page.master_columns[0].setText("D")
        with patch("tools.content_sync.qt_page.show_error") as error:
            self.page.run_sync()
            error.assert_called_once()
        self.page.run_in_background.assert_not_called()

    def test_partial_failure_is_visible_and_controls_are_restored(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            summary = SyncSummary(root, 1, [],
                                  [FileResult("broken.xlsx", error="locked")])
            self.page._set_busy(True)
            with patch("tools.content_sync.qt_page.show_warning") as warning:
                self.page._finish(summary)
                warning.assert_called_once()
            self.assertIn("失败: 1", self.page.result.toPlainText())
            self.assertTrue(self.page.run_button.isEnabled())

    def test_real_background_run_generates_output_and_updates_gui(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            targets = root / "targets"
            targets.mkdir()
            master = root / "master.xlsx"
            for path, rows in (
                (master, [["id", "key", "match", "value"], [1, "k", "s", "  nan  "]]),
                (targets / "a.xlsx", [["key", "match", "value"], ["k", "s", "old"]]),
            ):
                wb = Workbook()
                for row in rows:
                    wb.active.append(row)
                wb.save(path)
                wb.close()
            self.page.master_picker.set_path(str(master))
            self.page.target_picker.set_path(str(targets))
            with patch("tools.content_sync.qt_page.show_error") as error:
                self.page.run_sync()
                deadline = time.monotonic() + 10
                while self.page.has_running_tasks() and time.monotonic() < deadline:
                    self.app.processEvents()
                    time.sleep(0.005)
                self.assertFalse(self.page.has_running_tasks(), "worker did not finish")
                error.assert_not_called()
            self.assertTrue(self.page.run_button.isEnabled())
            self.assertIn("实际更新单元格: 1", self.page.result.toPlainText())
            self.assertEqual(self.page.output_picker.path(), "")
            wb = load_workbook(targets / "a.xlsx")
            try:
                self.assertEqual(wb.active["C2"].value, "  nan  ")
            finally:
                wb.close()


if __name__ == "__main__":
    unittest.main()
