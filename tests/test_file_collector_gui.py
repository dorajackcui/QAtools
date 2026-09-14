import os
from pathlib import Path
import tempfile
import time
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QCoreApplication, QEvent
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication, QMessageBox

from toolshub_gui import ToolshubApp


class FileCollectorGuiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        for module in ("tools.header_aliases", "tools.workflow.gui_options"):
            self.enterContext(patch(f"{module}.default_header_aliases_path", return_value=self.root / "aliases.json"))
        self.enterContext(patch("tools.tb_projects.default_tb_projects_path", return_value=self.root / "projects.json"))
        self.errors = self.enterContext(patch.object(QMessageBox, "critical"))
        self.callback_errors = self.enterContext(patch("sys.excepthook"))
        self.window = ToolshubApp(show_window=False)
        self.page = self.window.tool_frames["file_collector"]
        self.window.select_tool("file_collector")
        self.addCleanup(self.dispose)
        self.source = self.root / "source"
        self.source.mkdir()
        self.output = self.root / "output"
        self.page.source_picker.set_path(str(self.source))
        self.page.output_picker.set_path(str(self.output))

    def dispose(self):
        self.window.close()
        self.window.deleteLater()
        QCoreApplication.sendPostedEvents(None, QEvent.Type.DeferredDelete)

    def file(self, relative, data=b"excel bytes"):
        path = self.source / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(data)
        return path

    def click(self, button):
        self.assertTrue(button.isEnabled())
        button.click()
        deadline = time.monotonic() + 15
        while self.page.has_running_tasks() and time.monotonic() < deadline:
            QTest.qWait(10)
        self.assertFalse(self.page.has_running_tasks())
        self.app.processEvents()
        self.callback_errors.assert_not_called()

    def test_preview_navigation_copy_and_input_invalidation(self):
        self.file("en/a.xlsx", b"en")
        self.file("zh/a.xlsx", b"zh")
        source = self.file("b.xlsm", b"unique")
        page = self.page
        page.names.setPlainText("a\nb\nmissing")
        self.click(page.preview_button)
        self.assertEqual((page.plan.ready_count, page.plan.conflict_count), (1, 2))
        self.assertEqual(page.model.rowCount(), 4)
        self.assertEqual(page.model.index(3, 3).data(), "未找到")
        self.assertFalse(self.output.exists())
        self.window.select_tool("workflow")
        self.window.select_tool("file_collector")
        self.assertTrue(page.run_button.isEnabled())
        self.click(page.run_button)
        self.assertEqual((self.output / "b.xlsm").read_bytes(), source.read_bytes())
        self.assertEqual(page.summary.succeeded_files, 1)
        self.assertIn("已复制", [page.model.index(i, 3).data() for i in range(page.model.rowCount())])
        self.assertFalse(page.run_button.isEnabled(), "plans must not be reused")
        self.assertTrue(page.open_button.isEnabled())
        self.assertTrue(page.preview_button.isEnabled())
        page.names.setPlainText("b")
        self.assertIsNone(page.plan)
        self.assertEqual(page.model.rowCount(), 0)
        self.assertFalse(page.open_button.isEnabled())
        self.errors.assert_not_called()

    def test_import_utf8_list_and_preserve_tree(self):
        self.file("sub/UI 文本.XLSX")
        listing = self.root / "names.txt"
        listing.write_text("UI 文本.xlsx", encoding="utf-8-sig")
        with patch("tools.file_collector.qt_page.QFileDialog.getOpenFileName", return_value=(str(listing), "")):
            self.click(self.page.import_button)
        self.assertEqual(self.page.names.toPlainText(), "UI 文本.xlsx")
        self.page.preserve_tree.setChecked(True)
        self.click(self.page.preview_button)
        self.click(self.page.run_button)
        self.assertEqual((self.output / "sub/UI 文本.XLSX").read_bytes(), b"excel bytes")
        self.errors.assert_not_called()

    def test_changed_source_requires_new_preview_and_can_retry(self):
        source = self.file("a.xlsx")
        self.page.names.setPlainText("a")
        self.click(self.page.preview_button)
        source.write_bytes(b"new content since preview")
        self.click(self.page.run_button)
        self.errors.assert_called_once()
        self.assertIn("重新预览", self.page.status.text())
        self.assertFalse(self.output.exists())
        self.assertIsNone(self.page.plan)
        self.assertFalse(self.page.run_button.isEnabled())
        self.errors.reset_mock()
        self.click(self.page.preview_button)
        self.click(self.page.run_button)
        self.assertEqual((self.output / "a.xlsx").read_bytes(), source.read_bytes())
        self.errors.assert_not_called()

    def test_all_options_invalidate_preview_and_empty_list_recovers(self):
        self.file("a.xlsx")
        self.click(self.page.preview_button)
        self.errors.assert_called_once()
        self.errors.reset_mock()
        self.page.names.setPlainText("a")
        for mutate in (lambda: self.page.preserve_tree.setChecked(True),
                       lambda: self.page.comma_separated.setChecked(True),
                       lambda: self.page.output_picker.set_path(str(self.root / "other"))):
            self.click(self.page.preview_button)
            self.assertTrue(self.page.run_button.isEnabled())
            mutate()
            self.assertIsNone(self.page.plan)
            self.assertFalse(self.page.run_button.isEnabled())
        self.errors.assert_not_called()


if __name__ == "__main__":
    unittest.main()
