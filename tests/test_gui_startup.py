from __future__ import annotations

import os
from pathlib import Path
import subprocess
import sys
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QCoreApplication, QEvent
from PySide6.QtWidgets import QApplication, QWidget

from toolshub_gui import ToolshubApp


class GuiStartupTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def tearDown(self):
        QCoreApplication.sendPostedEvents(None, QEvent.Type.DeferredDelete)

    def run_fresh(self, code):
        result = subprocess.run(
            [sys.executable, "-c", code],
            cwd=Path(__file__).resolve().parents[1],
            env=dict(os.environ, QT_QPA_PLATFORM="offscreen"),
            capture_output=True, text=True, timeout=30,
        )
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_entry_import_does_not_import_workbooks_or_any_tool_page(self):
        self.run_fresh(
            "import sys\n"
            "import toolshub_gui\n"
            "assert 'openpyxl' not in sys.modules\n"
            "assert 'phraseloom.strings_workflow' not in sys.modules\n"
            "assert 'tools.workflow.workflow_runner' not in sys.modules\n"
            "assert not any(n.endswith('.qt_page') for n in sys.modules)\n"
        )

    def test_shell_is_visible_before_first_page_factory_runs(self):
        observed = []

        def factory(**kwargs):
            windows = [w for w in self.app.topLevelWidgets() if isinstance(w, ToolshubApp) and w.isVisible()]
            self.assertEqual(len(windows), 1)
            shell = windows[0]
            self.assertTrue(shell.loading_page.isVisible())
            self.assertIn("正在", shell.loading_page.text())
            self.assertEqual(len(shell.tool_frames), 0)
            observed.append(shell.loading_page.text())
            return QWidget()

        with patch.dict("toolshub_gui.PAGE_FACTORIES", {"workflow": factory}):
            window = ToolshubApp(show_window=True)
        try:
            self.assertEqual(len(observed), 1)
            self.assertEqual(window.current_tool_key, "workflow")
            self.assertEqual(window.page_stack.count(), 1)
            self.assertFalse(window.loading_page.isVisible())
        finally:
            window.close()
            window.deleteLater()

    def test_requested_directory_page_skips_qa_and_other_business_imports(self):
        self.run_fresh(
            "import sys\n"
            "from toolshub_gui import ToolshubApp, create_qt_application\n"
            "app, _ = create_qt_application([])\n"
            "window = ToolshubApp(show_window=False, initial_tool='compatibility')\n"
            "assert set(window.tool_frames) == {'compatibility'}\n"
            "assert 'tools.workflow.qt_page' not in sys.modules\n"
            "assert 'phraseloom.qt_page' not in sys.modules\n"
            "assert 'openpyxl' not in sys.modules\n"
            "window.close()\n"
        )

    def test_navigation_caches_pages_and_close_does_not_create_unused_pages(self):
        calls = []

        def factory(name):
            def create(**kwargs):
                calls.append(name)
                return QWidget()
            return create

        with patch.dict("toolshub_gui.PAGE_FACTORIES", {
            "workflow": factory("workflow"), "phraseloom": factory("phraseloom"),
        }):
            window = ToolshubApp(show_window=False)
            try:
                original = window.tool_frames["workflow"]
                window.select_tool("phraseloom")
                window.select_tool("workflow")
                self.assertIs(window.current_tool_frame, original)
                self.assertEqual(calls, ["workflow", "phraseloom"])
                self.assertEqual(window.page_stack.count(), 2)
                window.close()
                self.assertEqual(calls, ["workflow", "phraseloom"])
            finally:
                window.close()
                window.deleteLater()


if __name__ == "__main__":
    unittest.main()
