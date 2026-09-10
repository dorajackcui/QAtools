from __future__ import annotations

import os
from pathlib import Path
from threading import Event, Thread
import time
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QApplication, QWidget

from tools.qt_operation_logs import LogBuffer, OperationLogDialog
from tools.excel_utilities_pages import UntranslatedStatsPage
from tools.excel_file_ops import BatchSummary
from tools.operation_logs import emit_log, log_result, log_summary
from tools.content_sync.master_to_target import FileResult
from tools.excel_file_ops import OperationResult


class OperationLogsTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def pump_until(self, predicate, seconds=5):
        deadline = time.monotonic() + seconds
        while not predicate() and time.monotonic() < deadline:
            self.app.processEvents()
            time.sleep(0.003)
        self.assertTrue(predicate())

    def test_bounded_buffer_reports_dropped_messages_and_keeps_latest(self):
        buffer = LogBuffer(capacity=3)
        for i in range(20):
            buffer.append(str(i))
        lines = buffer.take()
        self.assertIn("省略 17 条", lines[0])
        self.assertTrue(lines[-1].endswith("19"))
        self.assertEqual(len(lines), 4)
        self.assertEqual(buffer.take(), [])
        buffer.append("x" * 100000)
        self.assertLess(len(buffer.take()[0]), 4100)

    def test_single_line_results_keep_failures_and_recovery_information(self):
        cases = [
            (FileResult("a.xlsx", status="updated", updated_cells=2,
                        postprocess_error="locked\nretry"), "[重存失败] locked\\nretry"),
            (OperationResult("b.xlsx", "failed", message="permission denied"), "permission denied"),
            (OperationResult("c.xlsx", "replaced", details={"backup": "backups/c.xlsx"}), "已备份"),
            (OperationResult("d.xlsx", "skipped", message="重名", details={"sources": ["one/d.xlsx", "two/d.xlsx"]}), "two/d.xlsx"),
            (OperationResult("e.xlsx", "read", details={"identities": 20}), "有效身份 20"),
        ]
        for result, expected in cases:
            with self.subTest(status=result.status):
                messages = []
                log_result(messages.append, result)
                self.assertEqual(len(messages), 1)
                self.assertNotIn("\n", messages[0])
                self.assertIn(expected, messages[0])

    def test_compact_summary_keeps_session_warnings(self):
        messages = []
        log_summary(messages.append, BatchSummary("compatibility", Path("out"),
                    warnings=["Excel 会话未正常结束"]))
        self.assertEqual(messages, ["完成：成功 0，失败 0，跳过 0", "[警告] Excel 会话未正常结束"])

    def test_large_diagnostics_are_bounded_before_formatting(self):
        class NoFullIteration(list):
            def __iter__(self):
                raise AssertionError("must not serialize the entire row list")
        messages = []
        emit_log(messages.append, "未匹配行", {"rows": NoFullIteration(range(100000))})
        self.assertIn("另有 99988 项", messages[0])
        self.assertLess(len(messages[0]), 500)

    def test_large_background_log_volume_does_not_starve_gui_events(self):
        parent = QWidget()
        dialog = OperationLogDialog(parent)
        self.addCleanup(parent.deleteLater)
        dialog.start_run("start")
        produced = Event()
        def produce():
            for i in range(50000):
                dialog.buffer.append(f"file {i}: " + "x" * 500)
            produced.set()
        thread = Thread(target=produce)
        thread.start()
        heartbeats = []
        timer = QTimer(parent)
        timer.setInterval(10)
        timer.timeout.connect(lambda: heartbeats.append(time.monotonic()))
        timer.start()
        self.pump_until(lambda: produced.is_set() and len(heartbeats) >= 5)
        thread.join(timeout=1)
        # A single refresh must not drain/render the entire backlog.
        dialog.text.clear()
        dialog.flush()
        self.assertLessEqual(dialog.text.document().blockCount(), 101)
        self.assertTrue(dialog.buffer.take())
        self.assertEqual(dialog.text.maximumBlockCount(), 10000)
        dialog.timer.stop()
        timer.stop()

    def test_logs_are_live_modeless_and_survive_close_and_reopen_during_task(self):
        page = UntranslatedStatsPage()
        self.addCleanup(page.deleteLater)
        page.input_picker.set_path("inputs")
        release = Event()
        self.addCleanup(release.set)
        def operation(**kwargs):
            kwargs["log_callback"]("正在处理 live.xlsx")
            release.wait(timeout=5)
            kwargs["log_callback"]("live.xlsx 已完成")
            return BatchSummary("stats", Path("output"))
        page.operation = lambda: (operation, {})
        page.run_operation()
        self.assertTrue(page.logs_button.isEnabled())
        self.assertFalse(page.content.isEnabled())
        page.logs_button.click()
        self.pump_until(lambda: "正在处理 live.xlsx" in page.logs.text.toPlainText())
        self.assertTrue(page.has_running_tasks())
        self.assertFalse(page.logs.isModal())
        page.logs.close()
        self.assertTrue(page.has_running_tasks())
        page.logs_button.click()
        self.assertIn("正在处理 live.xlsx", page.logs.text.toPlainText())
        release.set()
        self.pump_until(lambda: not page.has_running_tasks())
        self.pump_until(lambda: "live.xlsx 已完成" in page.logs.text.toPlainText())
        self.pump_until(lambda: not page.logs.timer.isActive())
        self.assertTrue(page.run_button.isEnabled())
        page.logs.start_run("second run")
        self.assertNotIn("live.xlsx", page.logs.text.toPlainText())
        page.logs.finish_run("done")

    def test_fatal_error_is_available_in_log(self):
        page = UntranslatedStatsPage()
        self.addCleanup(page.deleteLater)
        with patch("tools.excel_utilities_pages.show_error"):
            page._fail("workbook is locked")
        page.logs.flush()
        self.assertIn("workbook is locked", page.logs.text.toPlainText())


if __name__ == "__main__":
    unittest.main()
