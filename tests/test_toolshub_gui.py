from __future__ import annotations

import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtGui import QCloseEvent  # noqa: E402
from PySide6.QtWidgets import QApplication, QLineEdit, QStackedWidget  # noqa: E402

from toolshub_gui import (  # noqa: E402
    TOOL_GROUPS,
    ToolshubApp,
    build_argument_parser,
    calculate_initial_window_size,
    main,
)
from tools.qt_pages import FrenchNbspPage, PhraseLoomPage, WorkflowPage  # noqa: E402


class ToolshubLayoutTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.qt_app = QApplication.instance() or QApplication([])

    def make_app(self) -> ToolshubApp:
        return ToolshubApp(show_window=False)

    def test_initial_window_size_adds_breathing_room(self) -> None:
        self.assertEqual(
            calculate_initial_window_size(
                requested_width=1238,
                requested_height=725,
                screen_width=1920,
                screen_height=1080,
            ),
            (1334, 840),
        )

    def test_initial_window_size_stays_within_available_screen(self) -> None:
        self.assertEqual(
            calculate_initial_window_size(
                requested_width=1500,
                requested_height=900,
                screen_width=1366,
                screen_height=768,
            ),
            (1286, 688),
        )

    def test_smoke_test_builds_complete_app_without_showing_window(self) -> None:
        receiver = Mock()
        receiver.start.return_value = False
        with patch("toolshub_gui.WorkflowFileReceiver", return_value=receiver):
            self.assertEqual(main(["--smoke-test"]), 0)
        receiver.close.assert_called()

    def test_navigation_only_exposes_top_level_workflows_and_utilities(self) -> None:
        grouped_tools = {
            group.title: [tool.title for tool in group.tools]
            for group in TOOL_GROUPS
        }
        self.assertEqual(grouped_tools["常用流程"], ["一键质量检查", "PhraseLoom"])
        self.assertEqual(grouped_tools["文本修复"], ["法语 NBSP 恢复"])
        self.assertEqual(
            grouped_tools["其他"],
            ["Batch 拆分", "合并表格", "Xbench QA 转换"],
        )

    def test_all_qt_pages_are_created_once_and_kept_in_stack(self) -> None:
        window = self.make_app()
        try:
            self.assertIsInstance(window.page_stack, QStackedWidget)
            self.assertEqual(
                set(window.tool_frames),
                {
                    "workflow",
                    "phraseloom",
                    "french_nbsp",
                    "excel_batcher",
                    "excel_merger",
                    "xbench_report",
                },
            )
            self.assertEqual(window.page_stack.count(), 6)
            self.assertIsInstance(window.tool_frames["workflow"], WorkflowPage)
            self.assertIsInstance(window.tool_frames["phraseloom"], PhraseLoomPage)
        finally:
            window.close()

    def test_selecting_a_tool_reuses_persistent_page(self) -> None:
        window = self.make_app()
        try:
            page = window.tool_frames["french_nbsp"]
            window.select_tool("french_nbsp")
            window.select_tool("workflow")
            window.select_tool("french_nbsp")
            self.assertEqual(window.current_tool_key, "french_nbsp")
            self.assertEqual(window.title_label.text(), "法语 NBSP 恢复")
            self.assertIs(window.current_tool_frame, page)
            self.assertIs(window.page_stack.currentWidget(), page)
            self.assertTrue(window.nav_buttons["french_nbsp"].isChecked())
        finally:
            window.close()

    def test_workflow_file_path_is_read_only_and_button_driven(self) -> None:
        window = self.make_app()
        try:
            workflow = window.tool_frames["workflow"]
            self.assertTrue(workflow.input_picker.line_edit.isReadOnly())
            self.assertIsInstance(workflow.input_picker.line_edit, QLineEdit)
            workflow.input_picker.set_path(r"C:\very\long\folder\sample.xlsx")
            self.assertEqual(
                workflow.input_picker.path(),
                r"C:\very\long\folder\sample.xlsx",
            )
        finally:
            window.close()

    def test_workflow_optional_tag_config_can_be_cleared(self) -> None:
        window = self.make_app()
        try:
            workflow = window.tool_frames["workflow"]
            workflow.angle_config.set_path(r"C:\configs\tags.json")
            workflow.angle_config.clear()
            self.assertEqual(workflow.angle_config.path(), "")
        finally:
            window.close()

    def test_workflow_runs_excel_processing_through_background_worker(self) -> None:
        window = self.make_app()
        try:
            workflow = window.tool_frames["workflow"]
            workflow.input_picker.set_path(r"C:\input\sample.xlsx")
            workflow.sheet.addItem("Sheet1")
            workflow.source_column.setText("A")
            workflow.target_column.setText("B")
            workflow.run_in_background = Mock()

            workflow.run_selected_tasks()

            workflow.run_in_background.assert_called_once()
            call_kwargs = workflow.run_in_background.call_args.kwargs
            self.assertEqual(call_kwargs["kwargs"]["input_file"], r"C:\input\sample.xlsx")
            self.assertEqual(call_kwargs["kwargs"]["sheet"], "Sheet1")
            self.assertTrue(call_kwargs["kwargs"]["run_term_pair_check"])
            self.assertTrue(call_kwargs["kwargs"]["run_target_text_check"])
            self.assertEqual(call_kwargs["kwargs"]["term_mark_styles"], ("【】", "[]"))
            self.assertEqual(
                call_kwargs["kwargs"]["tag_token_types"],
                ("angle", "square_color", "brace", "newline"),
            )
        finally:
            window.close()

    def test_batch_worker_uses_values_captured_on_gui_thread(self) -> None:
        window = self.make_app()
        try:
            batcher = window.tool_frames["excel_batcher"]
            batcher.split_input.set_path(r"C:\input\sample.xlsx")
            batcher.split_sheet.addItem("Sheet1")
            batcher.batch_size.setValue(250)
            batcher.header_rows.setValue(2)
            batcher.split_output.set_path(r"C:\output\batches")
            batcher.run_in_background = Mock()

            batcher.run_split()
            task = batcher.run_in_background.call_args.args[0]

            batcher.split_sheet.setCurrentText("")
            batcher.batch_size.setValue(999)
            batcher.header_rows.setValue(0)
            batcher.split_output.clear()
            with patch("tools.qt_pages.split_workbook", return_value=Mock()) as split:
                task()

            split.assert_called_once_with(
                input_file=r"C:\input\sample.xlsx",
                sheet="Sheet1",
                batch_size=250,
                header_rows=2,
                output_dir=r"C:\output\batches",
            )
        finally:
            window.close()

    def test_window_refuses_to_close_while_background_task_is_running(self) -> None:
        window = self.make_app()
        workflow = window.tool_frames["workflow"]
        worker = Mock()
        workflow._workers.add(worker)
        event = QCloseEvent()
        try:
            with patch("toolshub_gui.show_warning") as warning:
                window.closeEvent(event)

            self.assertFalse(event.isAccepted())
            warning.assert_called_once()
            self.assertIn("一键质量检查", warning.call_args.args[2])
        finally:
            workflow._workers.clear()
            window.close()

    def test_qa_workflow_argument_accepts_finder_excel_path(self) -> None:
        args = build_argument_parser().parse_args(["--qa-workflow", "/tmp/QA input.xlsx"])
        self.assertEqual(args.qa_workflow, "/tmp/QA input.xlsx")

    def test_nbsp_restore_argument_accepts_finder_excel_path(self) -> None:
        args = build_argument_parser().parse_args(["--nbsp-restore", "/tmp/French input.xlsx"])
        self.assertEqual(args.nbsp_restore, "/tmp/French input.xlsx")

    def test_nbsp_finder_action_loads_safe_defaults_and_schedules_restore(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            workbook_path = Path(tmp_dir) / "French input.xlsx"
            workbook_path.touch()
            window = self.make_app()
            try:
                restorer = window.tool_frames["french_nbsp"]
                self.assertIsInstance(restorer, FrenchNbspPage)
                restorer.load_input_file = Mock()
                restorer.run_restore = Mock()
                window._bring_window_to_front = Mock()
                with patch("toolshub_gui.QTimer.singleShot", side_effect=lambda _delay, callback: callback()):
                    window.open_french_nbsp_restore_file(str(workbook_path))
                restorer.load_input_file.assert_called_once_with(
                    str(workbook_path.absolute()),
                    reset_options=True,
                )
                restorer.run_restore.assert_called_once_with()
                window._bring_window_to_front.assert_called_once_with()
            finally:
                window.close()


if __name__ == "__main__":
    unittest.main()
