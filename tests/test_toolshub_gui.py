from __future__ import annotations

import tkinter as tk
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from phraseloom.gui import PhraseLoomApp
from toolshub_gui import TOOL_GROUPS, ToolshubApp, build_argument_parser, main
from tools.french_nbsp_restorer.restore_french_nbsp_gui import FrenchNbspRestorerApp


class ToolshubLayoutTests(unittest.TestCase):
    def test_smoke_test_initializes_tk_without_building_the_app(self) -> None:
        root = Mock()
        with patch("toolshub_gui.tk.Tk", return_value=root):
            self.assertEqual(main(["--smoke-test"]), 0)

        root.withdraw.assert_called_once_with()
        root.update_idletasks.assert_called_once_with()
        root.destroy.assert_called_once_with()

    def make_app(self) -> tuple[tk.Tk, ToolshubApp]:
        try:
            root = tk.Tk()
        except tk.TclError as exc:
            self.skipTest(f"Tk display is unavailable: {exc}")

        root.withdraw()
        root.deiconify = lambda: None
        return root, ToolshubApp(root)

    def test_initial_window_is_not_smaller_than_requested_content(self) -> None:
        root, _ = self.make_app()

        try:
            root.update()

            content = root.winfo_children()[0]
            self.assertGreaterEqual(root.winfo_width(), content.winfo_reqwidth())
            self.assertGreaterEqual(root.winfo_height(), content.winfo_reqheight())
        finally:
            root.destroy()

    def test_tools_are_grouped_by_workflow_quality_text_repair_and_other(self) -> None:
        grouped_tools = {
            group.title: [tool.title for tool in group.tools]
            for group in TOOL_GROUPS
        }

        self.assertEqual(
            grouped_tools["常用流程"],
            ["一键质量检查", "PhraseLoom"],
        )
        self.assertNotIn("术语处理", grouped_tools)
        self.assertEqual(
            grouped_tools["质量检查"],
            [
                "术语检查",
                "Tag 检查",
                "换行数量检查",
                "同源译文一致性",
                "Target 中文检查",
            ],
        )
        self.assertEqual(grouped_tools["文本修复"], ["法语 NBSP 恢复"])
        self.assertEqual(grouped_tools["其他"], ["Xbench QA 转换"])

    def test_selecting_a_tool_updates_heading_and_visible_page(self) -> None:
        root, app = self.make_app()

        try:
            app.select_tool("tag_checker")
            root.update()

            self.assertEqual(app.selected_tool_key.get(), "tag_checker")
            self.assertEqual(app.current_tool_title.get(), "Tag 检查")
            self.assertIn("tag", app.current_tool_description.get().lower())
            self.assertIs(app.current_tool_frame, app.tool_frames["tag_checker"])
        finally:
            root.destroy()

    def test_phraseloom_is_embedded_as_a_tool_page(self) -> None:
        root, app = self.make_app()

        try:
            app.select_tool("phraseloom")
            root.update()

            self.assertIsInstance(app.current_tool_frame, PhraseLoomApp)
            self.assertEqual(app.current_tool_title.get(), "PhraseLoom")
            self.assertIn("Strings", app.current_tool_description.get())
        finally:
            root.destroy()

    def test_tool_pages_do_not_show_output_path_controls(self) -> None:
        root, _ = self.make_app()

        try:
            root.update()
            widget_texts = self.collect_widget_texts(root)

            self.assertNotIn("输出 Excel", widget_texts)
            self.assertNotIn("另存为", widget_texts)
        finally:
            root.destroy()

    def test_qa_workflow_argument_accepts_finder_excel_path(self) -> None:
        args = build_argument_parser().parse_args(
            ["--qa-workflow", "/tmp/QA input.xlsx"]
        )

        self.assertEqual(args.qa_workflow, "/tmp/QA input.xlsx")

    def test_nbsp_restore_argument_accepts_finder_excel_path(self) -> None:
        args = build_argument_parser().parse_args(
            ["--nbsp-restore", "/tmp/French input.xlsx"]
        )

        self.assertEqual(args.nbsp_restore, "/tmp/French input.xlsx")

    def test_nbsp_finder_action_loads_safe_defaults_and_runs_restore(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            workbook_path = Path(tmp_dir) / "French input.xlsx"
            workbook_path.touch()
            restorer = FrenchNbspRestorerApp.__new__(FrenchNbspRestorerApp)
            restorer.load_input_file = Mock()
            restorer.run_restore = Mock()
            app = ToolshubApp.__new__(ToolshubApp)
            app.root = SimpleNamespace(
                after_idle=lambda callback: callback(),
            )
            app.tool_frames = {"french_nbsp": restorer}
            app.select_tool = Mock()
            app._bring_window_to_front = Mock()

            app.open_french_nbsp_restore_file(str(workbook_path))

            app.select_tool.assert_called_once_with("french_nbsp")
            restorer.load_input_file.assert_called_once_with(
                str(workbook_path.absolute()),
                reset_options=True,
            )
            app._bring_window_to_front.assert_called_once()
            restorer.run_restore.assert_called_once()

    def collect_widget_texts(self, widget: tk.Misc) -> set[str]:
        texts: set[str] = set()
        for child in widget.winfo_children():
            try:
                text = child.cget("text")
            except tk.TclError:
                text = ""
            if text:
                texts.add(str(text))
            texts.update(self.collect_widget_texts(child))
        return texts


if __name__ == "__main__":
    unittest.main()
