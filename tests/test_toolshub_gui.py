from __future__ import annotations

import tkinter as tk
import unittest

from toolshub_gui import ToolshubApp


class ToolshubLayoutTests(unittest.TestCase):
    def make_app(self) -> tuple[tk.Tk, ToolshubApp]:
        try:
            root = tk.Tk()
        except tk.TclError as exc:
            self.skipTest(f"Tk display is unavailable: {exc}")

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

    def test_tools_are_grouped_by_workflow_terms_quality_and_text_repair(self) -> None:
        root, app = self.make_app()

        try:
            grouped_tools = {
                group.title: [tool.title for tool in group.tools]
                for group in app.tool_groups
            }

            self.assertEqual(grouped_tools["常用流程"], ["Workflow 编排"])
            self.assertEqual(
                grouped_tools["术语处理"],
                ["术语对检查", "LLM 术语提取", "术语表命中检查"],
            )
            self.assertEqual(grouped_tools["质量检查"], ["Tag 检查", "Target 中文检查"])
            self.assertEqual(grouped_tools["文本修复"], ["分行拆列", "法语 NBSP 恢复"])
        finally:
            root.destroy()

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

    def test_tool_pages_do_not_show_output_path_controls(self) -> None:
        root, _ = self.make_app()

        try:
            root.update()
            widget_texts = self.collect_widget_texts(root)

            self.assertNotIn("输出 Excel", widget_texts)
            self.assertNotIn("另存为", widget_texts)
        finally:
            root.destroy()

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
