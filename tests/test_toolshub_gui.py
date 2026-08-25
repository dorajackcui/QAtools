from __future__ import annotations

import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from phraseloom.gui import PhraseLoomApp
from toolshub_gui import (
    TOOL_GROUPS,
    ToolshubApp,
    build_argument_parser,
    calculate_initial_window_size,
    main,
)
from tools.gui_common import FILE_PATH_DISPLAY_WIDTH
from tools.french_nbsp_restorer.restore_french_nbsp_gui import FrenchNbspRestorerApp
from tools.gui_common import APP_MAIN_BACKGROUND


class ToolshubLayoutTests(unittest.TestCase):
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
        root = Mock()
        with (
            patch("toolshub_gui.create_application_root", return_value=root),
            patch("toolshub_gui.ToolshubApp") as app_class,
        ):
            self.assertEqual(main(["--smoke-test"]), 0)

        root.withdraw.assert_called_once_with()
        app_class.assert_called_once_with(root, show_window=False)
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

    def test_minimum_width_keeps_the_requested_layout_visible(self) -> None:
        root, _ = self.make_app()

        try:
            root.update_idletasks()
            content = root.winfo_children()[0]
            minimum_width, _ = root.minsize()

            self.assertGreaterEqual(minimum_width, content.winfo_reqwidth())
        finally:
            root.destroy()

    def test_sun_valley_dark_theme_and_toggle_navigation_are_active(self) -> None:
        root, _ = self.make_app()

        try:
            style = ttk.Style(root)

            self.assertEqual(style.theme_use(), "sun-valley-dark")
            self.assertEqual(
                style.layout("Toolshub.Nav.TRadiobutton"),
                style.layout("Toggle.TButton"),
            )
            self.assertIn("Checkbutton.indicator", repr(style.layout("TCheckbutton")))
            body_font = tkfont.nametofont("SunValleyBodyFont", root=root)
            default_font = tkfont.nametofont("TkDefaultFont", root=root)
            self.assertEqual(body_font.actual("family"), default_font.actual("family"))
            self.assertEqual(body_font.actual("size"), 10)
        finally:
            root.destroy()

    def test_workflow_file_picker_is_button_driven_and_non_editable(self) -> None:
        root, app = self.make_app()

        try:
            workflow_frame = app.tool_frames["workflow"]
            variable_name = str(workflow_frame.input_file_var)
            matching_entries = [
                widget
                for widget in self.collect_widgets(workflow_frame)
                if isinstance(widget, ttk.Entry)
                and str(widget.cget("textvariable")) == variable_name
            ]

            self.assertEqual(matching_entries, [])

            path_display = workflow_frame.input_file_display
            self.assertIsInstance(path_display, ttk.Label)
            self.assertEqual(int(path_display.cget("width")), FILE_PATH_DISPLAY_WIDTH)
            self.assertEqual(str(path_display.grid_info()["sticky"]), "w")
            self.assertEqual(
                path_display._file_path_display_variable.get(),
                "未选择文件",
            )

            workflow_frame.input_file_var.set(r"C:\very\long\folder\sample.xlsx")
            self.assertEqual(
                path_display._file_path_display_variable.get(),
                "sample.xlsx",
            )
        finally:
            root.destroy()

    def test_workflow_optional_tag_config_can_be_cleared(self) -> None:
        root, app = self.make_app()

        try:
            workflow = app.tool_frames["workflow"]
            workflow.tag_angle_config_file_var.set(r"C:\configs\tags.json")
            self.assertEqual(
                workflow.tag_angle_config_entry._file_path_display_variable.get(),
                "tags.json",
            )

            workflow.tag_angle_config_clear_button.invoke()

            self.assertEqual(workflow.tag_angle_config_file_var.get(), "")
            self.assertEqual(
                workflow.tag_angle_config_entry._file_path_display_variable.get(),
                "未选择文件",
            )
        finally:
            root.destroy()

    def test_navigation_only_exposes_top_level_workflows_and_utilities(self) -> None:
        grouped_tools = {
            group.title: [tool.title for tool in group.tools]
            for group in TOOL_GROUPS
        }

        self.assertEqual(
            grouped_tools["常用流程"],
            ["一键质量检查", "PhraseLoom"],
        )
        self.assertNotIn("术语处理", grouped_tools)
        self.assertNotIn("质量检查", grouped_tools)
        self.assertEqual(grouped_tools["文本修复"], ["法语 NBSP 恢复"])
        self.assertEqual(
            grouped_tools["其他"],
            ["Batch 拆分", "合并表格", "Xbench QA 转换"],
        )

    def test_all_tool_pages_are_created_at_startup(self) -> None:
        root, app = self.make_app()

        try:
            self.assertEqual(
                set(app.tool_frames),
                {
                    "workflow",
                    "phraseloom",
                    "french_nbsp",
                    "excel_batcher",
                    "excel_merger",
                    "xbench_report",
                },
            )
            for frame in app.tool_frames.values():
                self.assertEqual(frame.winfo_manager(), "grid")
        finally:
            root.destroy()

    def test_selecting_a_tool_updates_heading_and_visible_page(self) -> None:
        root, app = self.make_app()

        try:
            workflow_frame = app.tool_frames["workflow"]
            app.select_tool("french_nbsp")
            root.update()

            self.assertEqual(app.selected_tool_key.get(), "french_nbsp")
            self.assertEqual(app.current_tool_title.get(), "法语 NBSP 恢复")
            self.assertIs(app.current_tool_frame, app.tool_frames["french_nbsp"])
            self.assertEqual(workflow_frame.winfo_manager(), "grid")
            self.assertEqual(app.current_tool_frame.winfo_manager(), "grid")
        finally:
            root.destroy()

    def test_selecting_an_existing_tool_page_reuses_its_frame(self) -> None:
        root, app = self.make_app()

        try:
            app.select_tool("french_nbsp")
            nbsp_frame = app.current_tool_frame
            app.select_tool("workflow")
            app.select_tool("french_nbsp")

            self.assertIs(app.current_tool_frame, nbsp_frame)
            self.assertEqual(nbsp_frame.winfo_manager(), "grid")
            self.assertEqual(app.tool_frames["workflow"].winfo_manager(), "grid")
        finally:
            root.destroy()

    def test_phraseloom_is_embedded_as_a_tool_page(self) -> None:
        root, app = self.make_app()

        try:
            app.select_tool("phraseloom")
            root.update()

            self.assertIsInstance(app.current_tool_frame, PhraseLoomApp)
            self.assertEqual(app.current_tool_title.get(), "PhraseLoom")
        finally:
            root.destroy()

    def test_redundant_shell_copy_and_empty_output_hint_are_hidden(self) -> None:
        root, app = self.make_app()

        try:
            root.update()
            widget_texts = self.collect_widget_texts(root)
            for redundant_text in (
                "本地化 QA 工作台",
                "集中执行全部质量检查项目，并统一写入输出 Excel。",
                "导出干净、去重的 Strings 工作簿，并在翻译后回填原始 Excel。",
                "输出文件：选择输入 Excel 后自动生成",
            ):
                self.assertNotIn(redundant_text, widget_texts)

            workflow = app.tool_frames["workflow"]
            self.assertEqual(workflow.output_preview_label.winfo_manager(), "")
            workflow.output_preview_var.set("输出文件：result.xlsx")
            self.assertEqual(workflow.output_preview_label.winfo_manager(), "grid")
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

    def test_scroll_canvases_use_the_application_background(self) -> None:
        root, app = self.make_app()

        try:
            self.assertEqual(
                app.tool_frames["workflow"].scroll_canvas.cget("background"),
                APP_MAIN_BACKGROUND,
            )
            self.assertEqual(
                app.tool_frames["phraseloom"].scroll_canvas.cget("background"),
                APP_MAIN_BACKGROUND,
            )
        finally:
            root.destroy()

    def test_workflow_short_content_stays_at_the_top_of_a_tall_viewport(self) -> None:
        try:
            root = tk.Tk()
        except tk.TclError as exc:
            self.skipTest(f"Tk display is unavailable: {exc}")

        try:
            app = ToolshubApp(root, show_window=False)
            root.geometry("1334x900")
            root.deiconify()
            root.update()

            workflow = app.tool_frames["workflow"]
            scroll_region = tuple(
                float(value)
                for value in workflow.scroll_canvas.cget("scrollregion").split()
            )
            region_height = scroll_region[3] - scroll_region[1]
            self.assertGreaterEqual(
                region_height,
                workflow.scroll_canvas.winfo_height(),
            )
            self.assertEqual(workflow.scroll_canvas.canvasy(0), 0.0)
            self.assertEqual(
                workflow.scroll_content.winfo_rooty(),
                workflow.scroll_canvas.winfo_rooty(),
            )
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

    def collect_widgets(self, widget: tk.Misc) -> list[tk.Misc]:
        widgets: list[tk.Misc] = []
        for child in widget.winfo_children():
            widgets.append(child)
            widgets.extend(self.collect_widgets(child))
        return widgets


if __name__ == "__main__":
    unittest.main()
