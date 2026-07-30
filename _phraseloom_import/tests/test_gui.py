from __future__ import annotations

import io
import runpy
import tkinter as tk
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest.mock import MagicMock, patch

from phraseloom.gui import (
    PhraseLoomGUI,
    TASK_BY_KEY,
    TASKS,
    build_cli_args,
    validate_task_specs,
)


class GuiTaskSpecTests(unittest.TestCase):
    def test_gui_module_can_be_loaded_as_a_direct_script(self) -> None:
        gui_path = Path(__file__).parents[1] / "phraseloom" / "gui.py"

        namespace = runpy.run_path(
            str(gui_path),
            run_name="phraseloom_gui_direct_test",
        )

        self.assertIn("PhraseLoomGUI", namespace)

    def test_gui_uses_one_export_page_with_restore_as_an_action(self) -> None:
        validate_task_specs()
        self.assertEqual(
            [(task.key, task.command) for task in TASKS],
            [("export_strings", "export")],
        )
        self.assertEqual(TASK_BY_KEY["restore_strings"].command, "restore")

    def test_export_builds_clean_cli_args(self) -> None:
        args = build_cli_args(
            TASK_BY_KEY["export_strings"],
            {
                "input": "/tmp/source.xlsx",
                "source_col": "en",
                "target_col": "fr",
                "context_col": "screen",
            },
        )
        self.assertEqual(args[0:2], ["export", "/tmp/source.xlsx"])
        self.assertIn("--source-col", args)
        self.assertIn("--target-col", args)
        self.assertIn("--context-col", args)
        self.assertNotIn("--group-similar", args)
        self.assertNotIn("--tm", args)

    def test_export_can_enable_similar_string_grouping(self) -> None:
        args = build_cli_args(
            TASK_BY_KEY["export_strings"],
            {
                "input": "/tmp/source.xlsx",
                "group_similar": True,
            },
        )
        self.assertIn("--group-similar", args)

    def test_restore_only_needs_the_strings_workbook(self) -> None:
        args = build_cli_args(
            TASK_BY_KEY["restore_strings"],
            {"input": "/tmp/source_strings.xlsx"},
        )
        self.assertEqual(args, ["restore", "/tmp/source_strings.xlsx"])

    def test_restore_button_selects_one_file_and_runs_immediately(self) -> None:
        app = object.__new__(PhraseLoomGUI)
        app.root = object()
        app.output_preview_var = MagicMock()
        app._start_task = MagicMock()

        with (
            patch(
                "phraseloom.gui.filedialog.askopenfilename",
                return_value="/tmp/source_strings.xlsx",
            ),
            patch(
                "phraseloom.gui.default_restored_output_path",
                return_value=Path("/tmp/source_translated.xlsx"),
            ),
        ):
            app._choose_and_restore()

        app.output_preview_var.set.assert_called_once_with(
            "回填输出：source_translated.xlsx"
        )
        app._start_task.assert_called_once_with(
            TASK_BY_KEY["restore_strings"],
            ["restore", "/tmp/source_strings.xlsx"],
        )

    def test_missing_input_is_rejected(self) -> None:
        with self.assertRaisesRegex(ValueError, "请填写"):
            build_cli_args(TASK_BY_KEY["restore_strings"], {})

    def test_cli_dispatches_gui_without_eagerly_starting_tk(self) -> None:
        from phraseloom.cli import _dispatch

        with patch("phraseloom.gui.main", return_value=0) as gui_main:
            self.assertEqual(_dispatch(["gui"]), 0)
        gui_main.assert_called_once_with()

    def test_top_level_help_lists_the_two_step_workflow(self) -> None:
        from phraseloom.cli import _dispatch

        output = io.StringIO()
        with redirect_stdout(output):
            self.assertEqual(_dispatch(["--help"]), 0)
        help_text = output.getvalue()
        self.assertIn("phraseloom export", help_text)
        self.assertIn("phraseloom restore", help_text)
        self.assertNotIn("TM", help_text)
        self.assertNotIn("entity", help_text.lower())


class GuiLayoutTests(unittest.TestCase):
    def make_app(self) -> tuple[tk.Tk, PhraseLoomGUI]:
        try:
            root = tk.Tk()
        except tk.TclError as exc:
            self.skipTest(f"Tk display is unavailable: {exc}")
        return root, PhraseLoomGUI(root)

    def test_gui_uses_qatools_style_sidebar_and_sections(self) -> None:
        root, _app = self.make_app()
        try:
            root.update()
            texts = self._collect_widget_texts(root)
            self.assertIn("常用流程", texts)
            self.assertIn("导出 Strings", texts)
            self.assertIn("回填译文…", texts)
            self.assertIn("输入文件", texts)
            self.assertIn("列与 Tag 设置", texts)
            self.assertIn(
                "启用相似句分组（未聚类在前，聚类内容在后）",
                texts,
            )
            self.assertIn("运行结果", texts)
        finally:
            root.destroy()

    def test_restore_is_a_footer_action_not_a_separate_workspace(self) -> None:
        root, app = self.make_app()
        try:
            root.update()
            self.assertEqual(app.current_title.get(), "导出 Strings")
            texts = self._collect_widget_texts(root)
            self.assertIn("列与 Tag 设置", texts)
            self.assertIn("回填译文…", texts)
            self.assertNotIn("开始回填", texts)
            self.assertNotIn("翻译完成的 Strings 工作簿", texts)
        finally:
            root.destroy()

    def test_initial_window_fits_requested_content(self) -> None:
        root, _app = self.make_app()
        try:
            root.update()
            content = root.winfo_children()[0]
            self.assertGreaterEqual(root.winfo_width(), content.winfo_reqwidth())
            self.assertGreaterEqual(root.winfo_height(), content.winfo_reqheight())
        finally:
            root.destroy()

    def _collect_widget_texts(self, widget: tk.Misc) -> set[str]:
        texts: set[str] = set()
        for child in widget.winfo_children():
            try:
                text = child.cget("text")
            except tk.TclError:
                text = ""
            if text:
                texts.add(str(text))
            texts.update(self._collect_widget_texts(child))
        return texts


if __name__ == "__main__":
    unittest.main()
