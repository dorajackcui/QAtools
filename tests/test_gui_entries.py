from __future__ import annotations

import ast
import importlib
import os
from pathlib import Path
import runpy
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import Mock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QCoreApplication, QEvent
from PySide6.QtWidgets import QApplication
from toolshub_gui import QA_CHECK_WIDGETS, ToolshubApp, build_argument_parser, main
from tools.qt_navigation import ToolNavigationServer, send_tool_selection


ENTRIES = {
    "phraseloom/gui.py": ["--tool", "phraseloom"],
    "tools/workflow/workflow_gui.py": ["--tool", "workflow"],
    "tools/term_pair_checker/extract_terms_gui.py": ["--tool", "workflow", "--check", "term"],
    "extract_terms_gui.py": ["--tool", "workflow", "--check", "term"],
    "tools/tag_placeholder_checker/check_tags_and_placeholders_gui.py": ["--tool", "workflow", "--check", "tag"],
    "tools/line_break_checker/check_line_breaks_gui.py": ["--tool", "workflow", "--check", "line-break"],
    "tools/source_consistency_checker/check_source_consistency_gui.py": ["--tool", "workflow", "--check", "consistency"],
    "tools/chinese_target_checker/check_chinese_target_gui.py": ["--tool", "workflow", "--check", "chinese"],
    "tools/excel_batcher/excel_batcher_gui.py": ["--tool", "excel_batcher"],
    "tools/excel_merger/merge_active_sheets_gui.py": ["--tool", "excel_merger"],
    "tools/french_nbsp_restorer/restore_french_nbsp_gui.py": ["--tool", "french_nbsp"],
    "tools/xbench_report_transformer/transform_xbench_report_gui.py": ["--tool", "xbench_report"],
}


class GuiEntryTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def test_all_legacy_modules_and_scripts_route_to_shared_qt_pages(self) -> None:
        root = Path(__file__).resolve().parents[1]
        for relative, expected in ENTRIES.items():
            with self.subTest(entry=relative), patch("toolshub_gui.main", return_value=0) as launch:
                module = importlib.import_module(relative[:-3].replace("/", "."))
                self.assertEqual(module.main(), 0)
                launch.assert_called_once_with(expected)
                launch.reset_mock()
                with patch("sys.argv", [str(root / relative)]):
                    self.run_script(root / relative)
                launch.assert_called_once_with(expected)

    def run_script(self, path: Path) -> None:
        with self.assertRaises(SystemExit) as result:
            runpy.run_path(str(path), run_name="__main__")
        self.assertEqual(result.exception.code, 0)

    def test_initial_page_and_single_check_presets_use_real_qt_widgets(self) -> None:
        with tempfile.TemporaryDirectory() as directory, patch(
            "tools.workflow.gui_options.default_header_aliases_path", return_value=Path(directory) / "aliases.json"
        ), patch("tools.tb_projects.default_tb_projects_path", return_value=Path(directory) / "projects.json"):
            window = ToolshubApp(show_window=False)
            try:
                for expected in ENTRIES.values():
                    args = build_argument_parser().parse_args(expected)
                    window.open_tool(args.tool, args.check)
                    self.assertEqual(window.current_tool_key, args.tool)
                    if args.check:
                        page = window.tool_frames["workflow"]
                        selected = [name for name, widget in QA_CHECK_WIDGETS.items() if getattr(page, widget).isChecked()]
                        self.assertEqual(selected, args.check)
            finally:
                window.close()

    def test_main_applies_requested_page_and_preset(self) -> None:
        window = Mock()
        with patch("toolshub_gui.ToolshubApp", return_value=window), patch("toolshub_gui.WorkflowFileReceiver"):
            self.assertEqual(main(["--tool", "workflow", "--check", "tag", "--smoke-test"]), 0)
        window.open_tool.assert_called_once_with("workflow", ["tag"])

    def test_existing_instance_receives_page_request_instead_of_opening_another_window(self) -> None:
        with (
            patch("toolshub_gui._acquire_gui_instance_lock", return_value=None),
            patch("toolshub_gui.send_tool_selection", return_value=True) as forward,
            patch("toolshub_gui.ToolshubApp") as window,
        ):
            self.assertEqual(main(["--tool", "phraseloom"]), 0)
        forward.assert_called_once_with("phraseloom", None)
        window.assert_not_called()

    def test_navigation_server_delivers_page_and_check_selection(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            name = str(Path(directory) / "navigation")
            server = ToolNavigationServer(name=name)
            received = []
            server.requested.connect(lambda tool, checks: received.append((tool, checks)))
            try:
                self.assertTrue(server.start(), server.server.errorString())
                with patch("tools.qt_navigation.navigation_server_name", return_value=name):
                    self.assertTrue(send_tool_selection("workflow", ["tag"]))
                server.server.waitForNewConnection(1000)
                self.app.processEvents()
                self.assertEqual(received, [("workflow", ["tag"])])
            finally:
                server.close()
                server.deleteLater()
                QCoreApplication.sendPostedEvents(None, QEvent.Type.DeferredDelete)

    def test_navigation_rejects_invalid_and_oversized_messages(self) -> None:
        server = ToolNavigationServer()
        self.addCleanup(server.deleteLater)
        received = []
        server.requested.connect(lambda *args: received.append(args))
        for payload in (b"[]\n", b"not json\n", b'{"tool": "workflow", "checks": [123]}\n'):
            socket = Mock()
            socket.bytesAvailable.return_value = len(payload)
            socket.canReadLine.return_value = True
            socket.readLine.return_value = payload
            server._read(socket)
            socket.disconnectFromServer.assert_called_once()
        self.assertEqual(received, [])
        socket = Mock()
        socket.bytesAvailable.return_value = 5000
        server._read(socket)
        socket.disconnectFromServer.assert_called_once()
        socket.readLine.assert_not_called()

    def test_invalid_check_without_workflow_is_rejected(self) -> None:
        with self.assertRaises(SystemExit) as result:
            main(["--tool", "phraseloom", "--check", "tag"])
        self.assertEqual(result.exception.code, 2)

    def test_fresh_process_can_build_all_pages_without_tk_or_tk_theme(self) -> None:
        root = Path(__file__).resolve().parents[1]
        result = subprocess.run(
            [sys.executable, "-c", "import sys; sys.modules['tkinter'] = None; sys.modules['_tkinter'] = None; sys.modules['sv_ttk'] = None; from phraseloom.gui import main; raise SystemExit(main(['--smoke-test']))"],
            cwd=root, capture_output=True, text=True, timeout=30,
        )
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_application_source_does_not_import_legacy_gui_framework(self) -> None:
        root = Path(__file__).resolve().parents[1]
        paths = list(root.glob("*.py"))
        for package in ("tools", "phraseloom", "qatools"):
            paths.extend((root / package).rglob("*.py"))
        forbidden = {"tkinter", "_tkinter", "sv_ttk"}
        for path in paths:
            tree = ast.parse(path.read_text(encoding="utf-8"))
            for node in ast.walk(tree):
                modules = []
                if isinstance(node, ast.Import):
                    modules = [alias.name for alias in node.names]
                elif isinstance(node, ast.ImportFrom):
                    modules = [node.module or ""]
                self.assertFalse(forbidden.intersection(name.split(".")[0] for name in modules), str(path))


if __name__ == "__main__":
    unittest.main()
