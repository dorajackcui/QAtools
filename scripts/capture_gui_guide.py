"""Render the real Qt pages for the offline guide, using isolated configuration."""
from __future__ import annotations

import os
from pathlib import Path
import stat
import sys
import tempfile
from unittest.mock import patch

os.environ["QT_QPA_PLATFORM"] = "offscreen"
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def main():
    from PySide6.QtCore import QTimer
    from PySide6.QtWidgets import QApplication
    from PySide6.QtGui import QFont, QFontDatabase
    from openpyxl import Workbook
    from toolshub_gui import ToolshubApp
    from tools.header_aliases import HeaderAliasStore
    from tools.qt_gui_common import configure_qt_application
    from tools.content_sync.master_to_target import sync_master_to_targets
    from tools.content_sync.preflight import inspect_targets

    assets = ROOT / "docs/qa-workflow-guide/assets"
    assets.mkdir(parents=True, exist_ok=True)
    app = QApplication.instance() or QApplication([])
    # The offscreen plugin has no Windows font discovery. Load local fonts for
    # rendering only; this does not change application preferences or ship fonts.
    if not QFontDatabase.families() and sys.platform == "win32":
        fonts = Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts"
        for name in ("msyh.ttc", "msyhbd.ttc", "seguisym.ttf"):
            QFontDatabase.addApplicationFont(str(fonts / name))
        app.setFont(QFont("Microsoft YaHei UI"))
    if not QFontDatabase.families():
        raise RuntimeError("截图环境缺少字体，请在有中文字体的 Windows 环境运行。")
    configure_qt_application(app)
    with tempfile.TemporaryDirectory(prefix="qatools-guide-") as directory:
        temporary = Path(directory)
        with patch("tools.tb_projects.default_tb_projects_path", return_value=temporary / "tb.json"), \
             patch("tools.workflow.gui_options.default_header_aliases_path", return_value=temporary / "aliases.json"):
            window = ToolshubApp(show_window=False, header_alias_store=HeaderAliasStore(temporary / "aliases.json"))
        window.resize(1260, 820)
        window.show()

        def capture(widget, name):
            app.processEvents()
            widget.layout().activate()
            app.processEvents()
            if not widget.grab().save(str(assets / name)):
                raise RuntimeError(f"Could not save {name}")
            print(name)

        capture(window, "workflow-main.png")
        workflow = window.tool_frames["workflow"]
        for name, dialog_name in (("term-settings", "term"), ("tag-settings", "tag"), ("target-text-settings", "target"), ("substring-settings", "substring")):
            workflow._open_settings_dialog(dialog_name)
            dialog = getattr(workflow, f"{dialog_name}_settings_dialog")
            capture(dialog, name + ".png")
            dialog.reject()

        for key, name in (("phraseloom", "phraseloom"), ("content_sync", "sync-master-to-target"), ("compatibility", "compatibility"),
                          ("column_tools", "columns"), ("deep_replace", "deep-replace"),
                          ("file_collector", "file-collector"),
                          ("excel_batcher", "batch-split"), ("excel_merger", "merge-sheets"),
                          ("french_nbsp", "french-nbsp"), ("xbench_report", "xbench"),
                          ("untranslated_stats", "untranslated-stats"), ("settings", "header-aliases")):
            window.select_tool(key)
            capture(window, name + ".png")
        window.select_tool("excel_batcher")
        window.tool_frames["excel_batcher"].tabs.setCurrentIndex(1)
        capture(window, "batch-restore.png")
        window.select_tool("content_sync")
        sync = window.tool_frames["content_sync"]
        sync.tabs.setCurrentIndex(1)
        capture(window, "sync-target-to-master.png")

        # Produce real compact logs from an actual synthetic sync. Omit only the
        # two introductory path records, so screenshots contain no personal paths.
        master, targets = temporary / "master.xlsx", temporary / "small"
        targets.mkdir()
        for path, rows in ((master, [["id", "key", "source", "target"], [1, "k", "Hello", "你好"]]),
                           (targets / "sample.xlsx", [["key", "source", "target"], ["k", "Hello", None]])):
            book = Workbook()
            try:
                for row in rows:
                    book.active.append(row)
                book.save(path)
            finally:
                book.close()
        logs = []
        sync_master_to_targets(master, targets, log_callback=logs.append)
        dialog = sync.master_to_target_page.logs
        dialog.start_run()
        for message in logs[2:]:
            dialog.buffer.append(message)
        dialog.finish_run()
        dialog.open_logs()
        capture(dialog, "operation-logs.png")
        dialog.close()
        # Capture the actual selection dialogs using only the synthetic files.
        def capture_check(name):
            failures = []

            def finish():
                popup = app.activeModalWidget()
                try:
                    capture(popup, name)
                except Exception as error:
                    failures.append(error)
                finally:
                    if popup is not None:
                        popup.accept()

            QTimer.singleShot(50, finish)
            sync.master_to_target_page._targets_checked(inspect_targets(str(targets)))
            if failures:
                raise failures[0]

        capture_check("sync-directory-ok.png")
        sample = targets / "sample.xlsx"
        try:
            sample.chmod(stat.S_IREAD)
            capture_check("sync-directory-readonly.png")
        finally:
            sample.chmod(stat.S_IREAD | stat.S_IWRITE)
        window.close()
        app.processEvents()


if __name__ == "__main__":
    main()
