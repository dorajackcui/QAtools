from pathlib import Path
import os
import stat
import sys
import tempfile
import unittest
from unittest.mock import patch

from tools.content_sync.preflight import DirectoryCheck, inspect_master, inspect_targets, is_readonly, master_warnings
from tools.excel_file_ops import create_output_directory


class SelectionCheckTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)

    def test_scan_matches_processing_exclusions_and_does_not_open_excel(self):
        master = self.root / "master.xlsx"
        for name in ("master.xlsx", "a.xlsx", "b.xlsm", "old.xls", "~$a.xlsx", "note.txt"):
            (self.root / name).write_bytes(b"only checking metadata")
        backup = self.root / ".qatools-backups"
        backup.mkdir()
        (backup / "ignored.xlsx").touch()
        output = self.root / "output"
        create_output_directory(output)
        (output / "ignored.xlsx").touch()
        for reverse in (False, True):
            with patch("openpyxl.load_workbook", side_effect=AssertionError("eager Excel read")):
                result = inspect_targets(str(self.root), master=str(master), reverse=reverse)
            self.assertEqual((result.total, result.unsupported), (2, 1))
            self.assertEqual(result.sampled, ("a.xlsx", "b.xlsm"))

    def test_random_sample_is_bounded_and_readonly_errors_are_reported_separately(self):
        files = [self.root / f"{i:02}.xlsx" for i in range(25)]
        for path in files:
            path.touch()
        def check(path):
            if path == files[1]:
                raise PermissionError("unavailable")
            return path == files[0]
        with patch("tools.content_sync.preflight.random.sample", return_value=files[:5]) as sample, \
             patch("tools.content_sync.preflight.is_readonly", side_effect=check) as readonly:
            result = inspect_targets(str(self.root))
        self.assertEqual(sample.call_args.args[1], 5)
        self.assertEqual(readonly.call_count, 5)
        self.assertEqual(result.total, 25)
        self.assertEqual(result.readonly, ("00.xlsx",))
        self.assertIn("01.xlsx", result.errors[0])
        self.assertEqual(result.describe(), "文件只读，请取消read-only或选择【新输出目录】")
        unknown = DirectoryCheck(25, 0, result.sampled, (), result.errors)
        self.assertEqual(unknown.describe(), "无法检查文件状态，请确认文件访问权限。")

    def test_empty_or_missing_directory(self):
        result = inspect_targets(str(self.root))
        self.assertEqual((result.total, result.sampled), (0, ()))
        self.assertEqual(result.describe(), "可处理小表：0个")
        with self.assertRaises(ValueError):
            inspect_targets(str(self.root / "missing"))

    def test_real_file_readonly_attribute_is_detected_without_modification(self):
        path = self.root / "a.xlsx"
        path.write_bytes(b"original")
        try:
            path.chmod(stat.S_IREAD)
            self.assertTrue(is_readonly(path))
            self.assertEqual(inspect_targets(str(self.root)).readonly, ("a.xlsx",))
            self.assertEqual(path.read_bytes(), b"original")
        finally:
            path.chmod(stat.S_IREAD | stat.S_IWRITE)

    def test_lock_error_is_distinct_from_permissions_and_preserved_if_sheet_read_fails(self):
        path = self.root / "a.xlsx"
        path.write_bytes(b"invalid workbook")
        error = OSError("shared")
        error.winerror = 32
        with patch("tools.content_sync.preflight.sys.platform", "win32"), \
             patch("tools.content_sync.preflight._probe_windows_file", side_effect=error):
            result = inspect_master(str(path))
        self.assertIn("占用", result.warnings[0])
        self.assertIsNone(result.choices)
        self.assertTrue(result.error)
        error.winerror = 5
        with patch("tools.content_sync.preflight.sys.platform", "win32"), \
             patch("tools.content_sync.preflight._probe_windows_file", side_effect=error):
            self.assertIn("无法确认", master_warnings(path)[0])

    def test_lock_marker_is_only_a_hint(self):
        path = self.root / "a.xlsx"
        path.touch()
        (self.root / "~$a.xlsx").touch()
        with patch("tools.content_sync.preflight._probe_windows_file"):
            message = master_warnings(path)[0]
        self.assertIn("可能", message)
        self.assertIn("残留", message)

    @unittest.skipUnless(sys.platform == "win32", "Windows sharing semantics")
    def test_real_open_handle_is_detected_and_probe_releases_its_handle(self):
        path = self.root / "a.xlsx"
        path.write_bytes(b"original")
        self.assertEqual(master_warnings(path), ())
        with path.open("rb"):
            self.assertIn("占用", master_warnings(path)[0])
        self.assertEqual(master_warnings(path), ())
        moved = self.root / "moved.xlsx"
        os.replace(path, moved)
        self.assertEqual(moved.read_bytes(), b"original")
