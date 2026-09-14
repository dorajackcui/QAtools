from contextlib import redirect_stderr, redirect_stdout
import io
import os
from pathlib import Path
import shutil
import tempfile
import unittest
from unittest.mock import patch

from qatools.cli import main
from tools.excel_file_ops import create_output_directory
from tools.file_collector.collector import build_copy_plan, execute_copy_plan, parse_names


class FileCollectorTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.source = self.root / "source"
        self.source.mkdir()
        self.output = self.root / "output"

    def file(self, relative, data=b"unchanged workbook bytes"):
        path = self.source / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(data)
        return path

    def plan(self, text, **kwargs):
        return build_copy_plan(self.source, text, output_dir=self.output, **kwargs)

    def invoke(self, *args):
        output, errors = io.StringIO(), io.StringIO()
        with redirect_stdout(output), redirect_stderr(errors):
            code = main([str(arg) for arg in args])
        return code, output.getvalue(), errors.getvalue()

    def test_parse_preserves_spaces_and_commas_deduplicates_case_and_handles_bom(self):
        self.assertEqual(parse_names("\ufeff UI 文本.xlsx \r\nui 文本.XLSX\tA,B\n\n"), ("UI 文本.xlsx", "A,B"))
        self.assertEqual(parse_names("A,B，C\t D E ", comma_separated=True), ("A", "B", "C", "D E"))
        for value in ("", " \n\t", "../a.xlsx", r"sub\a", "*.xlsx", "A.xlsx:stream", "a\0b"):
            with self.subTest(value=value), self.assertRaises(ValueError):
                parse_names(value)

    def test_exact_extension_stem_multiple_formats_and_overlapping_requests(self):
        self.file("nested/UI 文本.XLSX")
        self.file("UI 文本.xlsm")
        self.file("UI 文本.txt")
        self.file("a.v2.xlsb")
        self.assertEqual(len(self.plan("ui 文本.xlsx").entries), 1)
        plan = self.plan("UI 文本\nui 文本.xlsx\na.v2\nmissing")
        self.assertEqual(len(plan.entries), 3)
        self.assertEqual(plan.not_found, ("missing",))
        self.assertEqual(next(e for e in plan.entries if e.source.suffix == ".XLSX").names,
                         ("UI 文本", "ui 文本.xlsx"))
        self.assertFalse(self.output.exists(), "preview must not create output")
        summary = execute_copy_plan(plan)
        self.assertEqual(summary.succeeded_files, 3)
        self.assertEqual(len(summary.warnings), 1)
        for entry in plan.entries:
            self.assertEqual((self.output / entry.destination).read_bytes(), (self.source / entry.source).read_bytes())
        self.assertEqual(len(list(self.output.iterdir())), 3, "no extra report files")

    def test_flat_conflicts_skip_all_candidates_but_copy_unambiguous_file(self):
        self.file("en/A.xlsx", b"english")
        self.file("zh/a.XLSX", b"chinese")
        self.file("b.xls")
        plan = self.plan("a\nb")
        self.assertEqual((plan.ready_count, plan.conflict_count), (1, 2))
        logs, progress = [], []
        summary = execute_copy_plan(plan, log_callback=logs.append, progress_callback=lambda *p: progress.append(p))
        self.assertEqual((summary.succeeded_files, summary.skipped_files), (1, 2))
        self.assertEqual([p.name for p in self.output.iterdir()], ["b.xls"])
        self.assertEqual(progress[-1], (3, 3))
        self.assertTrue(any("[跳过]" in line and "en" in line for line in logs))

    def test_preserve_tree_copies_every_duplicate_and_keeps_sources(self):
        originals = {self.file("en/a.xlsx", b"en"): b"en", self.file("zh/a.xlsx", b"zh"): b"zh"}
        plan = self.plan("a", preserve_tree=True)
        self.assertEqual(plan.ready_count, 2)
        self.assertEqual(execute_copy_plan(plan).succeeded_files, 2)
        for path, data in originals.items():
            self.assertEqual(path.read_bytes(), data)
            self.assertEqual((self.output / path.relative_to(self.source)).read_bytes(), data)

    @unittest.skipIf(os.name == "nt", "requires a case-sensitive source filesystem")
    def test_casefolded_file_directory_output_conflicts(self):
        self.file("A.xlsx", b"file")
        self.file("a.XLSX/b.xls", b"nested")
        plan = self.plan("a.xlsx\nb.xls", preserve_tree=True)
        self.assertEqual(plan.conflict_count, 2)
        self.assertEqual(execute_copy_plan(plan).succeeded_files, 0)
        self.assertFalse(self.output.exists())

    def test_scanner_excludes_locks_backups_generated_trees_and_non_excel(self):
        self.file("~$a.xlsx")
        self.file(".qatools-backups/a.xlsx")
        self.file("a.txt")
        generated = self.source / "earlier"
        create_output_directory(generated)
        self.file("earlier/a.xlsx")
        self.assertEqual(self.plan("a").not_found, ("a",))

    def test_symlinks_are_not_candidates(self):
        self.file("real/a.xlsx")
        try:
            (self.source / "alias.xlsx").symlink_to(self.source / "real/a.xlsx")
            (self.source / "loop").symlink_to(self.source, target_is_directory=True)
        except OSError as exc:
            self.skipTest(f"Symlink creation unavailable: {exc}")
        plan = self.plan("a\nalias")
        self.assertEqual(len(plan.entries), 1)
        self.assertEqual(plan.not_found, ("alias",))

    def test_invalid_output_and_empty_source_arguments_never_write(self):
        for output in ("", self.source, self.source / "child", self.root):
            with self.subTest(output=output), self.assertRaises(ValueError):
                build_copy_plan(self.source, "a", output_dir=output)
        with self.assertRaises(ValueError):
            build_copy_plan("", "a", output_dir=self.output)
        self.assertEqual(list(self.source.iterdir()), [])

    def test_no_matches_or_only_conflicts_do_not_create_output(self):
        summary = execute_copy_plan(self.plan("missing"))
        self.assertEqual(summary.succeeded_files, 0)
        self.assertEqual(summary.details["not_found"], ["missing"])
        self.assertFalse(self.output.exists())
        self.file("one/a.xlsx")
        self.file("two/a.xlsx")
        self.assertEqual(execute_copy_plan(self.plan("a")).skipped_files, 2)
        self.assertFalse(self.output.exists())

    def test_source_mutation_removal_and_new_duplicate_invalidate_preview(self):
        path = self.file("a.xlsx")
        plan = self.plan("a")
        path.write_bytes(b"changed size")
        with self.assertRaisesRegex(ValueError, "重新预览"):
            execute_copy_plan(plan)
        plan = self.plan("a")
        duplicate = self.file("sub/a.xlsx")
        with self.assertRaisesRegex(ValueError, "重新预览"):
            execute_copy_plan(plan)
        duplicate.unlink()
        plan = self.plan("a")
        path.unlink()
        with self.assertRaisesRegex(ValueError, "重新预览"):
            execute_copy_plan(plan)
        self.assertFalse(self.output.exists())

    def test_new_output_created_after_preview_is_untouched(self):
        self.file("a.xlsx")
        plan = self.plan("a")
        self.output.mkdir()
        target = self.output / "a.xlsx"
        target.write_bytes(b"other work")
        with self.assertRaisesRegex(ValueError, "已存在"):
            execute_copy_plan(plan)
        self.assertEqual(target.read_bytes(), b"other work")

    def test_publish_race_does_not_overwrite_and_removes_temporary_file(self):
        self.file("a.xlsx")
        plan = self.plan("a")
        method = "rename" if os.name == "nt" else "link"
        publish = getattr(os, method)

        def raced(source, destination):
            Path(destination).write_bytes(b"concurrent writer")
            publish(source, destination)

        with patch(f"tools.file_collector.collector.os.{method}", side_effect=raced):
            summary = execute_copy_plan(plan)
        self.assertEqual(summary.failed_files, 1)
        self.assertEqual((self.output / "a.xlsx").read_bytes(), b"concurrent writer")
        self.assertEqual(len(list(self.output.iterdir())), 1)

    def test_copy_failure_continues_and_changed_file_is_not_published(self):
        first = self.file("a.xlsx", b"original")
        self.file("b.xlsx", b"second")
        copyfile = shutil.copyfile

        def changed(source, destination):
            copyfile(source, destination)
            if Path(source) == first:
                first.write_bytes(b"changed during copy")

        with patch("tools.file_collector.collector.shutil.copyfile", side_effect=changed):
            summary = execute_copy_plan(self.plan("a\nb"))
        self.assertEqual((summary.failed_files, summary.succeeded_files), (1, 1))
        self.assertEqual([p.name for p in self.output.iterdir()], ["b.xlsx"])

    def test_cli_dry_run_real_copy_and_problem_exit_codes(self):
        self.file("nested/A 文本.xlsm")
        listing = self.root / "names.txt"
        listing.write_text("A 文本,missing", encoding="utf-8-sig")
        args = ("collect-files", self.source, "--names-file", listing, "-o", self.output, "--comma-separated")
        code, output, _ = self.invoke(*args, "--dry-run")
        self.assertEqual(code, 3)
        self.assertIn("未找到", output)
        self.assertFalse(self.output.exists())
        code, output, errors = self.invoke(*args, "--preserve-tree", "--quiet")
        self.assertEqual(code, 3, errors)
        self.assertIn("missing", errors)
        self.assertEqual((self.output / "nested/A 文本.xlsm").read_bytes(), (self.source / "nested/A 文本.xlsm").read_bytes())
        self.assertEqual(self.invoke("collect-files", self.source, "--names-file", listing)[0], 2)
        self.assertEqual(self.invoke(*args)[0], 1)

    def test_cli_success_and_all_failure_exit_codes(self):
        self.file("a.xlsx")
        listing = self.root / "names.txt"
        listing.write_text("a", encoding="utf-8")
        args = ("collect-files", self.source, "--names-file", listing, "-o", self.output)
        self.assertEqual(self.invoke(*args, "--dry-run")[0], 0)
        with patch("tools.file_collector.collector.shutil.copyfile", side_effect=PermissionError("file locked")):
            code, _, errors = self.invoke(*args, "--quiet")
            self.assertEqual(code, 1)
            self.assertIn("file locked", errors)
        self.output = self.root / "successful"
        self.assertEqual(self.invoke(*args[:-1], self.output)[0], 0)


if __name__ == "__main__":
    unittest.main()
