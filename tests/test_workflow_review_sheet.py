from __future__ import annotations

import unittest
from io import BytesIO
from pathlib import Path

from openpyxl import Workbook, load_workbook

from tools.excel_output import PROBLEM_BASE_HEADERS
from tools.workflow.review_sheet import collect_review_rows, write_review_sheet


class ReviewSheetOrderingTests(unittest.TestCase):
    def test_saved_review_only_unlocks_revision_cells_and_allows_selection_and_filter(self):
        for has_issue in (False, True):
            with self.subTest(has_issue=has_issue), BytesIO() as output:
                workbook = Workbook()
                try:
                    workbook.active.title = "Data"
                    workbook.active.append(["source", "target"])
                    workbook.active.append(["Hello {name}", "Hello"])
                    problem = workbook.create_sheet("Tag 问题")
                    problem.append(PROBLEM_BASE_HEADERS)
                    if has_issue:
                        problem.append([2, "Hello {name}", "Hello", "缺少：{name}"])
                    count = write_review_sheet(
                        workbook, current_sheet_name="Data", input_file=Path("input.xlsx"),
                        source_column="A", target_column="B", start_row=2,
                        problem_sheets=(("Tag 检查", problem.title),),
                        generated_sheet_names=("问题处理",),
                    )
                    self.assertEqual(count, int(has_issue))
                    workbook.save(output)
                finally:
                    workbook.close()
                output.seek(0)
                saved = load_workbook(output)
                try:
                    sheet = saved["问题处理"]
                    self.assertTrue(sheet.protection.sheet)
                    self.assertIsNone(sheet.protection.password)
                    self.assertFalse(sheet.protection.selectLockedCells)
                    self.assertFalse(sheet.protection.selectUnlockedCells)
                    self.assertFalse(sheet.protection.autoFilter)
                    self.assertEqual(sheet.auto_filter.ref, f"A1:F{count + 1}")
                    self.assertTrue(all(cell.protection.locked for cell in sheet[1]))
                    for column in (1, 2, 3, 5, 6):
                        self.assertTrue(sheet.cell(2, column).protection.locked)
                    self.assertEqual(sheet["D2"].protection.locked, not has_issue)
                    self.assertTrue(sheet["D3"].protection.locked)
                    self.assertFalse(saved["Data"].protection.sheet)
                    self.assertEqual(saved["Data"]["B2"].value, "Hello")
                    if has_issue:
                        self.assertEqual(sheet["A2"].hyperlink.location, "'Data'!B2")
                        self.assertEqual(sheet["B2"].value, "Hello {name}")
                        self.assertEqual(sheet["C2"].value, "Hello")
                finally:
                    saved.close()

    def test_consistency_groups_come_first_without_splitting_multiple_issues(self) -> None:
        workbook = Workbook()
        try:
            tag_sheet = workbook.active
            tag_sheet.title = "Tag 问题"
            tag_sheet.append(PROBLEM_BASE_HEADERS)
            tag_sheet.append([5, "Other", "Other target", "Tag 缺失"])
            tag_sheet.append([20, "Save", "Save A", "Tag 缺失"])
            tag_sheet.append([2, "Plain", "Plain target", "Tag 缺失"])

            source_sheet = workbook.create_sheet("Source 问题")
            source_sheet.append(PROBLEM_BASE_HEADERS + ("同组行号",))
            for row, source, target, grouped_rows in (
                (90, "Save", "Save B", "20、50、80、90"),
                (60, "Open", "Open B", "40、60"),
                (20, "Save", "Save A", "20、50、80、90"),
                (40, "Open", "Open A", "40、60"),
                (50, "Save", "Save B", "20、50、80、90"),
                (80, "Save", "Save A", "20、50、80、90"),
            ):
                source_sheet.append([row, source, target, f"2 种译法：1: {source} B；2: {source} A", grouped_rows])

            target_sheet = workbook.create_sheet("Target 问题")
            target_sheet.append(PROBLEM_BASE_HEADERS + ("同组行号",))
            for row, source, target, grouped_rows in (
                (100, "Extra", "Save A", "20、80、100"),
                (80, "Save", "Save A", "20、80、100"),
                (70, "Zulu", "Shared Z", "10、70"),
                (20, "Save", "Save A", "20、80、100"),
                (30, "Beta", "Shared A", "30、110"),
                (10, "Alpha", "Shared Z", "10、70"),
                (110, "Gamma", "Shared A", "30、110"),
            ):
                variants = {"Save A": "1: Extra；2: Save", "Shared Z": "1: Zulu；2: Alpha", "Shared A": "1: Beta；2: Gamma"}
                target_sheet.append([row, source, target, f"2 种原文：{variants[target]}", grouped_rows])

            checks = (
                ("Tag 检查", tag_sheet.title),
                ("同 Target 不同 Source", target_sheet.title),
                ("同 Source 不同 Target", source_sheet.title),
            )
            rows = collect_review_rows(workbook, checks)
            self.assertEqual(
                [row[0] for row in rows],
                [20, 80, 50, 90, 40, 60, 10, 70, 30, 110, 100, 2, 5],
            )
            self.assertEqual(len({row[0] for row in rows}), len(rows))
            self.assertEqual(
                rows[0],
                (
                    20,
                    "Save",
                    "Save A",
                    None,
                    "【Tag 检查】Tag 缺失；"
                    "【同 Target 不同 Source】2 种原文：1: Extra；2: Save；"
                    "【同 Source 不同 Target】2 种译法：1: Save B；2: Save A",
                    "Tag 检查；同 Target 不同 Source；同 Source 不同 Target",
                ),
            )
            self.assertEqual(
                [row[:4] for row in collect_review_rows(workbook, reversed(checks))],
                [row[:4] for row in rows],
            )
        finally:
            workbook.close()

    def test_ordering_uses_normalized_groups_and_keeps_blank_target_variants(self) -> None:
        workbook = Workbook()
        try:
            sheet = workbook.active
            sheet.append(PROBLEM_BASE_HEADERS)
            for row, source, target in (
                (12, "Same ", "Y"),
                (4, "Same", None),
                (10, "Same", "X"),
                (6, "【Same】", "“X”"),
                (14, "Same", None),
                (8, "same", "X"),
                (16, "same", "Y"),
            ):
                sheet.append([row, source, target, "译文不一致"])

            rows = collect_review_rows(
                workbook, (("同 Source 不同 Target", sheet.title),)
            )
            self.assertEqual([row[0] for row in rows], [4, 14, 6, 10, 12, 8, 16])
            self.assertEqual([row[2] for row in rows[:3]], ["", "", "“X”"])
        finally:
            workbook.close()


if __name__ == "__main__":
    unittest.main()
