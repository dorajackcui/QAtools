from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from tools.chinese_target_checker.check_chinese_target import (
    contains_chinese,
    extract_chinese_characters,
    process_excel,
)


class ChineseTargetTextTests(unittest.TestCase):
    def test_contains_chinese_detects_cjk_ideographs_only(self) -> None:
        self.assertTrue(contains_chinese("Keep 中文 text"))
        self.assertFalse(contains_chinese("Only punctuation ，。！？"))
        self.assertFalse(contains_chinese(""))
        self.assertFalse(contains_chinese(None))
        self.assertFalse(contains_chinese(42))

    def test_extract_chinese_characters_preserves_match_order(self) -> None:
        self.assertEqual(extract_chinese_characters("A中B文C中"), "中文中")
        self.assertEqual(extract_chinese_characters("No Chinese"), "")


class ChineseTargetExcelTests(unittest.TestCase):
    def create_workbook(self, path: Path) -> None:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Data"
        worksheet["A1"] = "source"
        worksheet["B1"] = "target"
        worksheet["C1"] = "old result"
        worksheet["D1"] = "custom result"
        worksheet["A2"] = "hello"
        worksheet["B2"] = "Bonjour"
        worksheet["C2"] = "old keep"
        worksheet["A3"] = "hello"
        worksheet["B3"] = "包含中文 target"
        worksheet["C3"] = "old clear"
        worksheet["A4"] = "hello"
        worksheet["B4"] = None
        worksheet["C4"] = "old clear none"
        workbook.save(path)

    def test_process_excel_marks_default_adjacent_result_column(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            expected_output_path = Path(tmp_dir) / "input_chinese_target_checked.xlsx"
            self.create_workbook(input_path)

            summary = process_excel(
                input_file=input_path,
                target_column="B",
                sheet="Data",
                start_row=2,
            )

            self.assertEqual(summary.output_path, expected_output_path.resolve())
            self.assertEqual(summary.worksheet_title, "Data")
            self.assertEqual(summary.target_column, "B")
            self.assertEqual(summary.result_column, "C")
            self.assertEqual(summary.start_row, 2)
            self.assertEqual(summary.processed_count, 3)
            self.assertEqual(summary.matched_count, 1)
            self.assertFalse(summary.problem_sheet_created)

            original_workbook = load_workbook(input_path)
            self.assertEqual(original_workbook["Data"]["C1"].value, "old result")
            self.assertEqual(original_workbook["Data"]["C3"].value, "old clear")

            output_workbook = load_workbook(summary.output_path)
            output_sheet = output_workbook["Data"]
            self.assertEqual(output_sheet["C1"].value, "中文检查")
            self.assertIsNone(output_sheet["C2"].value)
            self.assertEqual(output_sheet["C3"].value, "含中文")
            self.assertIsNone(output_sheet["C4"].value)
            self.assertNotIn("中文检查问题", output_workbook.sheetnames)

    def test_process_excel_supports_custom_result_column_and_problem_sheet(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            output_path = Path(tmp_dir) / "output.xlsx"
            self.create_workbook(input_path)

            summary = process_excel(
                input_file=input_path,
                target_column="B",
                result_column="D",
                sheet="Data",
                start_row=2,
                create_problem_sheet=True,
                output_file=output_path,
            )

            self.assertEqual(summary.output_path, output_path.resolve())
            self.assertEqual(summary.result_column, "D")
            self.assertEqual(summary.processed_count, 3)
            self.assertEqual(summary.matched_count, 1)
            self.assertTrue(summary.problem_sheet_created)

            output_workbook = load_workbook(summary.output_path)
            output_sheet = output_workbook["Data"]
            self.assertEqual(output_sheet["D1"].value, "中文检查")
            self.assertIsNone(output_sheet["D2"].value)
            self.assertEqual(output_sheet["D3"].value, "含中文")
            self.assertIsNone(output_sheet["D4"].value)

            problem_sheet = output_workbook["中文检查问题"]
            self.assertEqual(problem_sheet["A1"].value, "行号")
            self.assertEqual(problem_sheet["B1"].value, "target文本")
            self.assertEqual(problem_sheet["C1"].value, "中文字符")
            self.assertEqual(problem_sheet["A2"].value, 3)
            self.assertEqual(problem_sheet["B2"].value, "包含中文 target")
            self.assertEqual(problem_sheet["C2"].value, "包含中文")

    def test_process_excel_rejects_invalid_start_row(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            self.create_workbook(input_path)

            with self.assertRaisesRegex(ValueError, "开始行必须大于等于 1"):
                process_excel(
                    input_file=input_path,
                    target_column="B",
                    start_row=0,
                )


if __name__ == "__main__":
    unittest.main()
