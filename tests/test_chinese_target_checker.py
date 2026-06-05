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
    def test_contains_chinese_detects_cjk_ideographs_and_punctuation(self) -> None:
        self.assertTrue(contains_chinese("Keep 中文 text"))
        self.assertTrue(contains_chinese("Only punctuation ，。！？"))
        self.assertTrue(contains_chinese("Fullwidth marks 【】（）《》“”"))
        self.assertTrue(contains_chinese("Typography marks “”‘’—…·"))
        self.assertFalse(contains_chinese(""))
        self.assertFalse(contains_chinese(None))
        self.assertFalse(contains_chinese(42))
        self.assertFalse(contains_chinese("ASCII punctuation ()[]!?"))
        self.assertFalse(contains_chinese("Fullwidth alnum ＡＢ１２"))

    def test_extract_chinese_characters_preserves_match_order(self) -> None:
        self.assertEqual(extract_chinese_characters("A中B，C（D）“E”…"), "中，（）“”…")
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

    def test_process_excel_inserts_default_result_column_in_place(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            self.create_workbook(input_path)

            summary = process_excel(
                input_file=input_path,
                target_column="B",
                sheet="Data",
                start_row=2,
            )

            self.assertEqual(summary.output_path, input_path.with_name("target_chinese_check_input.xlsx").resolve())
            self.assertEqual(summary.worksheet_title, "Data")
            self.assertEqual(summary.target_column, "B")
            self.assertEqual(summary.result_column, "C")
            self.assertEqual(summary.start_row, 2)
            self.assertEqual(summary.processed_count, 3)
            self.assertEqual(summary.matched_count, 1)

            workbook = load_workbook(summary.output_path)
            worksheet = workbook["Data"]
            self.assertEqual(worksheet["C1"].value, "中文检查")
            self.assertIsNone(worksheet["C2"].value)
            self.assertEqual(worksheet["C3"].value, "含中文")
            self.assertIsNone(worksheet["C4"].value)
            self.assertEqual(worksheet["D1"].value, "old result")
            self.assertEqual(worksheet["D2"].value, "old keep")
            self.assertEqual(worksheet["D3"].value, "old clear")
            self.assertEqual(worksheet["D4"].value, "old clear none")
            self.assertEqual(worksheet["E1"].value, "custom result")
            self.assertNotIn("中文检查问题", workbook.sheetnames)

    def test_process_excel_supports_custom_result_column_without_problem_sheet(self) -> None:
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
                output_file=output_path,
            )

            self.assertEqual(summary.output_path, output_path.resolve())
            self.assertEqual(summary.result_column, "D")
            self.assertEqual(summary.processed_count, 3)
            self.assertEqual(summary.matched_count, 1)

            output_workbook = load_workbook(summary.output_path)
            output_sheet = output_workbook["Data"]
            self.assertEqual(output_sheet["D1"].value, "中文检查")
            self.assertIsNone(output_sheet["D2"].value)
            self.assertEqual(output_sheet["D3"].value, "含中文")
            self.assertIsNone(output_sheet["D4"].value)
            self.assertNotIn("中文检查问题", output_workbook.sheetnames)

            input_workbook = load_workbook(input_path)
            self.assertEqual(input_workbook["Data"]["D1"].value, "custom result")
            self.assertIsNone(input_workbook["Data"]["D3"].value)

    def test_process_excel_removes_stale_problem_sheet(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            self.create_workbook(input_path)
            workbook = load_workbook(input_path)
            workbook.create_sheet("中文检查问题")
            workbook.save(input_path)

            summary = process_excel(
                input_file=input_path,
                target_column="B",
                sheet="Data",
                start_row=2,
            )

            output_workbook = load_workbook(summary.output_path)
            self.assertNotIn("中文检查问题", output_workbook.sheetnames)

    def test_process_excel_reuses_existing_adjacent_result_column(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            self.create_workbook(input_path)
            workbook = load_workbook(input_path)
            worksheet = workbook["Data"]
            worksheet.insert_cols(3)
            worksheet["C1"] = "中文检查"
            worksheet["D1"] = "old result"
            worksheet["E1"] = "custom result"
            workbook.save(input_path)

            summary = process_excel(
                input_file=input_path,
                target_column="B",
                sheet="Data",
                start_row=2,
            )

            self.assertEqual(summary.result_column, "C")
            output_workbook = load_workbook(input_path)
            output_sheet = output_workbook["Data"]
            self.assertEqual(output_sheet.max_column, 5)
            self.assertEqual(output_sheet["C1"].value, "中文检查")
            self.assertEqual(output_sheet["D1"].value, "old result")
            self.assertEqual(output_sheet["E1"].value, "custom result")

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
