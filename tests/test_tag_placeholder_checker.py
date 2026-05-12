from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from tools.tag_placeholder_checker.check_tags_and_placeholders import (
    extract_tokens,
    process_excel,
)


class ExtractTokensTests(unittest.TestCase):
    def test_extract_tokens_supports_mixed_token_types_in_text_order(self) -> None:
        text = r"前缀</text>{name}\n中间<color=red>后缀"
        self.assertEqual(
            extract_tokens(text, token_types=("angle", "brace", "newline")),
            ["</text>", "{name}", r"\n", "<color=red>"],
        )

    def test_extract_tokens_requires_at_least_one_type(self) -> None:
        with self.assertRaisesRegex(ValueError, "请至少选择一种检查类型"):
            extract_tokens("任意文本", token_types=())

    def test_extract_tokens_only_keeps_configured_angle_tag_patterns_by_default(self) -> None:
        text = "忽略 <apple>，保留 </text> 和 <color=red>"
        self.assertEqual(
            extract_tokens(text, token_types=("angle",)),
            ["</text>", "<color=red>"],
        )


class ProcessExcelTests(unittest.TestCase):
    def create_workbook(self, path: Path) -> None:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Data"
        worksheet["A1"] = "source"
        worksheet["B1"] = "target"
        worksheet["A2"] = "保留 <color=red>{name}</text>"
        worksheet["B2"] = "保留 <color=red>{name}</text>"
        worksheet["A3"] = "缺少占位 {name}"
        worksheet["B3"] = "缺少占位"
        worksheet["A4"] = "多出tag <color=red>"
        worksheet["B4"] = "多出tag <color=red></text>"
        worksheet["A5"] = "数量不一致 {count} {count}"
        worksheet["B5"] = "数量不一致 {count}"
        worksheet["A6"] = r"缺少换行 mark \n"
        worksheet["B6"] = "缺少换行 mark"
        worksheet["A7"] = "普通尖括号 <apple>"
        worksheet["B7"] = "普通尖括号"
        workbook.save(path)

    def test_process_excel_writes_problem_and_summary_sheets(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            expected_output_path = Path(tmp_dir) / "input_tag_placeholder_checked.xlsx"
            self.create_workbook(input_path)

            summary = process_excel(
                input_file=input_path,
                source_column="A",
                target_column="B",
                start_row=2,
                token_types=("angle", "brace", "newline"),
            )

            self.assertEqual(summary.worksheet_title, "Data")
            self.assertEqual(summary.output_path, expected_output_path.resolve())
            self.assertEqual(summary.total_rows_checked, 6)
            self.assertEqual(summary.rows_with_selected_tokens, 5)
            self.assertEqual(summary.angle_rows, 2)
            self.assertEqual(summary.brace_rows, 3)
            self.assertEqual(summary.newline_rows, 1)
            self.assertEqual(summary.problem_rows, 4)
            self.assertEqual(summary.problem_count, 4)
            self.assertEqual(summary.selected_token_types, ("angle", "brace", "newline"))

            original_workbook = load_workbook(input_path)
            self.assertEqual(original_workbook.sheetnames, ["Data"])
            self.assertEqual(original_workbook["Data"]["A2"].value, "保留 <color=red>{name}</text>")

            workbook = load_workbook(summary.output_path)
            self.assertIn("标签占位问题", workbook.sheetnames)
            self.assertIn("检查汇总", workbook.sheetnames)

            problem_sheet = workbook["标签占位问题"]
            self.assertEqual(problem_sheet["A2"].value, 3)
            self.assertEqual(problem_sheet["B2"].value, "花括号placeholder不一致")
            self.assertIn("target缺少={name}", str(problem_sheet["C2"].value))
            self.assertEqual(problem_sheet["D2"].value, "缺少占位 {name}")
            self.assertEqual(problem_sheet["E2"].value, "缺少占位")

            self.assertEqual(problem_sheet["A3"].value, 4)
            self.assertEqual(problem_sheet["B3"].value, "尖括号tag不一致")
            self.assertIn("target多出=</text>", str(problem_sheet["C3"].value))

            self.assertEqual(problem_sheet["A4"].value, 5)
            self.assertEqual(problem_sheet["B4"].value, "花括号placeholder不一致")
            self.assertIn("target缺少={count}", str(problem_sheet["C4"].value))

            self.assertEqual(problem_sheet["A5"].value, 6)
            self.assertEqual(problem_sheet["B5"].value, r"\n mark不一致")
            self.assertIn(r"target缺少=\n", str(problem_sheet["C5"].value))

            summary_sheet = workbook["检查汇总"]
            summary_values = {
                summary_sheet.cell(row_index, 1).value: summary_sheet.cell(row_index, 2).value
                for row_index in range(2, 14)
            }
            self.assertEqual(summary_values["检查工作表"], "Data")
            self.assertEqual(summary_values["source列"], "A")
            self.assertEqual(summary_values["target列"], "B")
            self.assertEqual(summary_values["开始行"], 2)
            self.assertEqual(summary_values["检查类型"], r"尖括号tag、花括号placeholder、\n mark")
            self.assertEqual(summary_values["总行数"], 6)
            self.assertEqual(summary_values["命中检查类型行数"], 5)
            self.assertEqual(summary_values["含尖括号tag行数"], 2)
            self.assertEqual(summary_values["含花括号placeholder行数"], 3)
            self.assertEqual(summary_values[r"含\n mark行数"], 1)
            self.assertEqual(summary_values["问题行数"], 4)
            self.assertEqual(summary_values["问题条数"], 4)

    def test_process_excel_supports_single_selected_type(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            self.create_workbook(input_path)

            summary = process_excel(
                input_file=input_path,
                source_column="A",
                target_column="B",
                start_row=2,
                token_types=("angle",),
            )

            self.assertEqual(summary.rows_with_selected_tokens, 2)
            self.assertEqual(summary.angle_rows, 2)
            self.assertEqual(summary.brace_rows, 0)
            self.assertEqual(summary.newline_rows, 0)
            self.assertEqual(summary.problem_rows, 1)
            self.assertEqual(summary.problem_count, 1)

            workbook = load_workbook(summary.output_path)
            problem_sheet = workbook["标签占位问题"]
            self.assertEqual(problem_sheet.max_row, 2)
            self.assertEqual(problem_sheet["B2"].value, "尖括号tag不一致")

    def test_process_excel_supports_newline_selected_type(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            self.create_workbook(input_path)

            summary = process_excel(
                input_file=input_path,
                source_column="A",
                target_column="B",
                start_row=2,
                token_types=("newline",),
            )

            self.assertEqual(summary.rows_with_selected_tokens, 1)
            self.assertEqual(summary.angle_rows, 0)
            self.assertEqual(summary.brace_rows, 0)
            self.assertEqual(summary.newline_rows, 1)
            self.assertEqual(summary.problem_rows, 1)
            self.assertEqual(summary.problem_count, 1)

            workbook = load_workbook(summary.output_path)
            problem_sheet = workbook["标签占位问题"]
            self.assertEqual(problem_sheet.max_row, 2)
            self.assertEqual(problem_sheet["B2"].value, r"\n mark不一致")


if __name__ == "__main__":
    unittest.main()
