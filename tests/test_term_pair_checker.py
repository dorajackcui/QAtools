from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from tools.term_pair_checker.extract_terms_from_excel import extract_terms, process_excel


class ExtractTermsTests(unittest.TestCase):
    def test_extract_terms_supports_multiple_selected_tag_types_in_text_order(self) -> None:
        text = "前缀[方括号]中间<尖括号>后缀【书名号】"
        self.assertEqual(
            extract_terms(text, mark_styles=("[]", "<>", "【】")),
            ["[方括号]", "<尖括号>", "【书名号】"],
        )

    def test_extract_terms_keeps_fullwidth_square_bracket_compatibility(self) -> None:
        self.assertEqual(extract_terms("这里有［全角方括号］", mark_styles=("[]",)), ["［全角方括号］"])

    def test_extract_terms_requires_at_least_one_tag_type(self) -> None:
        with self.assertRaisesRegex(ValueError, "请至少选择一种 mark 类型"):
            extract_terms("任意文本", mark_styles=())


class ProcessExcelTests(unittest.TestCase):
    def create_workbook(self, path: Path) -> None:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Data"
        worksheet["A1"] = "source"
        worksheet["B1"] = "target"
        worksheet["A2"] = "第一行 [Alpha] 和 <Beta>"
        worksheet["B2"] = "第一行 [阿尔法] 和 <贝塔>"
        worksheet["A3"] = "第二行【Gamma】"
        worksheet["B3"] = "第二行【伽马】"
        worksheet["A4"] = "第三行复用 <Beta>"
        worksheet["B4"] = "第三行复用 <错误贝塔>"
        worksheet["A5"] = "第四行 [Alpha] 加【Gamma】"
        worksheet["B5"] = "第四行只有 [阿尔法]"
        workbook.save(path)

    def test_process_excel_supports_multiple_tag_types(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            self.create_workbook(input_path)
            expected_output_path = Path(tmp_dir) / "input_term_pairs.xlsx"

            worksheet_title, source_col, target_col, saved_path, term_count, problem_count = process_excel(
                input_file=input_path,
                source_column="A",
                target_column="B",
                start_row=2,
                mark_styles=("[]", "<>", "【】"),
            )

            self.assertEqual(worksheet_title, "Data")
            self.assertEqual(source_col, "A")
            self.assertEqual(target_col, "B")
            self.assertEqual(saved_path, expected_output_path.resolve())
            self.assertEqual(term_count, 3)
            self.assertEqual(problem_count, 2)

            original_workbook = load_workbook(input_path)
            self.assertEqual(original_workbook.sheetnames, ["Data"])
            self.assertEqual(original_workbook["Data"]["A2"].value, "第一行 [Alpha] 和 <Beta>")

            workbook = load_workbook(saved_path)
            term_sheet = workbook["术语表"]
            self.assertEqual(term_sheet["A2"].value, "[Alpha]")
            self.assertEqual(term_sheet["B2"].value, "[阿尔法]")
            self.assertEqual(term_sheet["A3"].value, "<Beta>")
            self.assertEqual(term_sheet["B3"].value, "<贝塔>")
            self.assertEqual(term_sheet["A4"].value, "【Gamma】")
            self.assertEqual(term_sheet["B4"].value, "【伽马】")

            problem_sheet = workbook["问题列"]
            self.assertEqual(problem_sheet["A2"].value, 4)
            self.assertEqual(problem_sheet["B2"].value, "术语未对齐")
            self.assertIn("预期target=<贝塔>", str(problem_sheet["C2"].value))
            self.assertEqual(problem_sheet["A3"].value, 5)
            self.assertEqual(problem_sheet["B3"].value, "术语数量不一致")
            self.assertIn("source=[Alpha]、【Gamma】", str(problem_sheet["C3"].value))
            self.assertIn("target=[阿尔法]", str(problem_sheet["C3"].value))

    def test_process_excel_rejects_empty_selected_tag_types(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            self.create_workbook(input_path)

            with self.assertRaisesRegex(ValueError, "请至少选择一种 mark 类型"):
                process_excel(
                    input_file=input_path,
                    source_column="A",
                    target_column="B",
                    start_row=2,
                    mark_styles=(),
                )


if __name__ == "__main__":
    unittest.main()
