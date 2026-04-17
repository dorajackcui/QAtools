from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from tools.term_pair_checker.extract_terms_from_excel import (
    extract_terms,
    process_excel,
)


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

    def test_extract_terms_uses_default_json_exclusions_for_false_positive_tags(self) -> None:
        text = "样式 </> <color=red> <outline color=blue> 真术语 <苹果>"
        self.assertEqual(extract_terms(text, mark_styles=("<>",)), ["<苹果>"])

    def test_extract_terms_supports_custom_json_exclusion_config(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            config_path = Path(tmp_dir) / "custom_exclusions.json"
            config_path.write_text(
                json.dumps({"patterns": [r"^apple$"]}, ensure_ascii=False),
                encoding="utf-8",
            )

            text = "保留 <color=red>，排除 <apple>"
            self.assertEqual(
                extract_terms(
                    text,
                    mark_styles=("<>",),
                    exclusion_config_file=config_path,
                ),
                ["<color=red>"],
            )


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
            plain_term_sheet = workbook["术语表（无mark）"]
            self.assertEqual(plain_term_sheet["A2"].value, "Alpha")
            self.assertEqual(plain_term_sheet["B2"].value, "阿尔法")
            self.assertEqual(plain_term_sheet["A3"].value, "Beta")
            self.assertEqual(plain_term_sheet["B3"].value, "贝塔")
            self.assertEqual(plain_term_sheet["A4"].value, "Gamma")
            self.assertEqual(plain_term_sheet["B4"].value, "伽马")

            problem_sheet = workbook["问题列"]
            self.assertEqual(problem_sheet["A2"].value, 4)
            self.assertEqual(problem_sheet["B2"].value, "术语未对齐")
            self.assertIn("source=Beta", str(problem_sheet["C2"].value))
            self.assertIn("预期target=贝塔", str(problem_sheet["C2"].value))
            self.assertIn("实际target=错误贝塔", str(problem_sheet["C2"].value))
            self.assertIn("术语对示例=<Beta> -> <贝塔>", str(problem_sheet["C2"].value))
            self.assertEqual(problem_sheet["D2"].value, "第三行复用 <Beta>")
            self.assertEqual(problem_sheet["E2"].value, "第三行复用 <错误贝塔>")
            self.assertEqual(problem_sheet["A3"].value, 5)
            self.assertEqual(problem_sheet["B3"].value, "术语数量不一致")
            self.assertIn("source=[Alpha]、【Gamma】", str(problem_sheet["C3"].value))
            self.assertIn("target=[阿尔法]", str(problem_sheet["C3"].value))
            self.assertEqual(problem_sheet["D3"].value, "第四行 [Alpha] 加【Gamma】")
            self.assertEqual(problem_sheet["E3"].value, "第四行只有 [阿尔法]")

    def test_process_excel_dedupes_same_plain_term_with_different_marks(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"

            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "Data"
            worksheet["A1"] = "source"
            worksheet["B1"] = "target"
            worksheet["A2"] = "第一行 [Alpha]"
            worksheet["B2"] = "第一行 [阿尔法]"
            worksheet["A3"] = "第二行 <Alpha>"
            worksheet["B3"] = "第二行 <阿尔法>"
            workbook.save(input_path)

            _, _, _, saved_path, term_count, problem_count = process_excel(
                input_file=input_path,
                source_column="A",
                target_column="B",
                start_row=2,
                mark_styles=("[]", "<>"),
            )

            self.assertEqual(term_count, 1)
            self.assertEqual(problem_count, 0)

            result_workbook = load_workbook(saved_path)
            term_sheet = result_workbook["术语表"]
            self.assertEqual(term_sheet["A2"].value, "[Alpha]")
            self.assertEqual(term_sheet["B2"].value, "[阿尔法]")
            self.assertEqual(term_sheet.max_row, 2)
            plain_term_sheet = result_workbook["术语表（无mark）"]
            self.assertEqual(plain_term_sheet["A2"].value, "Alpha")
            self.assertEqual(plain_term_sheet["B2"].value, "阿尔法")
            self.assertEqual(plain_term_sheet.max_row, 2)

    def test_process_excel_retroactively_checks_unmarked_rows(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"

            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "Data"
            worksheet["A1"] = "source"
            worksheet["B1"] = "target"
            worksheet["A2"] = "第三行先出现苹果"
            worksheet["B2"] = "第三行先出现banana"
            worksheet["A3"] = "第十一行才标记出 <苹果>"
            worksheet["B3"] = "第十一行才标记出 <apple>"
            workbook.save(input_path)

            _, _, _, saved_path, term_count, problem_count = process_excel(
                input_file=input_path,
                source_column="A",
                target_column="B",
                start_row=2,
                mark_styles=("<>",),
            )

            self.assertEqual(term_count, 1)
            self.assertEqual(problem_count, 1)

            result_workbook = load_workbook(saved_path)
            problem_sheet = result_workbook["问题列"]
            self.assertEqual(problem_sheet["A2"].value, 2)
            self.assertEqual(problem_sheet["B2"].value, "术语未对齐")
            self.assertIn("source=苹果", str(problem_sheet["C2"].value))
            self.assertIn("预期target=apple", str(problem_sheet["C2"].value))
            self.assertIn("术语对示例=<苹果> -> <apple>", str(problem_sheet["C2"].value))
            self.assertEqual(problem_sheet["D2"].value, "第三行先出现苹果")
            self.assertEqual(problem_sheet["E2"].value, "第三行先出现banana")

    def test_process_excel_treats_marked_target_as_aligned_for_unmarked_source(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"

            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "Data"
            worksheet["A1"] = "source"
            worksheet["B1"] = "target"
            worksheet["A2"] = "前面未标记的苹果"
            worksheet["B2"] = "前面已写成 <apple>"
            worksheet["A3"] = "后面用 <苹果> 建立术语"
            worksheet["B3"] = "后面用 <apple> 建立术语"
            workbook.save(input_path)

            _, _, _, saved_path, term_count, problem_count = process_excel(
                input_file=input_path,
                source_column="A",
                target_column="B",
                start_row=2,
                mark_styles=("<>",),
            )

            self.assertEqual(term_count, 1)
            self.assertEqual(problem_count, 0)

            result_workbook = load_workbook(saved_path)
            problem_sheet = result_workbook["问题列"]
            self.assertEqual(problem_sheet.max_row, 1)

    def test_process_excel_uses_hybrid_boundary_for_retroactive_checks(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"

            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "Data"
            worksheet["A1"] = "source"
            worksheet["B1"] = "target"
            worksheet["A2"] = "account setup"
            worksheet["B2"] = "account setup"
            worksheet["A3"] = "定义 <ACC>"
            worksheet["B3"] = "定义 <ACC>"
            workbook.save(input_path)

            _, _, _, saved_path, term_count, problem_count = process_excel(
                input_file=input_path,
                source_column="A",
                target_column="B",
                start_row=2,
                mark_styles=("<>",),
            )

            self.assertEqual(term_count, 1)
            self.assertEqual(problem_count, 0)

            result_workbook = load_workbook(saved_path)
            self.assertEqual(result_workbook["问题列"].max_row, 1)

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

    def test_process_excel_ignores_false_positive_markup_from_default_config(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"

            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "Data"
            worksheet["A1"] = "source"
            worksheet["B1"] = "target"
            worksheet["A2"] = "</> <color=red> <outline color=blue>"
            worksheet["B2"] = "</> <color=red> <outline color=blue>"
            worksheet["A3"] = "真实术语 <苹果>"
            worksheet["B3"] = "真实术语 <apple>"
            workbook.save(input_path)

            _, _, _, saved_path, term_count, problem_count = process_excel(
                input_file=input_path,
                source_column="A",
                target_column="B",
                start_row=2,
                mark_styles=("<>",),
            )

            self.assertEqual(term_count, 1)
            self.assertEqual(problem_count, 0)

            result_workbook = load_workbook(saved_path)
            term_sheet = result_workbook["术语表"]
            self.assertEqual(term_sheet["A2"].value, "<苹果>")
            self.assertEqual(term_sheet["B2"].value, "<apple>")
            self.assertEqual(term_sheet.max_row, 2)
            plain_term_sheet = result_workbook["术语表（无mark）"]
            self.assertEqual(plain_term_sheet["A2"].value, "苹果")
            self.assertEqual(plain_term_sheet["B2"].value, "apple")
            self.assertEqual(plain_term_sheet.max_row, 2)


if __name__ == "__main__":
    unittest.main()
