from pathlib import Path
import unittest

from openpyxl import Workbook

from tools.consistency_text import normalize_consistency_text, normalized_text_with_offset
from tools.source_consistency_checker.check_source_consistency import process_workbook as check_source
from tools.target_consistency_checker.check_target_consistency import process_workbook as check_target


class ConsistencyNormalizationTests(unittest.TestCase):
    def test_wrappers_and_whitespace_normalize_to_same_key_and_keep_offset(self):
        for text in (
            "阿童木", " “阿童木” ", "'阿童木'", '"阿童木"', "‘阿童木’", "「阿童木」",
            "『阿童木』", "«阿童木»", "‹阿童木›", "＂阿童木＂", "＇阿童木＇",
            "【阿童木】", "[阿童木]", "(阿童木)", "（阿童木）", "《阿童木》",
            "〈阿童木〉", "〔阿童木〕", "〖阿童木〗", "［阿童木］",
            "\t【 ‘ 阿童木 ’ 】\u00a0",
        ):
            with self.subTest(text=text):
                normalized, start = normalized_text_with_offset(text)
                self.assertEqual(normalized, "阿童木")
                self.assertEqual(text[start:start + len(normalized)], normalized)
                self.assertEqual(normalize_consistency_text(normalized), normalized)

    def test_internal_symbols_case_whitespace_and_placeholders_are_preserved(self):
        for text in (
            "【阿童木】登场", "“阿童木”和“悟空”", '"Astro" and "Boy"',
            "'阿童木'和'悟空'", "＇阿童木＇和＇悟空＇", "'A'和'B'",
            "[阿童木]与[悟空]", "“阿童木", "阿童木”", "【阿童木)",
            "don't", "A  B", "A\nB", "Ａ", "{player}", "<name>",
            "cafe\u0301", "Astro",
        ):
            with self.subTest(text=text):
                self.assertEqual(normalize_consistency_text(text), text)
        self.assertEqual(normalize_consistency_text("'don't'"), "don't")
        self.assertEqual(normalize_consistency_text("‘don’t’"), "don’t")
        self.assertEqual(normalize_consistency_text("‘l’été’"), "l’été")
        self.assertEqual(normalize_consistency_text('"Say \\"hi\\""'), 'Say \\"hi\\"')

    def test_empty_wrappers_and_non_string_cells(self):
        for value in (None, "", " \n", '""', "【 ‘ ’ 】"):
            self.assertEqual(normalize_consistency_text(value), "")
        self.assertEqual(normalize_consistency_text(123), "123")

    def test_both_checks_normalize_group_keys_and_versions_preserving_raw_rows(self):
        for checker, problem_sheet, reverse in (
            (check_source, "同源译文不一致", False),
            (check_target, "同Target不同Source", True),
        ):
            with self.subTest(reverse=reverse):
                workbook = Workbook()
                self.addCleanup(workbook.close)
                worksheet = workbook.active
                worksheet.append(["source", "target"])
                rows = [
                    ("【阿童木】", "“Astro Boy”"), ("阿童木", "Astro Boy"),
                    ("‘悟空’", "Goku"), ("悟空", "【Monkey King】"),
                    ("【悟空】", "‘Goku’"),
                    ("【 】", "ignored"),
                ]
                for source, target in rows:
                    worksheet.append((target, source) if reverse else (source, target))
                summary = checker(
                    workbook=workbook, output_path=Path("unused.xlsx"),
                    source_column="A", target_column="B",
                )
                self.assertEqual(summary.problem_rows, 3)
                issues = list(workbook[problem_sheet].values)[1:]
                self.assertEqual([row[0] for row in issues], [4, 5, 6])
                self.assertTrue(all(row[4] == 2 and row[5] == "4、5、6" for row in issues))
                expected = [tuple(reversed(row)) if reverse else row for row in rows[2:5]]
                self.assertEqual([row[1:3] for row in issues], expected)


if __name__ == "__main__":
    unittest.main()
