from __future__ import annotations

from pathlib import Path
from unittest.mock import patch
import random
import unittest

from openpyxl import Workbook

from tools.substring_consistency_checker.check_substring_consistency import (
    DETAIL_LIMIT,
    PROBLEM_SHEET_NAME,
    process_workbook,
)
from tools.term_matching import text_contains_term
from tools.substring_consistency_checker import check_substring_consistency as checker


class SubstringConsistencyTests(unittest.TestCase):
    def test_default_cjk_minimum_and_custom_thresholds(self):
        rows = [
            ("确认", "Confirm"), ("合作", "Cooperate"), ("【阿童木】", "Astro Boy"),
            ("确认操作", "Wrong"), ("双方合作", "Wrong"), ("阿童木登场", "Wrong"),
        ]
        self.assertEqual([row[0] for row in self.check_rows(rows)[1]], [7])
        self.assertEqual([row[0] for row in self.check_rows(rows, min_cjk_chars=2)[1]], [5, 6, 7])
        self.assertEqual(self.check_rows(rows, min_cjk_chars=4)[1], [])

    def test_effective_length_excludes_symbols_and_uses_cjk_for_mixed_text(self):
        for child in ("确认！", "确认123", "确认{name}", "确认<br/>", "A中", "確認", "확인"):
            with self.subTest(child=child):
                self.assertEqual(self.check_rows([(child, "Expected"), (f"前 {child} 后", "Wrong")])[1], [])
        for child in ("确认操作", "確認済", "확인됨", "𠀀𠀁𠀂", "A中B"):
            with self.subTest(child=child):
                self.assertEqual(len(self.check_rows([(child, "Expected"), (f"前 {child} 后", "Wrong")])[1]), 1)
        rows = [("Go{名字}", "Expected"), ("Go{名字} now", "Wrong")]
        self.assertEqual(len(self.check_rows(rows)[1]), 1)
        self.assertEqual(self.check_rows(rows, min_other_chars=3)[1], [])

    def test_invalid_minimum_is_rejected_before_processing(self):
        for value in (0, -1, 1_000_001, True, 2.5, "3"):
            for option in ("min_cjk_chars", "min_other_chars"):
                with self.subTest(value=value, option=option), self.assertRaisesRegex(ValueError, "最小有效字符数"):
                    self.check_rows([], **{option: value})

    def test_terms_exclude_whole_child_or_parent_before_indexing(self):
        rows = [
            ("【阿童木】", "Astro Boy"), ("阿童木登场", "Wrong"),
            ("光合作", "Cooperation"), ("光合作用", "Wrong"),
            ("合作效率", "Efficiency"), ("提高合作效率", "Wrong"),
        ]
        with patch.object(checker, "_build_index", wraps=checker._build_index) as build:
            _, issues, _ = self.check_rows(rows, checked_source_terms=("阿童木", "【光合作用】", "合作"))
        self.assertEqual([issue[0] for issue in issues], [7])
        patterns = build.call_args_list[0].args[0]
        self.assertNotIn("阿童木", patterns)
        self.assertNotIn("光合作用", patterns)
        self.assertIn("合作效率", patterns)

    def test_term_filter_uses_source_equality_not_target_or_partial_overlap(self):
        rows = [("Save", "保存"), ("Save now", "Wrong")]
        self.assertEqual(self.check_rows(rows, checked_source_terms=("【SAVE】",))[1], [])
        self.assertEqual(len(self.check_rows(rows, checked_source_terms=("保存", "Sav"))[1]), 1)
        # A match inside a term occurrence is kept when the whole parent is
        # longer than that term: this filter intentionally uses whole strings.
        rows = [("光合作", "Expected"), ("植物光合作用", "Wrong")]
        self.assertEqual(len(self.check_rows(rows, checked_source_terms=("光合作用",))[1]), 1)

    def test_normalized_references_share_versions_and_keep_original_text(self):
        summary, issues, _ = self.check_rows([
            ("【阿童木】", "‘Astro Boy’"), ("阿童木", "Astro Boy"),
            ("【 阿童木登场 】", "‘Wrong’"), ("阿童木登场", "Wrong"),
            ("“阿童木登场”", "‘Astro Boy appears’"),
        ])
        self.assertEqual(summary.problem_count, 2)
        self.assertEqual([row[0] for row in issues], [4, 5])
        self.assertEqual(issues[0][1:3], ("【 阿童木登场 】", "‘Wrong’"))
        self.assertIn("参考第 2、3", issues[0][3])
        self.assertEqual(issues[0][3], "“【阿童木】” → “‘Astro Boy’”（参考第 2、3 行）")
        self.assertEqual(issues[0][3], issues[1][3])

    def test_normalized_equal_sources_are_not_containment_and_conflicts_still_skip(self):
        summary, issues, _ = self.check_rows([
            ("【阿童木】", "Astro Boy"), ("阿童木", "Astro Boy"),
            ("“悟空”", "Goku"), ("悟空", "Monkey King"), ("悟空登场", "Wrong"),
            ("【 】", "Ignored"), ("阿童木来了", "‘ ’"),
        ])
        self.assertEqual(summary.skipped_conflicting_sources, 1)
        self.assertEqual(issues, [])

    def check_rows(self, rows, **kwargs):
        workbook = Workbook()
        self.addCleanup(workbook.close)
        worksheet = workbook.active
        worksheet.title = "Data"
        worksheet.append(["source", "target"])
        for row in rows:
            worksheet.append(row)
        worksheet["Z10000"] = "unrelated tail"
        summary = process_workbook(
            workbook=workbook, output_path=Path("unused.xlsx"),
            source_column="a", target_column="b", sheet="Data", **kwargs,
        )
        return summary, list(workbook[PROBLEM_SHEET_NAME].values)[1:], workbook

    def test_reports_parent_with_concise_mapping_and_reference_rows(self):
        summary, issues, workbook = self.check_rows([
            ("保存更改", "Save changes"),
            ("保存更改", "Save changes"),
            ("是否保存更改？", "Save modifications?"),
            ("请保存更改。", "Please SAVE CHANGES."),
        ])
        self.assertEqual(summary.total_rows_checked, 4)
        self.assertEqual(summary.problem_count, 1)
        self.assertEqual(summary.problem_rows, 1)
        self.assertEqual(issues[0][:3], (4, "是否保存更改？", "Save modifications?"))
        self.assertIn("参考第 2、3", issues[0][3])
        self.assertEqual(issues[0][3], "“保存更改” → “Save changes”（参考第 2、3 行）")
        self.assertEqual(workbook[PROBLEM_SHEET_NAME]["A2"].hyperlink.location, "'Data'!B4")
        self.assertEqual(workbook["Data"]["B4"].value, "Save modifications?")

    def test_source_is_exact_and_both_sides_use_ascii_word_boundaries(self):
        _, issues, _ = self.check_rows([
            ("he", "il"), ("the hero", "wrong"),
            ("Save", "save"), ("Save now", "saver"),
            ("save later", "wrong"), ("Save later", "Please save now"),
            ("New\nGame", "Play"), ("New Game now", "wrong"),
        ])
        self.assertEqual([issue[0] for issue in issues], [5])

    def test_empty_conflicting_and_structural_references_are_skipped(self):
        summary, issues, _ = self.check_rows([
            ("重复短句", "One"), ("重复短句", "Two"), ("包含重复短句", "wrong"),
            ("未完短句", "Ready"), ("未完短句", None), ("包含未完短句", "wrong"),
            ("空白短句", "  "), ("包含空白短句", "wrong"),
            ("保存更改", "Save"), ("请保存更改", None),
            (None, "ignored"), ("  ", "ignored"),
            ("{player}", "{player}"), ("Hello {player}", "wrong"),
            ("<br/>", "Break"), ("Hello <br/>", "wrong"),
            ("%s", "String"), ("Hello %s", "wrong"),
            ("123", "Number"), ("123 apples", "wrong"),
            ("的", "of"), ("我的", "wrong"),
            ("!!", "Alert"), ("Oh!!", "wrong"),
        ])
        self.assertEqual(summary.skipped_conflicting_sources, 2)
        self.assertEqual(issues, [])

    def test_duplicate_occurrences_deduplicate_but_all_ancestors_are_checked(self):
        summary, issues, _ = self.check_rows([
            ("保存", "Save"), ("保存更改", "Save changes"),
            ("请保存更改并保存", "Wrong"),
            ("请保存更改并保存", "Wrong"),
        ], min_cjk_chars=2)
        self.assertEqual(summary.problem_count, 4)
        self.assertEqual([issue[0] for issue in issues], [4, 5])
        self.assertEqual(len(issues[0][3].splitlines()), 2)
        self.assertIn("参考第 2", issues[0][3])
        self.assertIn("参考第 3", issues[0][3])

    def test_details_are_bounded_but_count_is_complete(self):
        children = [f"测试词条{index:03}" for index in range(DETAIL_LIMIT + 5)]
        summary, issues, _ = self.check_rows([
            *((child, f"Translation {index}") for index, child in enumerate(children)),
            (" ".join(children), "wrong"),
        ])
        self.assertEqual(summary.problem_count, len(children))
        self.assertEqual(len(issues), 1)
        self.assertEqual(issues[0][3].count("参考第"), DETAIL_LIMIT)
        self.assertIn("另 5 条未展示", issues[0][3])

    def test_long_reference_excerpts_and_duplicate_row_lists_are_explicitly_truncated(self):
        summary, issues, _ = self.check_rows([
            *(("保存" * 100, "Save " * 100) for _ in range(12)),
            ("请" + "保存" * 100, "wrong"),
        ])
        self.assertEqual(summary.problem_count, 1)
        self.assertIn("共 12 行，行号已截断", issues[0][3])
        self.assertIn("…（已截断）", issues[0][3])
        self.assertLess(len(issues[0][3]), 1000)

    def test_empty_index_single_character_target_and_start_row(self):
        self.assertEqual(self.check_rows([])[1], [])
        self.assertEqual(self.check_rows([("单字", "字"), ("一个单字", "汉字")])[1], [])
        self.assertEqual(self.check_rows([
            ("保存", "Save"), ("请保存", "wrong"),
        ], start_row=3)[1], [])
        with self.assertRaisesRegex(ValueError, "开始行"):
            self.check_rows([], start_row=0)

    def test_matches_brute_force_oracle_on_generated_corpus(self):
        rng = random.Random(42)
        children = [f"测试词条{index:03}" for index in range(30)]
        rows = [(child, f"Term{index:03}") for index, child in enumerate(children)]
        for index in range(80):
            chosen = rng.sample(range(30), rng.randrange(1, 8))
            source = " ".join(children[i] for i in chosen) + f" 结尾{index}"
            target = " ".join(f"Term{i:03}" for i in chosen if rng.random() > 0.5) or "wrong"
            rows.append((source, target))
        expected = {}
        for row, (source, target) in enumerate(rows, 2):
            count = sum(
                len(child) < len(source)
                and text_contains_term(source, child, "hybrid-boundary")
                and not text_contains_term(target.casefold(), translation.casefold(), "hybrid-boundary")
                for child, translation in rows
            )
            if count:
                expected[row] = count
        summary, issues, _ = self.check_rows(rows)
        self.assertEqual(summary.problem_count, sum(expected.values()))
        self.assertEqual([issue[0] for issue in issues], list(expected))
        for issue in issues:
            self.assertEqual(len(issue[3].splitlines()), expected[issue[0]])


if __name__ == "__main__":
    unittest.main()
