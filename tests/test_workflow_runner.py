from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from openpyxl import Workbook, load_workbook

from tools.chinese_target_checker.check_chinese_target import (
    PROBLEM_SHEET_NAME as CHINESE_PROBLEM_SHEET_NAME,
)
from tools.target_text_checker.check_target_text import (
    ABNORMAL_PUNCTUATION_RULE,
    PROBLEM_SHEET_NAME as TARGET_TEXT_PROBLEM_SHEET_NAME,
)
from tools.workflow.workflow_runner import (
    WORKFLOW_SUMMARY_SHEET_NAME,
    WORKFLOW_TERM_PROBLEM_SHEET_NAME,
    count_unique_problem_rows,
    run_workflow,
)
from tools.workflow.review_sheet import (
    WORKFLOW_METADATA_SHEET_NAME,
    WORKFLOW_REVIEW_SHEET_NAME,
    collect_review_rows,
    read_review_metadata,
)
from tools.workflow.revision_applier import (
    apply_workflow_revisions,
    build_default_revised_output_path,
)


class WorkflowRunnerTests(unittest.TestCase):
    def test_normalized_checks_merge_raw_rows_and_allow_revision_roundtrip(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            workbook = Workbook()
            source_rows = [
                ("【阿童木】", "“Astro Boy”"), ("阿童木", "“Other”"),
                ("悟空", "【Other】"), ("【保存】", "Save"), ("【请保存】", "Wrong"),
            ]
            try:
                workbook.active.title = "Data"
                workbook.active.append(["source", "target"])
                for row in source_rows:
                    workbook.active.append(row)
                workbook.save(input_path)
            finally:
                workbook.close()
            summary = run_workflow(
                input_file=input_path, source_column="A", target_column="B",
                run_term_pair_check=False, run_tag_check=False, run_line_break_check=False,
                run_number_check=False, run_url_check=False, run_chinese_target_check=False,
                run_target_text_check=False, run_source_consistency_check=True,
                run_target_consistency_check=True, run_substring_consistency_check=True,
                substring_min_cjk_chars=2,
            )
            self.assertEqual(summary.source_consistency_problem_rows, 2)
            self.assertEqual(summary.target_consistency_problem_rows, 2)
            self.assertEqual(summary.substring_consistency_problem_rows, 1)
            report = load_workbook(summary.output_path)
            try:
                review = report["问题处理"]
                rows = [row for row in review.iter_rows(min_row=2, max_col=6, values_only=True)
                        if isinstance(row[0], int)]
                self.assertEqual([row[0] for row in rows], [2, 3, 4, 6])
                for row in rows:
                    self.assertEqual(row[1:3], source_rows[row[0] - 2])
                self.assertIn("同 Source 不同 Target", rows[1][5])
                self.assertIn("同 Target 不同 Source", rows[1][5])
                self.assertEqual(rows[-1][4], "【子串译文一致性】“【保存】” → “Save”（参考原表第 5 行）")
                review["D2"] = "Astro"
                review["D5"] = "Please save"
                report.save(summary.output_path)
            finally:
                report.close()
            revision = apply_workflow_revisions(summary.output_path)
            self.assertEqual(revision.revised_count, 2)
            self.assertEqual(revision.conflict_rows, ())
            revised = load_workbook(revision.output_path)
            try:
                self.assertEqual(revised["Data"]["A2"].value, "【阿童木】")
                self.assertEqual(revised["Data"]["B2"].value, "Astro")
                self.assertEqual(revised["Data"]["B6"].value, "Please save")
            finally:
                revised.close()

    def test_optional_tag_order_check_reaches_the_unified_report(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            workbook = Workbook()
            try:
                workbook.active.append(["source", "target"])
                workbook.active.append(["<br/>{name}", "{name}<br/>"])
                workbook.save(input_path)
            finally:
                workbook.close()
            for options, expected_rows in (({}, 0), ({"tag_check_order": True}, 1)):
                with self.subTest(options=options):
                    summary = run_workflow(
                        input_file=input_path, source_column="A", target_column="B",
                        run_term_pair_check=False, run_line_break_check=False,
                        run_source_consistency_check=False, run_number_check=False,
                        run_url_check=False, run_chinese_target_check=False,
                        run_target_text_check=False, **options,
                    )
                    self.assertEqual(summary.tag_problem_rows, expected_rows)
                    self.assertEqual(summary.tag_problem_count, expected_rows)
                    result = load_workbook(summary.output_path)
                    try:
                        review = result[WORKFLOW_REVIEW_SHEET_NAME]
                        issues = [
                            row for row in review.iter_rows(min_row=2, max_col=6, values_only=True)
                            if isinstance(row[0], int)
                        ]
                        self.assertEqual(len(issues), expected_rows)
                        if expected_rows:
                            self.assertIn("【Tag 检查】原文：", review["E2"].value)
                            self.assertEqual(review["F2"].value, "Tag 检查")
                    finally:
                        result.close()

    def test_substring_only_report_supports_revision_roundtrip(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            workbook = Workbook()
            try:
                workbook.active.title = "Data"
                workbook.active.append(["source", "target"])
                workbook.active.append(["保存更改", "Save changes"])
                workbook.active.append(["是否保存更改？", "Save modifications?"])
                workbook.save(input_path)
            finally:
                workbook.close()
            summary = run_workflow(
                input_file=input_path, source_column="A", target_column="B",
                run_term_pair_check=False, run_tag_check=False,
                run_line_break_check=False, run_source_consistency_check=False,
                run_number_check=False, run_url_check=False,
                run_chinese_target_check=False, run_target_text_check=False,
                run_substring_consistency_check=True,
            )
            self.assertTrue(summary.ran_substring_consistency_check)
            self.assertEqual(summary.substring_consistency_problem_rows, 1)
            self.assertEqual(summary.substring_consistency_problem_count, 1)
            report = load_workbook(summary.output_path)
            try:
                self.assertEqual(report.sheetnames, [
                    "Data", "问题处理", WORKFLOW_METADATA_SHEET_NAME, "质量检查汇总",
                ])
                review = report["问题处理"]
                self.assertEqual(review["A2"].value, 3)
                self.assertIn("参考原表第 2 行", review["E2"].value)
                self.assertEqual(review["F2"].value, "子串译文一致性")
                self.assertEqual(list(report["质量检查汇总"].values), [
                    ("检查项", "问题行数"), ("子串译文一致性", 1),
                ])
                review["D2"] = "Save changes?"
                report.save(summary.output_path)
            finally:
                report.close()
            result = apply_workflow_revisions(summary.output_path)
            revised = load_workbook(result.output_path)
            try:
                self.assertEqual(revised["Data"]["B3"].value, "Save changes?")
                self.assertEqual(revised["Data"]["B2"].value, "Save changes")
                self.assertEqual(revised.sheetnames, ["Data"])
            finally:
                revised.close()

    def create_workbook(self, path: Path) -> None:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Data"
        worksheet["A1"] = "source"
        worksheet["B1"] = "target"
        worksheet["A2"] = "第一行 [Alpha] 和 <color=red>{name}"
        worksheet["B2"] = "第一行 [阿尔法] 和 <color=red>{name}"
        worksheet["A3"] = "第二行\n复用 [Alpha] 和 <color=red>{name}"
        worksheet["B3"] = "第二行复用 [错误阿尔法] 和 {name}"
        worksheet["A4"] = "Same source"
        worksheet["B4"] = "译文一"
        worksheet["A5"] = "Same source"
        worksheet["B5"] = "译文二"
        worksheet["C1"] = "note"
        worksheet["C3"] = "keep me"
        worksheet["Z1000"] = "unrelated tail"
        worksheet["A1001"].number_format = "@"
        workbook.save(path)
        workbook.close()

    def test_run_workflow_writes_all_quality_check_results_into_same_output_file(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            output_path = Path(tmp_dir) / "workflow_output.xlsx"
            self.create_workbook(input_path)

            summary = run_workflow(
                input_file=input_path,
                output_file=output_path,
                source_column="A",
                target_column="B",
                sheet="Data",
                start_row=2,
                run_term_pair_check=True,
                term_mark_styles=("[]",),
                run_tag_check=True,
                tag_token_types=("angle", "brace"),
                run_line_break_check=True,
                run_source_consistency_check=True,
                run_chinese_target_check=True,
            )

            self.assertEqual(summary.output_path, output_path.resolve())
            self.assertTrue(summary.ran_term_pair_check)
            self.assertTrue(summary.ran_tag_check)
            self.assertTrue(summary.ran_line_break_check)
            self.assertTrue(summary.ran_source_consistency_check)
            self.assertFalse(summary.ran_target_consistency_check)
            self.assertTrue(summary.ran_number_check)
            self.assertTrue(summary.ran_url_check)
            self.assertTrue(summary.ran_chinese_target_check)
            self.assertTrue(summary.ran_target_text_check)
            self.assertEqual(summary.term_problem_count, 1)
            self.assertEqual(summary.term_problem_rows, 1)
            self.assertEqual(summary.tag_problem_count, 1)
            self.assertEqual(summary.tag_problem_rows, 1)
            self.assertEqual(summary.line_break_problem_count, 1)
            self.assertEqual(summary.source_consistency_problem_count, 1)
            self.assertEqual(summary.source_consistency_problem_rows, 2)
            self.assertEqual(summary.chinese_target_problem_count, 4)
            self.assertEqual(summary.target_text_problem_count, 0)
            self.assertEqual(summary.target_text_problem_rows, 0)

            workbook = load_workbook(summary.output_path)
            self.assertEqual(
                workbook.sheetnames,
                [
                    "Data",
                    "术语表",
                    WORKFLOW_REVIEW_SHEET_NAME,
                    WORKFLOW_METADATA_SHEET_NAME,
                    WORKFLOW_SUMMARY_SHEET_NAME,
                ],
            )
            self.assertNotIn("问题列", workbook.sheetnames)
            self.assertNotIn("检查汇总", workbook.sheetnames)
            self.assertNotIn(WORKFLOW_TERM_PROBLEM_SHEET_NAME, workbook.sheetnames)
            self.assertNotIn("标签占位问题", workbook.sheetnames)
            self.assertNotIn("换行数量问题", workbook.sheetnames)
            self.assertNotIn("同源译文不一致", workbook.sheetnames)
            self.assertNotIn(CHINESE_PROBLEM_SHEET_NAME, workbook.sheetnames)
            self.assertNotIn(TARGET_TEXT_PROBLEM_SHEET_NAME, workbook.sheetnames)
            self.assertEqual(workbook["Data"]["C1"].value, "note")
            self.assertEqual(workbook["Data"]["C3"].value, "keep me")
            summary_sheet = workbook[WORKFLOW_SUMMARY_SHEET_NAME]
            self.assertEqual(
                list(summary_sheet.values),
                [
                    ("检查项", "问题行数"),
                    ("术语检查", 1),
                    ("同 Source 不同 Target", 2),
                    ("Tag 检查", 1),
                    ("换行数量检查", 1),
                    ("数字一致性", 0),
                    ("URL 一致性", 0),
                    ("Target 中文检查", 4),
                    ("Target 文本规范检查", 0),
                ],
            )
            review_sheet = workbook[WORKFLOW_REVIEW_SHEET_NAME]
            self.assertEqual(
                [review_sheet.cell(1, column).value for column in range(1, 7)],
                [
                    "行号",
                    "source",
                    "target",
                    "修改后target",
                    "问题描述",
                    "检查项",
                ],
            )
            self.assertEqual(
                [review_sheet.cell(row, 1).value for row in range(2, 6)],
                [4, 5, 2, 3],
            )
            merged_check_items = [
                item
                for row in range(2, 6)
                for item in review_sheet.cell(row, 6).value.split("；")
            ]
            self.assertEqual(merged_check_items.count("术语检查"), 1)
            self.assertEqual(merged_check_items.count("Tag 检查"), 1)
            self.assertEqual(merged_check_items.count("换行数量检查"), 1)
            self.assertEqual(merged_check_items.count("同 Source 不同 Target"), 2)
            self.assertEqual(merged_check_items.count("Target 中文检查"), 4)
            self.assertEqual(review_sheet["A5"].value, 3)
            self.assertIsNone(review_sheet["D5"].value)
            self.assertIn("术语检查", review_sheet["F5"].value)
            self.assertIn("Tag 检查", review_sheet["F5"].value)
            self.assertIn("换行数量检查", review_sheet["F5"].value)
            self.assertIn("Target 中文检查", review_sheet["F5"].value)
            row_three_description = review_sheet["E5"].value
            self.assertIn("【术语检查】“Alpha” → “阿尔法”（当前：“错误阿尔法”；本批次新增）", row_three_description)
            self.assertNotIn("source术语：", row_three_description)
            self.assertNotIn("预期target术语：", row_three_description)
            self.assertNotIn("术语来源：", row_three_description)
            self.assertIn("【Tag 检查】缺少：<color=red>", row_three_description)
            self.assertIn("【换行数量检查】原文 1 个，译文 0 个", row_three_description)
            self.assertIn("【Target 中文检查】命中：", row_three_description)
            for redundant in ("问题类型：", "source换行数：", "target换行数：", "数量差：", "命中字符："):
                self.assertNotIn(redundant, row_three_description)
            row_four_description = review_sheet["E2"].value
            self.assertIn("【同 Source 不同 Target】2 种译法：1: 译文一；2: 译文二", row_four_description)
            self.assertNotIn("target版本数：", row_four_description)
            self.assertNotIn("同组行号：", row_four_description)
            for review_row, source_row in enumerate((4, 5, 2, 3), start=2):
                row_cell = review_sheet.cell(review_row, 1)
                self.assertEqual(row_cell.hyperlink.location, f"'Data'!B{source_row}")
                self.assertIsNone(row_cell.hyperlink.target)
            self.assertEqual(review_sheet.max_column, 6)
            self.assertEqual(workbook[WORKFLOW_METADATA_SHEET_NAME].sheet_state, "veryHidden")
            self.assertEqual(len(review_sheet.data_validations.dataValidation), 0)
            metadata = read_review_metadata(workbook)
            self.assertEqual(metadata["data_sheet_name"], "Data")
            self.assertEqual(metadata["source_column"], "A")
            self.assertEqual(metadata["target_column"], "B")
            self.assertNotIn("remove_term_helper", metadata)
            workbook.close()

    def test_workflow_loads_and_saves_the_main_workbook_once(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            output_path = Path(tmp_dir) / "workflow_output.xlsx"
            self.create_workbook(input_path)
            shared_workbook = load_workbook(input_path)
            real_save = shared_workbook.save
            shared_workbook.save = Mock(wraps=real_save)

            with patch(
                "tools.workflow.workflow_runner.load_workbook_for_editing",
                return_value=shared_workbook,
            ) as load_mock:
                run_workflow(
                    input_file=input_path,
                    output_file=output_path,
                    source_column="A",
                    target_column="B",
                    sheet="Data",
                    run_term_pair_check=True,
                    term_mark_styles=("[]",),
                    run_tag_check=True,
                    tag_token_types=("angle", "brace"),
                    run_line_break_check=True,
                    run_source_consistency_check=True,
                    run_target_consistency_check=True,
                    run_substring_consistency_check=True,
                    run_number_check=True,
                    run_url_check=True,
                    run_chinese_target_check=True,
                    run_target_text_check=True,
                )

            load_mock.assert_called_once_with(input_path.resolve())
            self.assertEqual(shared_workbook.save.call_count, 1)
            self.assertEqual(shared_workbook.save.call_args.args, (output_path.resolve(),))

    def test_workflow_closes_the_main_workbook_when_a_check_fails(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            output_path = Path(tmp_dir) / "workflow_output.xlsx"
            self.create_workbook(input_path)
            shared_workbook = load_workbook(input_path)
            real_close = shared_workbook.close
            shared_workbook.close = Mock(wraps=real_close)

            with (
                patch(
                    "tools.workflow.workflow_runner.load_workbook_for_editing",
                    return_value=shared_workbook,
                ),
                patch(
                    "tools.workflow.workflow_runner.run_tag_check_workbook",
                    side_effect=RuntimeError("check failed"),
                ),
                self.assertRaisesRegex(RuntimeError, "check failed"),
            ):
                run_workflow(
                    input_file=input_path,
                    output_file=output_path,
                    source_column="A",
                    target_column="B",
                    sheet="Data",
                    run_term_pair_check=False,
                    run_tag_check=True,
                    run_line_break_check=False,
                    run_source_consistency_check=False,
                    run_target_consistency_check=False,
                    run_number_check=False,
                    run_url_check=False,
                    run_chinese_target_check=False,
                    run_target_text_check=False,
                )

            shared_workbook.close.assert_called_once_with()
            self.assertFalse(output_path.exists())

    def test_run_workflow_surfaces_angle_tag_structure_mismatches(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "Data"
            worksheet.append(["source", "target"])
            worksheet.append(["<b>A</b><i>B</i>", "<b>A<i>B</i></b>"])
            workbook.save(input_path)

            summary = run_workflow(
                input_file=input_path,
                source_column="A",
                target_column="B",
                sheet="Data",
                run_term_pair_check=False,
                run_tag_check=True,
                tag_token_types=("angle",),
                run_line_break_check=False,
                run_source_consistency_check=False,
                run_chinese_target_check=False,
                run_target_text_check=False,
            )

            self.assertEqual(summary.tag_problem_count, 1)
            self.assertEqual(summary.tag_problem_rows, 1)
            output_workbook = load_workbook(summary.output_path)
            review_sheet = output_workbook[WORKFLOW_REVIEW_SHEET_NAME]
            self.assertEqual(review_sheet["A2"].value, 2)
            self.assertIn("【Tag 检查】嵌套或闭合结构不同", review_sheet["E2"].value)
            self.assertEqual(review_sheet["F2"].value, "Tag 检查")
            output_workbook.close()

    def test_run_workflow_integrates_reverse_number_and_url_checks(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "Data"
            worksheet.append(["source", "target"])
            worksheet.append(
                [
                    "Open 10 files at https://old.example/path",
                    "打开 11 个文件：https://new.example/path",
                ]
            )
            worksheet.append(["Save", "保存"])
            worksheet.append(["Store", "保存"])
            workbook.save(input_path)
            workbook.close()

            summary = run_workflow(
                input_file=input_path,
                source_column="A",
                target_column="B",
                sheet="Data",
                run_term_pair_check=False,
                run_tag_check=False,
                run_line_break_check=False,
                run_source_consistency_check=False,
                run_target_consistency_check=True,
                run_number_check=True,
                run_url_check=True,
                run_chinese_target_check=False,
                run_target_text_check=False,
            )

            self.assertEqual(summary.target_consistency_problem_count, 1)
            self.assertEqual(summary.target_consistency_problem_rows, 2)
            self.assertEqual(summary.number_problem_rows, 1)
            self.assertEqual(summary.url_problem_rows, 1)
            output_workbook = load_workbook(summary.output_path)
            try:
                review_sheet = output_workbook[WORKFLOW_REVIEW_SHEET_NAME]
                self.assertEqual(
                    [review_sheet[f"A{row}"].value for row in range(2, 5)],
                    [3, 4, 2],
                )
                self.assertEqual(
                    review_sheet["F4"].value,
                    "数字一致性；URL 一致性",
                )
                self.assertEqual(review_sheet["F2"].value, "同 Target 不同 Source")
                self.assertIn("缺少：10", review_sheet["E4"].value)
                self.assertIn(
                    "多出：https://new.example/path",
                    review_sheet["E4"].value,
                )
                self.assertEqual(
                    list(output_workbook[WORKFLOW_SUMMARY_SHEET_NAME].values),
                    [
                        ("检查项", "问题行数"),
                        ("同 Target 不同 Source", 2),
                        ("数字一致性", 1),
                        ("URL 一致性", 1),
                    ],
                )
            finally:
                output_workbook.close()

    def test_run_workflow_surfaces_empty_angle_and_brace_tokens(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "Data"
            worksheet.append(["source", "target"])
            worksheet.append(["保留 {}、<> 和 < >", "全部缺少"])
            workbook.save(input_path)
            workbook.close()

            summary = run_workflow(
                input_file=input_path,
                source_column="A",
                target_column="B",
                sheet="Data",
                run_term_pair_check=False,
                run_tag_check=True,
                tag_token_types=("angle", "brace"),
                run_line_break_check=False,
                run_source_consistency_check=False,
                run_chinese_target_check=False,
                run_target_text_check=False,
            )

            self.assertEqual(summary.tag_problem_count, 2)
            self.assertEqual(summary.tag_problem_rows, 1)
            output_workbook = load_workbook(summary.output_path)
            review_sheet = output_workbook[WORKFLOW_REVIEW_SHEET_NAME]
            self.assertIn("缺少：< >、<>", review_sheet["E2"].value)
            self.assertIn("缺少：{}", review_sheet["E2"].value)
            self.assertEqual(review_sheet["F2"].value, "Tag 检查")
            output_workbook.close()

    def test_run_workflow_passes_the_angle_filter_config(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            config_path = Path(tmp_dir) / "angle-tags.json"
            config_path.write_text('{"patterns": ["^/?b$"]}', encoding="utf-8")
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "Data"
            worksheet.append(["source", "target"])
            worksheet.append(["<i>Italic</i>", "斜体"])
            workbook.save(input_path)

            summary = run_workflow(
                input_file=input_path,
                source_column="A",
                target_column="B",
                sheet="Data",
                run_term_pair_check=False,
                run_tag_check=True,
                tag_token_types=("angle",),
                tag_angle_config_file=config_path,
                run_line_break_check=False,
                run_source_consistency_check=False,
                run_chinese_target_check=False,
                run_target_text_check=False,
            )

            self.assertEqual(summary.tag_problem_count, 0)
            self.assertEqual(summary.tag_problem_rows, 0)

    def test_apply_workflow_revisions_writes_targets_and_removes_qa_artifacts(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            report_path = Path(tmp_dir) / "workflow_output.xlsx"
            revised_path = Path(tmp_dir) / "revised_input.xlsx"
            self.create_workbook(input_path)
            run_workflow(
                input_file=input_path,
                output_file=report_path,
                source_column="A",
                target_column="B",
                sheet="Data",
                start_row=2,
                term_mark_styles=("[]",),
                tag_token_types=("angle", "brace"),
            )

            report_workbook = load_workbook(report_path)
            review_sheet = report_workbook[WORKFLOW_REVIEW_SHEET_NAME]
            self.assertEqual(review_sheet["A2"].value, 4)
            self.assertEqual(review_sheet["A5"].value, 3)
            review_sheet["D2"] = "第四行修订"
            review_sheet["D5"] = "第三行修订"
            review_sheet["Z1000"] = "unrelated tail"
            review_sheet["A1001"].number_format = "@"
            report_workbook.save(report_path)
            report_workbook.close()

            summary = apply_workflow_revisions(report_path, output_file=revised_path)

            self.assertEqual(summary.revised_count, 2)
            self.assertEqual(summary.ignored_count, 2)
            self.assertEqual(summary.unchanged_count, 0)
            self.assertEqual(summary.conflict_rows, ())
            revised_workbook = load_workbook(revised_path)
            self.assertEqual(revised_workbook.sheetnames, ["Data"])
            data_sheet = revised_workbook["Data"]
            self.assertEqual(data_sheet["B2"].value, "第一行 [阿尔法] 和 <color=red>{name}")
            self.assertEqual(data_sheet["B3"].value, "第三行修订")
            self.assertEqual(data_sheet["B4"].value, "第四行修订")
            self.assertEqual(data_sheet["B5"].value, "译文二")
            self.assertEqual(data_sheet["C1"].value, "note")
            self.assertEqual(data_sheet["C3"].value, "keep me")
            revised_workbook.close()

    def test_apply_workflow_revisions_ignores_empty_string_targets(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            report_path = Path(tmp_dir) / "workflow_output.xlsx"
            revised_path = Path(tmp_dir) / "revised_input.xlsx"
            self.create_workbook(input_path)
            run_workflow(
                input_file=input_path,
                output_file=report_path,
                source_column="A",
                target_column="B",
                sheet="Data",
                start_row=2,
                term_mark_styles=("[]",),
                tag_token_types=("angle", "brace"),
            )

            report_workbook = load_workbook(report_path)
            review_sheet = report_workbook[WORKFLOW_REVIEW_SHEET_NAME]
            review_sheet["D2"] = ""
            review_sheet["D3"] = " \t"
            review_sheet["D4"] = "第一行修订"

            with patch(
                "tools.workflow.revision_applier.load_workbook_for_editing",
                return_value=report_workbook,
            ):
                summary = apply_workflow_revisions(
                    report_path,
                    output_file=revised_path,
                )
            report_workbook.close()

            self.assertEqual(summary.revised_count, 1)
            self.assertEqual(summary.ignored_count, 3)
            self.assertEqual(summary.unchanged_count, 0)
            self.assertEqual(summary.conflict_rows, ())
            revised_workbook = load_workbook(revised_path)
            data_sheet = revised_workbook["Data"]
            self.assertEqual(data_sheet["B2"].value, "第一行修订")
            self.assertEqual(data_sheet["B3"].value, "第二行复用 [错误阿尔法] 和 {name}")
            self.assertEqual(data_sheet["B4"].value, "译文一")
            self.assertEqual(data_sheet["B5"].value, "译文二")
            revised_workbook.close()

    def test_apply_workflow_revisions_overwrites_changed_target_when_source_matches(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            report_path = Path(tmp_dir) / "workflow_output.xlsx"
            revised_path = Path(tmp_dir) / "revised_input.xlsx"
            self.create_workbook(input_path)
            run_workflow(
                input_file=input_path,
                output_file=report_path,
                source_column="A",
                target_column="B",
                sheet="Data",
                start_row=2,
                term_mark_styles=("[]",),
                tag_token_types=("angle", "brace"),
            )

            report_workbook = load_workbook(report_path)
            report_workbook["Data"]["B2"] = "人工直接修改"
            report_workbook[WORKFLOW_REVIEW_SHEET_NAME]["D4"] = "准备回填的修改"
            report_workbook.save(report_path)
            report_workbook.close()

            summary = apply_workflow_revisions(report_path, output_file=revised_path)

            self.assertEqual(summary.revised_count, 1)
            self.assertEqual(summary.conflict_rows, ())
            revised_workbook = load_workbook(revised_path)
            self.assertEqual(revised_workbook["Data"]["B2"].value, "准备回填的修改")
            revised_workbook.close()

    def test_apply_workflow_revisions_ignores_edited_review_targets(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            report_path = Path(tmp_dir) / "workflow_output.xlsx"
            revised_path = Path(tmp_dir) / "revised_input.xlsx"
            self.create_workbook(input_path)
            run_workflow(
                input_file=input_path,
                output_file=report_path,
                source_column="A",
                target_column="B",
                sheet="Data",
                term_mark_styles=("[]",),
                tag_token_types=("angle", "brace"),
            )

            report_workbook = load_workbook(report_path)
            try:
                review_sheet = report_workbook[WORKFLOW_REVIEW_SHEET_NAME]
                self.assertEqual(
                    [review_sheet.cell(row, 1).value for row in range(2, 6)],
                    [4, 5, 2, 3],
                )
                review_sheet["C2"] = "误改原 target"
                review_sheet["D2"] = "第四行修订"
                review_sheet["C3"] = None
                review_sheet["D3"] = "第五行修订"
                review_sheet["C4"] = "只修改参考值，不回填"
                review_sheet["C5"] = "误改参考值，但最终译文未变化"
                review_sheet["D5"] = report_workbook["Data"]["B3"].value
                report_workbook.save(report_path)
            finally:
                report_workbook.close()

            summary = apply_workflow_revisions(report_path, output_file=revised_path)

            self.assertEqual(summary.revised_count, 2)
            self.assertEqual(summary.ignored_count, 1)
            self.assertEqual(summary.unchanged_count, 1)
            self.assertEqual(summary.conflict_rows, ())
            revised_workbook = load_workbook(revised_path)
            try:
                data_sheet = revised_workbook["Data"]
                self.assertEqual(data_sheet["B2"].value, "第一行 [阿尔法] 和 <color=red>{name}")
                self.assertEqual(data_sheet["B3"].value, "第二行复用 [错误阿尔法] 和 {name}")
                self.assertEqual(data_sheet["B4"].value, "第四行修订")
                self.assertEqual(data_sheet["B5"].value, "第五行修订")
                self.assertEqual(data_sheet["A4"].value, "Same source")
            finally:
                revised_workbook.close()

    def test_apply_workflow_revisions_skips_mismatched_sources(self) -> None:
        for changed_field, conflict_row in (("data_source", 2), ("review_source", 2), ("row_number", 3)):
            with self.subTest(changed_field=changed_field), tempfile.TemporaryDirectory() as tmp_dir:
                input_path = Path(tmp_dir) / "input.xlsx"
                report_path = Path(tmp_dir) / "workflow_output.xlsx"
                revised_path = Path(tmp_dir) / "revised_input.xlsx"
                self.create_workbook(input_path)
                run_workflow(
                    input_file=input_path,
                    output_file=report_path,
                    source_column="A",
                    target_column="B",
                    sheet="Data",
                    term_mark_styles=("[]",),
                    tag_token_types=("angle", "brace"),
                )

                report_workbook = load_workbook(report_path)
                try:
                    review_sheet = report_workbook[WORKFLOW_REVIEW_SHEET_NAME]
                    self.assertEqual(review_sheet["A4"].value, 2)
                    if changed_field == "data_source":
                        report_workbook["Data"]["A2"] = "原数据原文已变化"
                    elif changed_field == "review_source":
                        review_sheet["B4"] = review_sheet["B4"].value + " "
                    else:
                        review_sheet["A4"] = 3
                    review_sheet["D4"] = "不应写入原文不匹配的行"
                    report_workbook.save(report_path)
                finally:
                    report_workbook.close()

                summary = apply_workflow_revisions(report_path, output_file=revised_path)

                self.assertEqual(summary.revised_count, 0)
                self.assertEqual(summary.conflict_rows, (conflict_row,))
                revised_workbook = load_workbook(revised_path)
                try:
                    data_sheet = revised_workbook["Data"]
                    self.assertEqual(data_sheet["B2"].value, "第一行 [阿尔法] 和 <color=red>{name}")
                    self.assertEqual(data_sheet["B3"].value, "第二行复用 [错误阿尔法] 和 {name}")
                    self.assertEqual(data_sheet["B4"].value, "译文一")
                    self.assertEqual(data_sheet["B5"].value, "译文二")
                finally:
                    revised_workbook.close()

    def test_default_revised_output_path_uses_original_input_name(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            report_path = Path(tmp_dir) / "workflow_output.xlsx"
            self.create_workbook(input_path)
            run_workflow(
                input_file=input_path,
                output_file=report_path,
                source_column="A",
                target_column="B",
                sheet="Data",
                start_row=2,
                run_term_pair_check=False,
                run_tag_check=False,
                run_line_break_check=False,
                run_source_consistency_check=False,
                run_chinese_target_check=True,
                run_target_text_check=False,
            )

            self.assertEqual(
                build_default_revised_output_path(report_path),
                (Path(tmp_dir) / "revised_input.xlsx").resolve(),
            )

    def test_apply_revisions_preserves_unrelated_sheets_from_disabled_checks(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            report_path = Path(tmp_dir) / "workflow_output.xlsx"
            revised_path = Path(tmp_dir) / "revised_input.xlsx"
            self.create_workbook(input_path)
            input_workbook = load_workbook(input_path)
            legitimate_term_sheet = input_workbook.create_sheet("术语表")
            legitimate_term_sheet["A1"] = "用户原有内容"
            input_workbook.save(input_path)
            run_workflow(
                input_file=input_path,
                output_file=report_path,
                source_column="A",
                target_column="B",
                sheet="Data",
                start_row=2,
                run_term_pair_check=False,
                run_tag_check=False,
                run_line_break_check=False,
                run_source_consistency_check=False,
                run_chinese_target_check=True,
                run_target_text_check=False,
            )

            report_workbook = load_workbook(report_path)
            report_workbook[WORKFLOW_REVIEW_SHEET_NAME]["D2"] = "第一行修订"
            report_workbook.save(report_path)
            apply_workflow_revisions(report_path, output_file=revised_path)

            revised_workbook = load_workbook(revised_path)
            self.assertIn("术语表", revised_workbook.sheetnames)
            self.assertEqual(revised_workbook["术语表"]["A1"].value, "用户原有内容")

    def test_count_unique_problem_rows_deduplicates_source_rows(self) -> None:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.append(["问题行号"])
        worksheet.append([3])
        worksheet.append([3])
        worksheet.append([5])

        self.assertEqual(count_unique_problem_rows(worksheet), 2)

    def test_review_merge_rejects_incompatible_problem_sheet_schema(self) -> None:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "旧问题表"
        worksheet.append(["问题行号", "source", "target", "描述"])
        worksheet.append([2, "Source", "Target", "Problem"])

        with self.assertRaisesRegex(ValueError, "前四列必须为"):
            collect_review_rows(workbook, (("测试检查", "旧问题表"),))

    def test_run_workflow_requires_at_least_one_task(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            self.create_workbook(input_path)

            with self.assertRaisesRegex(ValueError, "请至少选择一个质量检查项目"):
                run_workflow(
                    input_file=input_path,
                    source_column="A",
                    target_column="B",
                    run_term_pair_check=False,
                    run_tag_check=False,
                    run_line_break_check=False,
                    run_source_consistency_check=False,
                    run_target_consistency_check=False,
                    run_number_check=False,
                    run_url_check=False,
                    run_chinese_target_check=False,
                    run_target_text_check=False,
                )

    def test_run_workflow_passes_history_tb_to_term_pair_check(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            history_path = Path(tmp_dir) / "history.xlsx"
            output_path = Path(tmp_dir) / "workflow_output.xlsx"

            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "Data"
            worksheet["A1"] = "source"
            worksheet["B1"] = "target"
            worksheet["A2"] = "本批次 [Alpha]"
            worksheet["B2"] = "本批次 [临时译法]"
            workbook.save(input_path)

            history_workbook = Workbook()
            history_sheet = history_workbook.active
            history_sheet.title = "TB"
            history_sheet["A1"] = "source"
            history_sheet["B1"] = "target"
            history_sheet["A2"] = "Alpha"
            history_sheet["B2"] = "历史译法"
            history_workbook.save(history_path)

            summary = run_workflow(
                input_file=input_path,
                output_file=output_path,
                source_column="A",
                target_column="B",
                sheet="Data",
                start_row=2,
                run_term_pair_check=True,
                term_mark_styles=("[]",),
                term_history_tb_file=history_path,
                term_history_sheet="TB",
                run_tag_check=False,
                run_line_break_check=False,
                run_source_consistency_check=False,
                run_chinese_target_check=False,
                run_target_text_check=False,
            )

            self.assertEqual(summary.term_count, 1)
            self.assertEqual(summary.term_problem_count, 1)
            self.assertEqual(summary.term_problem_rows, 1)

            workbook = load_workbook(summary.output_path)
            term_sheet = workbook["术语表"]
            self.assertEqual(term_sheet["B2"].value, "历史译法")
            self.assertEqual(term_sheet["E2"].value, "历史TB")

    def test_run_workflow_can_check_history_tb_without_term_marks(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            history_path = Path(tmp_dir) / "history.xlsx"

            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "Data"
            worksheet["A1"] = "source"
            worksheet["B1"] = "target"
            worksheet["A2"] = "Alpha appears here"
            worksheet["B2"] = "这里使用错误译法"
            workbook.save(input_path)

            history_workbook = Workbook()
            history_sheet = history_workbook.active
            history_sheet.title = "TB"
            history_sheet["A1"] = "source"
            history_sheet["B1"] = "target"
            history_sheet["A2"] = "Alpha"
            history_sheet["B2"] = "历史译法"
            history_workbook.save(history_path)

            summary = run_workflow(
                input_file=input_path,
                source_column="A",
                target_column="B",
                sheet="Data",
                start_row=2,
                run_term_pair_check=True,
                term_mark_styles=(),
                term_history_tb_file=history_path,
                term_history_sheet="TB",
                run_tag_check=False,
                run_line_break_check=False,
                run_source_consistency_check=False,
                run_chinese_target_check=False,
                run_target_text_check=False,
            )

            self.assertEqual(summary.term_count, 1)
            self.assertEqual(summary.term_problem_count, 1)
            self.assertEqual(summary.term_problem_rows, 1)

    def test_term_check_does_not_shift_a_source_column_right_of_target(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "Data"
            worksheet["B1"] = "target"
            worksheet["D1"] = "source"
            worksheet["B2"] = "[One]"
            worksheet["D2"] = "[Term]"
            worksheet["B3"] = "Two"
            worksheet["D3"] = "Same"
            worksheet["B4"] = "Three"
            worksheet["D4"] = "Same"
            workbook.save(input_path)
            workbook.close()

            summary = run_workflow(
                input_file=input_path,
                source_column="D",
                target_column="B",
                sheet="Data",
                term_mark_styles=("[]",),
                run_tag_check=False,
                run_line_break_check=False,
                run_source_consistency_check=True,
                run_chinese_target_check=False,
                run_target_text_check=False,
            )

            self.assertEqual(summary.source_consistency_problem_count, 1)
            self.assertEqual(summary.source_consistency_problem_rows, 2)
            output_workbook = load_workbook(summary.output_path)
            self.assertEqual(output_workbook["Data"]["D1"].value, "source")
            review_sheet = output_workbook[WORKFLOW_REVIEW_SHEET_NAME]
            self.assertEqual(review_sheet["A2"].value, 3)
            review_sheet["C2"] = "误改原 target"
            review_sheet["D2"] = "Revised"
            output_workbook.save(summary.output_path)
            output_workbook.close()

            revision = apply_workflow_revisions(summary.output_path)

            self.assertEqual(revision.revised_count, 1)
            self.assertEqual(revision.conflict_rows, ())
            revised_workbook = load_workbook(revision.output_path)
            try:
                self.assertEqual(revised_workbook["Data"]["B3"].value, "Revised")
                self.assertEqual(revised_workbook["Data"]["D3"].value, "Same")
                self.assertEqual(revised_workbook["Data"]["B4"].value, "Three")
            finally:
                revised_workbook.close()

    def test_run_workflow_integrates_selected_target_text_rules(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "Data"
            worksheet.append(["source", "target"])
            worksheet.append(["row 2", "Wait..  now,，"])
            workbook.save(input_path)

            summary = run_workflow(
                input_file=input_path,
                source_column="A",
                target_column="B",
                sheet="Data",
                run_term_pair_check=False,
                run_tag_check=False,
                run_line_break_check=False,
                run_source_consistency_check=False,
                run_number_check=False,
                run_url_check=False,
                run_chinese_target_check=False,
                run_target_text_check=True,
                target_text_rules=(ABNORMAL_PUNCTUATION_RULE,),
            )

            self.assertTrue(summary.ran_target_text_check)
            self.assertEqual(summary.target_text_problem_count, 1)
            self.assertEqual(summary.target_text_problem_rows, 1)
            output_workbook = load_workbook(summary.output_path)
            review_sheet = output_workbook[WORKFLOW_REVIEW_SHEET_NAME]
            self.assertEqual(review_sheet["F2"].value, "Target 文本规范检查")
            self.assertIn("异常标点：", review_sheet["E2"].value)
            self.assertNotIn("连续空格", review_sheet["E2"].value)
            self.assertEqual(
                list(output_workbook[WORKFLOW_SUMMARY_SHEET_NAME].values),
                [
                    ("检查项", "问题行数"),
                    ("Target 文本规范检查", 1),
                ],
            )

if __name__ == "__main__":
    unittest.main()
