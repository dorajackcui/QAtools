from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from tools.workflow.workflow_runner import (
    WORKFLOW_SUMMARY_SHEET_NAME,
    WORKFLOW_TERM_PROBLEM_SHEET_NAME,
    count_unique_problem_rows,
    run_workflow,
)


class WorkflowRunnerTests(unittest.TestCase):
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
        workbook.save(path)

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
            self.assertTrue(summary.ran_chinese_target_check)
            self.assertEqual(summary.term_problem_count, 1)
            self.assertEqual(summary.term_problem_rows, 1)
            self.assertEqual(summary.tag_problem_count, 1)
            self.assertEqual(summary.tag_problem_rows, 1)
            self.assertEqual(summary.line_break_problem_count, 1)
            self.assertEqual(summary.source_consistency_problem_count, 1)
            self.assertEqual(summary.source_consistency_problem_rows, 2)
            self.assertEqual(summary.chinese_target_problem_count, 4)

            workbook = load_workbook(summary.output_path)
            self.assertEqual(
                workbook.sheetnames,
                [
                    "Data",
                    "术语表",
                    WORKFLOW_TERM_PROBLEM_SHEET_NAME,
                    "标签占位问题",
                    "换行数量问题",
                    "同源译文不一致",
                    WORKFLOW_SUMMARY_SHEET_NAME,
                ],
            )
            self.assertNotIn("问题列", workbook.sheetnames)
            self.assertNotIn("检查汇总", workbook.sheetnames)
            self.assertEqual(workbook[WORKFLOW_TERM_PROBLEM_SHEET_NAME]["A2"].value, 3)
            self.assertEqual(workbook["标签占位问题"]["A2"].value, 3)
            self.assertEqual(workbook["换行数量问题"]["A2"].value, 3)
            self.assertEqual(workbook["同源译文不一致"]["C2"].value, 4)
            self.assertEqual(workbook["同源译文不一致"]["C3"].value, 5)
            self.assertEqual(workbook["Data"]["C1"].value, "中文检查")
            summary_sheet = workbook[WORKFLOW_SUMMARY_SHEET_NAME]
            self.assertEqual(
                list(summary_sheet.values),
                [
                    ("检查项", "问题行数"),
                    ("术语检查", 1),
                    ("Tag 检查", 1),
                    ("换行数量检查", 1),
                    ("同源译文一致性", 2),
                    ("Target 中文检查", 4),
                ],
            )

    def test_count_unique_problem_rows_deduplicates_source_rows(self) -> None:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.append(["问题行号"])
        worksheet.append([3])
        worksheet.append([3])
        worksheet.append([5])

        self.assertEqual(count_unique_problem_rows(worksheet), 2)

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
                    run_chinese_target_check=False,
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
            )

            self.assertEqual(summary.term_count, 1)
            self.assertEqual(summary.term_problem_count, 1)
            self.assertEqual(summary.term_problem_rows, 1)

if __name__ == "__main__":
    unittest.main()
