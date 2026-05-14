from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from tools.workflow.workflow_runner import run_workflow


class WorkflowRunnerTests(unittest.TestCase):
    def create_workbook(self, path: Path) -> None:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Data"
        worksheet["A1"] = "source"
        worksheet["B1"] = "target"
        worksheet["A2"] = "第一行 [Alpha] 和 <color=red>{name}"
        worksheet["B2"] = "第一行 [阿尔法] 和 <color=red>{name}"
        worksheet["A3"] = "第二行复用 [Alpha] 和 <color=red>{name}"
        worksheet["B3"] = "第二行复用 [错误阿尔法] 和 {name}"
        workbook.save(path)

    def test_run_workflow_writes_term_and_tag_results_into_same_output_file(self) -> None:
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
            )

            self.assertEqual(summary.output_path, output_path.resolve())
            self.assertTrue(summary.ran_term_pair_check)
            self.assertTrue(summary.ran_tag_check)
            self.assertEqual(summary.term_problem_count, 1)
            self.assertEqual(summary.tag_problem_count, 1)

            workbook = load_workbook(summary.output_path)
            self.assertEqual(
                workbook.sheetnames,
                ["Data", "术语表", "问题列", "标签占位问题", "检查汇总"],
            )
            self.assertEqual(workbook["问题列"]["A2"].value, 3)
            self.assertEqual(workbook["标签占位问题"]["A2"].value, 3)
            self.assertEqual(workbook["检查汇总"]["B2"].value, "Data")

    def test_run_workflow_requires_at_least_one_task(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            self.create_workbook(input_path)

            with self.assertRaisesRegex(ValueError, "请至少选择一个 workflow 任务"):
                run_workflow(
                    input_file=input_path,
                    source_column="A",
                    target_column="B",
                    run_term_pair_check=False,
                    run_tag_check=False,
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
            )

            self.assertEqual(summary.term_count, 1)
            self.assertEqual(summary.term_problem_count, 1)

            workbook = load_workbook(summary.output_path)
            term_sheet = workbook["术语表"]
            self.assertEqual(term_sheet["B2"].value, "历史译法")
            self.assertEqual(term_sheet["E2"].value, "历史TB")


if __name__ == "__main__":
    unittest.main()
