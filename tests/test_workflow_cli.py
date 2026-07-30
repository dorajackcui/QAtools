from __future__ import annotations

import io
from pathlib import Path
import unittest
from contextlib import redirect_stdout
from unittest.mock import patch

from tools.workflow.cli import build_parser, main
from tools.workflow.workflow_runner import WorkflowSummary


class WorkflowCliTests(unittest.TestCase):
    def test_parser_defaults_to_all_checks(self) -> None:
        args = build_parser().parse_args(
            ["input.xlsx", "-c", "A", "-t", "B"]
        )

        self.assertIsNone(args.check)
        self.assertEqual(args.start_row, 2)
        self.assertFalse(args.no_term_mark)

    def test_main_maps_selected_checks_and_options_to_runner(self) -> None:
        summary = WorkflowSummary(
            output_path=Path("output.xlsx"),
            worksheet_title="Sheet1",
            source_column="A",
            target_column="B",
            start_row=2,
            ran_term_pair_check=False,
            ran_tag_check=True,
            ran_line_break_check=False,
            ran_source_consistency_check=False,
            ran_chinese_target_check=True,
            term_count=0,
            term_problem_count=0,
            term_problem_rows=0,
            tag_problem_count=2,
            tag_problem_rows=1,
            line_break_problem_count=0,
            source_consistency_problem_count=0,
            source_consistency_problem_rows=0,
            chinese_target_problem_count=3,
        )

        with (
            patch(
                "tools.workflow.cli.run_workflow",
                return_value=summary,
            ) as run,
            redirect_stdout(io.StringIO()) as output,
        ):
            exit_code = main(
                [
                    "input.xlsx",
                    "-s",
                    "Game",
                    "-c",
                    "A",
                    "-t",
                    "B",
                    "--check",
                    "tag",
                    "--check",
                    "chinese",
                    "--tag-token-type",
                    "angle",
                    "-o",
                    "output.xlsx",
                ]
            )

        self.assertEqual(exit_code, 0)
        run.assert_called_once()
        kwargs = run.call_args.kwargs
        self.assertFalse(kwargs["run_term_pair_check"])
        self.assertTrue(kwargs["run_tag_check"])
        self.assertFalse(kwargs["run_line_break_check"])
        self.assertFalse(kwargs["run_source_consistency_check"])
        self.assertTrue(kwargs["run_chinese_target_check"])
        self.assertEqual(kwargs["tag_token_types"], ["angle"])
        self.assertIn("质量检查完成", output.getvalue())

    def test_no_term_mark_is_forwarded_as_empty_mark_collection(self) -> None:
        summary = WorkflowSummary(
            output_path=Path("output.xlsx"),
            worksheet_title="Sheet1",
            source_column="A",
            target_column="B",
            start_row=2,
            ran_term_pair_check=True,
            ran_tag_check=False,
            ran_line_break_check=False,
            ran_source_consistency_check=False,
            ran_chinese_target_check=False,
            term_count=0,
            term_problem_count=0,
            term_problem_rows=0,
            tag_problem_count=0,
            tag_problem_rows=0,
            line_break_problem_count=0,
            source_consistency_problem_count=0,
            source_consistency_problem_rows=0,
            chinese_target_problem_count=0,
        )

        with (
            patch(
                "tools.workflow.cli.run_workflow",
                return_value=summary,
            ) as run,
            redirect_stdout(io.StringIO()),
        ):
            main(
                [
                    "input.xlsx",
                    "-c",
                    "A",
                    "-t",
                    "B",
                    "--check",
                    "term",
                    "--no-term-mark",
                    "--history-tb",
                    "history.xlsx",
                ]
            )

        self.assertEqual(run.call_args.kwargs["term_mark_styles"], ())
        self.assertEqual(
            run.call_args.kwargs["term_history_tb_file"],
            "history.xlsx",
        )


if __name__ == "__main__":
    unittest.main()
