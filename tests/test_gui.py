from __future__ import annotations

import io
import unittest
from contextlib import redirect_stdout
from unittest.mock import patch

from phraseloom.gui import TASK_BY_KEY, TASKS, build_cli_args, validate_task_specs


class GuiTaskSpecTests(unittest.TestCase):
    def test_task_specs_are_valid_and_cover_cli_workflows(self) -> None:
        validate_task_specs()
        expected_commands = {
            "prepare",
            "fill",
            "tm-extract",
            "extract",
            "entity-tm",
            "entity-prepare",
            "entity-fill-pack",
            "entity-merge-pack",
            "entity-split",
            "entity-prefill",
            "entity-extract-tm",
            "entity-fill",
            "entity-merge",
        }
        self.assertEqual({task.command for task in TASKS}, expected_commands)

    def test_prepare_builds_cli_args_with_explicit_prefill_choices(self) -> None:
        args = build_cli_args(
            TASK_BY_KEY["prepare"],
            {
                "input": "/tmp/source.xlsx",
                "tm": "/tmp/tm.xlsx",
                "use_existing_targets": True,
                "source_col": "source",
                "target_col": "target",
                "context_col": "notes",
            },
        )
        self.assertEqual(args[0:2], ["prepare", "/tmp/source.xlsx"])
        self.assertIn("--tm", args)
        self.assertIn("/tmp/tm.xlsx", args)
        self.assertIn("--use-existing-targets", args)
        self.assertIn("--context-col", args)
        self.assertIn("notes", args)
        self.assertNotIn("--min-group-size", args)

    def test_prepare_and_tm_forms_include_optional_context_column(self) -> None:
        for task_key in ("prepare", "tm_extract"):
            fields = {field.key: field for field in TASK_BY_KEY[task_key].fields}
            self.assertIn("context_col", fields)
            self.assertEqual(fields["context_col"].flag, "--context-col")
            self.assertTrue(fields["context_col"].advanced)

    def test_daily_fill_needs_only_translator_workbook(self) -> None:
        args = build_cli_args(
            TASK_BY_KEY["fill"],
            {"input": "/tmp/source_translator_todo.xlsx"},
        )
        self.assertEqual(args, ["fill", "/tmp/source_translator_todo.xlsx"])

    def test_repeat_fields_and_entity_merge_flags_are_supported(self) -> None:
        extract_args = build_cli_args(
            TASK_BY_KEY["extract_report"],
            {
                "input": "/tmp/source.xlsx",
                "examples": "One=一\nTwo=二\n",
                "min_group_size": "2",
            },
        )
        self.assertEqual(extract_args.count("--example"), 2)

        merge_args = build_cli_args(
            TASK_BY_KEY["entity_merge"],
            {
                "entity": "/tmp/entity.xlsx",
                "non_entity": "/tmp/non_entity.xlsx",
            },
        )
        self.assertEqual(
            merge_args,
            [
                "entity-merge",
                "--entity",
                "/tmp/entity.xlsx",
                "--non-entity",
                "/tmp/non_entity.xlsx",
            ],
        )

    def test_invalid_or_conflicting_values_are_rejected(self) -> None:
        with self.assertRaisesRegex(ValueError, "请填写"):
            build_cli_args(TASK_BY_KEY["fill"], {})
        with self.assertRaisesRegex(ValueError, "大于 0"):
            build_cli_args(
                TASK_BY_KEY["tm_extract"],
                {"input": "/tmp/tm.xlsx", "min_group_size": "0"},
            )
        with self.assertRaisesRegex(ValueError, "不能同时使用"):
            build_cli_args(
                TASK_BY_KEY["entity_fill_pack"],
                {
                    "input": "/tmp/entity_pack.xlsx",
                    "output": "/tmp/result.xlsx",
                    "in_place": True,
                },
            )

    def test_cli_dispatches_gui_without_eagerly_starting_tk(self) -> None:
        from phraseloom.cli import _dispatch

        with patch("phraseloom.gui.main", return_value=0) as gui_main:
            self.assertEqual(_dispatch(["gui"]), 0)
        gui_main.assert_called_once_with()

    def test_top_level_help_lists_gui(self) -> None:
        from phraseloom.cli import _dispatch

        output = io.StringIO()
        with redirect_stdout(output):
            self.assertEqual(_dispatch(["--help"]), 0)
        self.assertIn("phraseloom gui", output.getvalue())


if __name__ == "__main__":
    unittest.main()
