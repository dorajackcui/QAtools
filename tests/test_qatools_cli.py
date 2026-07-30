from __future__ import annotations

import io
from pathlib import Path
from types import SimpleNamespace
import sys
import unittest
from contextlib import redirect_stderr, redirect_stdout
from unittest.mock import patch

from qatools.cli import COMMANDS, command_map, format_help, main


class CommandRegistryTests(unittest.TestCase):
    def test_command_names_and_aliases_are_unique(self) -> None:
        mapping = command_map()
        expected_name_count = sum(
            1 + len(command.aliases)
            for command in COMMANDS
        )

        self.assertEqual(len(mapping), expected_name_count)
        self.assertEqual(mapping["strings"].name, "phraseloom")
        self.assertEqual(mapping["workflow"].name, "qa")

    def test_top_level_help_lists_every_canonical_command(self) -> None:
        help_text = format_help()

        for command in COMMANDS:
            self.assertIn(command.name, help_text)

    def test_unknown_command_returns_usage_error(self) -> None:
        error = io.StringIO()

        with redirect_stderr(error):
            exit_code = main(["does-not-exist"])

        self.assertEqual(exit_code, 2)
        self.assertIn("未知命令", error.getvalue())

    def test_argv_command_receives_unmodified_remaining_arguments(self) -> None:
        received: list[str] = []
        fake_module = SimpleNamespace(
            main=lambda argv: received.extend(argv) or 0,
        )

        with patch(
            "qatools.cli.importlib.import_module",
            return_value=fake_module,
        ):
            exit_code = main(["qa", "input.xlsx", "-c", "A", "-t", "B"])

        self.assertEqual(exit_code, 0)
        self.assertEqual(received, ["input.xlsx", "-c", "A", "-t", "B"])

    def test_legacy_command_uses_scoped_sys_argv_and_restores_it(self) -> None:
        original_argv = sys.argv[:]
        received: list[str] = []

        def fake_main() -> None:
            received.extend(sys.argv)

        fake_module = SimpleNamespace(main=fake_main)
        with patch(
            "qatools.cli.importlib.import_module",
            return_value=fake_module,
        ):
            exit_code = main(
                ["tag-check", "input.xlsx", "-c", "A", "-t", "B"]
            )

        self.assertEqual(exit_code, 0)
        self.assertEqual(
            received,
            ["qatools tag-check", "input.xlsx", "-c", "A", "-t", "B"],
        )
        self.assertEqual(sys.argv, original_argv)

    def test_help_subcommand_delegates_to_native_command_help(self) -> None:
        output = io.StringIO()

        with redirect_stdout(output):
            exit_code = main(["help", "line-break-check"])

        self.assertEqual(exit_code, 0)
        self.assertIn("--source-column", output.getvalue())
        self.assertIn("--target-column", output.getvalue())

    def test_every_registered_command_exposes_help(self) -> None:
        for command in COMMANDS:
            with (
                self.subTest(command=command.name),
                redirect_stdout(io.StringIO()),
                redirect_stderr(io.StringIO()),
            ):
                self.assertEqual(main([command.name, "--help"]), 0)

    def test_cli_guide_mentions_every_canonical_command(self) -> None:
        guide = (
            Path(__file__).parents[1] / "docs" / "cli-usage.md"
        ).read_text(encoding="utf-8")

        for command in COMMANDS:
            self.assertIn(f"qatools {command.name}", guide)


if __name__ == "__main__":
    unittest.main()
