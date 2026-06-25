from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from subprocess import CompletedProcess
from unittest.mock import patch

from tools.codex_runner import build_codex_exec_command, run_codex_exec_prompt


class CodexRunnerTests(unittest.TestCase):
    def test_build_codex_exec_command_sets_model_reasoning_and_output(self) -> None:
        command = build_codex_exec_command(
            output_path=Path("/tmp/codex-output.txt"),
            model="gpt-5.3-codex-spark",
            reasoning_effort="high",
        )

        self.assertEqual(command[0], "codex")
        self.assertIn("--ask-for-approval", command)
        self.assertIn("never", command)
        self.assertIn('model_reasoning_effort="high"', command)
        self.assertIn("--ephemeral", command)
        self.assertIn("--ignore-rules", command)
        self.assertIn("--skip-git-repo-check", command)
        self.assertIn("read-only", command)
        self.assertEqual(
            command[command.index("--output-last-message") + 1],
            str(Path("/tmp/codex-output.txt")),
        )
        self.assertEqual(command[command.index("--model") + 1], "gpt-5.3-codex-spark")
        self.assertEqual(command[-1], "-")

    def test_run_codex_exec_prompt_reads_output_last_message(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            output_path = Path(tmp_dir) / "codex-output.txt"

            def fake_run(command, input, **kwargs):
                self.assertEqual(input, "prompt text")
                self.assertEqual(command[command.index("--output-last-message") + 1], str(output_path))
                output_path.write_text('{"ok": true}', encoding="utf-8")
                return CompletedProcess(command, 0, "", "")

            with patch("tools.codex_runner.subprocess.run", side_effect=fake_run):
                output = run_codex_exec_prompt(
                    "prompt text",
                    output_path=output_path,
                    model="gpt-5.3-codex-spark",
                    reasoning_effort="high",
                )

        self.assertEqual(output, '{"ok": true}')

    def test_run_codex_exec_prompt_raises_on_nonzero_exit(self) -> None:
        def fake_run(command, input, **kwargs):
            return CompletedProcess(command, 1, "stdout text", "stderr text")

        with tempfile.TemporaryDirectory() as tmp_dir:
            output_path = Path(tmp_dir) / "codex-output.txt"
            with patch("tools.codex_runner.subprocess.run", side_effect=fake_run):
                with self.assertRaises(RuntimeError) as context:
                    run_codex_exec_prompt("prompt text", output_path=output_path)

        error_text = str(context.exception)
        self.assertIn("exit 1", error_text)
        self.assertIn("stdout text", error_text)
        self.assertIn("stderr text", error_text)


if __name__ == "__main__":
    unittest.main()
