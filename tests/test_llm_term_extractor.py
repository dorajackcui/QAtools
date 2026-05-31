from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from subprocess import CompletedProcess
from unittest.mock import patch

from tools.llm_term_extractor.codex_term_review import (
    DEFAULT_CODEX_MODEL,
    DEFAULT_REASONING_EFFORT,
    ConflictGroup,
    InputBatchRow,
    build_codex_command,
    load_prompt_template,
    parse_conflict_response,
    parse_extraction_response,
    render_conflict_prompt,
    render_extraction_prompt,
    run_codex_prompt,
)


class CodexTermReviewTests(unittest.TestCase):
    def test_render_extraction_prompt_includes_mode_rows_and_schema(self) -> None:
        template = load_prompt_template(
            Path("tools/llm_term_extractor/prompts/extract_terms_zh_target.md")
        )
        rows = [
            InputBatchRow(
                row_id="row-1",
                source_text="Unlock the Abyssal Vault in Event Clash.",
                target_text="在活动冲突中解锁深渊宝库。",
            )
        ]

        prompt = render_extraction_prompt(template, mode="source_target", rows=rows)

        self.assertIn("source_target", prompt)
        self.assertIn('"row_id": "row-1"', prompt)
        self.assertIn("Abyssal Vault", prompt)
        self.assertIn("source decides whether something is a term", prompt)
        self.assertIn("fixed names/events/gameplay/items/currencies/systems/titles/fixed phrases/cross-context anchors", prompt)
        self.assertIn("do not collect ordinary UI state/action/adjective/full sentence text", prompt)
        self.assertIn("extract existing target expression only; do not recommend or rewrite", prompt)
        self.assertIn('"rows"', prompt)
        self.assertIn('"terms"', prompt)
        self.assertNotIn("{{MODE}}", prompt)
        self.assertNotIn("{{ROWS_JSON}}", prompt)
        self.assertNotIn("{{OUTPUT_SCHEMA}}", prompt)

    def test_parse_extraction_response_reads_plain_and_fenced_json(self) -> None:
        plain = (
            '{"rows":[{"row_id":"row-1","terms":[{"source_term":"Abyssal Vault",'
            '"target_term":"深渊宝库","category":"item","note":"第一行\n第二行"}]}]}'
        )
        fenced = (
            "```json\n"
            '{"rows":[{"row_id":"row-2","terms":[{"source_term":"Event Clash",'
            '"target_term":"活动冲突","category":"event","note":"fixed event"}]}]}\n'
            "```"
        )

        plain_rows = parse_extraction_response(plain)
        fenced_rows = parse_extraction_response(fenced)

        self.assertEqual(plain_rows[0].row_id, "row-1")
        self.assertEqual(plain_rows[0].terms[0].source_term, "Abyssal Vault")
        self.assertEqual(plain_rows[0].terms[0].target_term, "深渊宝库")
        self.assertEqual(plain_rows[0].terms[0].note, "第一行\n第二行")
        self.assertEqual(fenced_rows[0].row_id, "row-2")
        self.assertEqual(fenced_rows[0].terms[0].source_term, "Event Clash")

    def test_parse_conflict_response_reads_decisions(self) -> None:
        template = load_prompt_template(
            Path("tools/llm_term_extractor/prompts/conflict_review_zh_target.md")
        )
        prompt = render_conflict_prompt(
            template,
            groups=[
                ConflictGroup(
                    group_id="group-1",
                    source_term="Abyssal Vault",
                    target_terms=("深渊宝库", "深渊金库"),
                    examples=("Unlock the Abyssal Vault.",),
                )
            ],
        )
        response = (
            '{"decisions":[{"group_id":"group-1","decision":"conflict",'
            '"canonical_target":"深渊宝库","reason":"official-looking variant"}]}'
        )

        decisions = parse_conflict_response(response)

        self.assertIn("ignore case, singular/plural, ordinary grammar, punctuation/mark-only differences", prompt)
        self.assertIn(
            "flag substantial target wording, official-looking variant, conceptual difference, or project-context decisions",
            prompt,
        )
        self.assertIn("return strict JSON only", prompt)
        self.assertIn('"groups"', prompt)
        self.assertIn('"decisions"', prompt)
        self.assertNotIn("{{GROUPS_JSON}}", prompt)
        self.assertNotIn("{{OUTPUT_SCHEMA}}", prompt)
        self.assertEqual(decisions[0].group_id, "group-1")
        self.assertEqual(decisions[0].decision, "conflict")
        self.assertEqual(decisions[0].canonical_target, "深渊宝库")
        self.assertEqual(decisions[0].reason, "official-looking variant")

    def test_build_codex_command_uses_spark_model_and_reasoning_effort(self) -> None:
        command = build_codex_command(
            Path("/tmp/codex-output.txt"),
            DEFAULT_CODEX_MODEL,
            DEFAULT_REASONING_EFFORT,
        )

        self.assertEqual(command[0], "codex")
        self.assertIn("--ask-for-approval", command)
        self.assertIn("never", command)
        self.assertIn('model_reasoning_effort="high"', command)
        self.assertIn("--ephemeral", command)
        self.assertIn("--ignore-rules", command)
        self.assertIn("--skip-git-repo-check", command)
        self.assertIn("read-only", command)
        self.assertEqual(command[command.index("--output-last-message") + 1], "/tmp/codex-output.txt")
        self.assertEqual(command[command.index("--model") + 1], "gpt-5.3-codex-spark")
        self.assertEqual(command[-1], "-")

    def test_run_codex_prompt_writes_prompt_to_stdin_and_reads_output_file(self) -> None:
        with tempfile.TemporaryDirectory(prefix="tag-exactor-llm-test-") as tmp_dir:
            output_path = Path(tmp_dir) / "codex-output.txt"

            def fake_run(command, input, **kwargs):
                self.assertEqual(input, "prompt text")
                self.assertEqual(command[command.index("--output-last-message") + 1], str(output_path))
                output_path.write_text('{"rows":[]}', encoding="utf-8")
                return CompletedProcess(command, 0, stdout="stdout fallback", stderr="")

            with patch("tools.llm_term_extractor.codex_term_review.subprocess.run", side_effect=fake_run):
                output = run_codex_prompt(
                    "prompt text",
                    output_path,
                    DEFAULT_CODEX_MODEL,
                    DEFAULT_REASONING_EFFORT,
                    timeout_seconds=12,
                )

        self.assertEqual(output, '{"rows":[]}')


if __name__ == "__main__":
    unittest.main()
