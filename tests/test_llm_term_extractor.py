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


class LlmTermWorkbookTests(unittest.TestCase):
    def test_process_excel_handles_source_only_and_source_target_rows(self) -> None:
        from openpyxl import Workbook, load_workbook

        from tools.llm_term_extractor.codex_term_review import ExtractedLlmTerm, RowExtraction
        from tools.llm_term_extractor.extract_llm_terms import process_excel

        with tempfile.TemporaryDirectory(prefix="tag-exactor-llm-workbook-") as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "Data"
            worksheet["A1"] = "source"
            worksheet["B1"] = "target"
            worksheet["A2"] = "Unlock the Abyssal Vault."
            worksheet["B2"] = "解锁深渊宝库。"
            worksheet["A3"] = ""
            worksheet["B3"] = ""
            worksheet["A4"] = "Claim the Heart Flower Gift Box."
            worksheet["B4"] = ""
            workbook.save(input_path)

            def fake_extractor(rows):
                self.assertEqual([row.row_id for row in rows], ["2", "4"])
                return [
                    RowExtraction(
                        row_id="2",
                        terms=(
                            ExtractedLlmTerm(
                                source_term="Abyssal Vault",
                                target_term="深渊宝库",
                                category="item",
                                note="fixed item name",
                            ),
                        ),
                    ),
                    RowExtraction(
                        row_id="4",
                        terms=(
                            ExtractedLlmTerm(
                                source_term="Heart Flower Gift Box",
                                target_term="",
                                category="item",
                                note="source-only item name",
                            ),
                        ),
                    ),
                    RowExtraction(
                        row_id="99",
                        terms=(ExtractedLlmTerm(source_term="Ignored Term", target_term="忽略"),),
                    ),
                ]

            summary = process_excel(
                input_file=input_path,
                source_column="A",
                target_column="B",
                sheet="Data",
                start_row=2,
                batch_extractor=fake_extractor,
            )

            self.assertEqual(summary.output_path, input_path.with_name("input_llm_terms.xlsx"))
            self.assertEqual(summary.worksheet_title, "Data")
            self.assertEqual(summary.scanned_row_count, 2)
            self.assertEqual(summary.batch_count, 1)
            self.assertEqual(summary.term_count, 2)
            self.assertEqual(summary.evidence_count, 2)
            self.assertEqual(summary.import_candidate_count, 1)
            self.assertEqual(summary.review_before_import_count, 1)
            self.assertEqual(summary.already_in_history_count, 0)

            result = load_workbook(summary.output_path)
            self.assertEqual(
                result.sheetnames[-7:],
                [
                    "Terms_Source_Dedup",
                    "Extraction_Evidence",
                    "Conflicts_To_Review",
                    "Import_Candidate",
                    "Review_Before_Import",
                    "Already_In_History",
                    "Summary",
                ],
            )

            terms = result["Terms_Source_Dedup"]
            self.assertEqual(terms["A2"].value, "Abyssal Vault")
            self.assertEqual(terms["B2"].value, "深渊宝库")
            self.assertEqual(terms["A3"].value, "Heart Flower Gift Box")

            import_candidates = result["Import_Candidate"]
            self.assertEqual(import_candidates["A2"].value, "Abyssal Vault")
            self.assertEqual(import_candidates["B2"].value, "深渊宝库")

            review = result["Review_Before_Import"]
            self.assertEqual(review["A2"].value, "Heart Flower Gift Box")
            self.assertEqual(review["D2"].value, "target缺失")

    def test_process_excel_routes_real_conflicts_to_review_sheets(self) -> None:
        from openpyxl import Workbook, load_workbook

        from tools.llm_term_extractor.codex_term_review import (
            ConflictDecision,
            ExtractedLlmTerm,
            RowExtraction,
        )
        from tools.llm_term_extractor.extract_llm_terms import process_excel

        with tempfile.TemporaryDirectory(prefix="tag-exactor-llm-conflict-") as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "Data"
            worksheet["A1"] = "source"
            worksheet["B1"] = "target"
            worksheet["A2"] = "Flower Art gameplay"
            worksheet["B2"] = "Art Floral"
            worksheet["A3"] = "Flower Art piece"
            worksheet["B3"] = "composition florale"
            workbook.save(input_path)

            def fake_extractor(rows):
                return [
                    RowExtraction(
                        row_id="2",
                        terms=(
                            ExtractedLlmTerm(
                                source_term="Flower Art",
                                target_term="Art Floral",
                                category="system_or_concept",
                                note="fixed concept",
                            ),
                        ),
                    ),
                    RowExtraction(
                        row_id="3",
                        terms=(
                            ExtractedLlmTerm(
                                source_term="  flower   art ",
                                target_term="composition florale",
                                category="system_or_concept",
                                note="same source with another target",
                            ),
                        ),
                    ),
                ]

            def fake_conflict_reviewer(groups):
                self.assertEqual(len(groups), 1)
                self.assertEqual(groups[0].source_term, "Flower Art")
                self.assertEqual(groups[0].target_terms, ("Art Floral", "composition florale"))
                return [
                    ConflictDecision(
                        group_id=groups[0].group_id,
                        decision="conflict",
                        reason="substantial target wording difference",
                    )
                ]

            summary = process_excel(
                input_file=input_path,
                source_column="A",
                target_column="B",
                sheet="Data",
                start_row=2,
                batch_extractor=fake_extractor,
                conflict_reviewer=fake_conflict_reviewer,
            )

            self.assertEqual(summary.term_count, 1)
            self.assertEqual(summary.evidence_count, 2)
            self.assertEqual(summary.conflict_count, 1)
            self.assertEqual(summary.import_candidate_count, 0)
            self.assertEqual(summary.review_before_import_count, 1)

            result = load_workbook(summary.output_path)
            terms = result["Terms_Source_Dedup"]
            self.assertEqual(terms["A2"].value, "Flower Art")
            self.assertIn("Art Floral", terms["B2"].value)
            self.assertIn("composition florale", terms["B2"].value)

            conflicts = result["Conflicts_To_Review"]
            self.assertEqual(conflicts["A2"].value, "Flower Art")
            self.assertIn("Art Floral", conflicts["B2"].value)
            self.assertIn("composition florale", conflicts["B2"].value)
            self.assertEqual(conflicts["D2"].value, "conflict")
            self.assertEqual(conflicts["F2"].value, "substantial target wording difference")

            review = result["Review_Before_Import"]
            self.assertEqual(review["A2"].value, "Flower Art")
            self.assertEqual(review["D2"].value, "conflict")

    def test_process_excel_routes_unreviewed_multi_target_terms_as_conflicts(self) -> None:
        from openpyxl import Workbook, load_workbook

        from tools.llm_term_extractor.codex_term_review import ExtractedLlmTerm, RowExtraction
        from tools.llm_term_extractor.extract_llm_terms import process_excel

        with tempfile.TemporaryDirectory(prefix="tag-exactor-llm-unreviewed-conflict-") as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "Data"
            worksheet["A1"] = "source"
            worksheet["B1"] = "target"
            worksheet["A2"] = "Guild Trial stage"
            worksheet["B2"] = "Epreuve de guilde"
            worksheet["A3"] = "Guild Trial reward"
            worksheet["B3"] = "Defi de guilde"
            workbook.save(input_path)

            def fake_extractor(rows):
                return [
                    RowExtraction(
                        row_id="2",
                        terms=(
                            ExtractedLlmTerm(
                                source_term="Guild Trial",
                                target_term="Epreuve de guilde",
                                category="event",
                                note="event name",
                            ),
                        ),
                    ),
                    RowExtraction(
                        row_id="3",
                        terms=(
                            ExtractedLlmTerm(
                                source_term="Guild Trial",
                                target_term="Defi de guilde",
                                category="event",
                                note="same event different target",
                            ),
                        ),
                    ),
                ]

            summary = process_excel(
                input_file=input_path,
                source_column="A",
                target_column="B",
                sheet="Data",
                start_row=2,
                batch_extractor=fake_extractor,
            )

            self.assertEqual(summary.term_count, 1)
            self.assertEqual(summary.conflict_count, 1)
            self.assertEqual(summary.import_candidate_count, 0)
            self.assertEqual(summary.review_before_import_count, 1)

            result = load_workbook(summary.output_path)
            conflicts = result["Conflicts_To_Review"]
            self.assertEqual(conflicts["A2"].value, "Guild Trial")
            self.assertIn("Epreuve de guilde", conflicts["B2"].value)
            self.assertIn("Defi de guilde", conflicts["B2"].value)
            self.assertEqual(conflicts["D2"].value, "review")
            self.assertEqual(conflicts["F2"].value, "多译法需确认")

            review = result["Review_Before_Import"]
            self.assertEqual(review["A2"].value, "Guild Trial")
            self.assertEqual(review["D2"].value, "多译法需确认")
