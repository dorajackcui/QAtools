# LLM Term Extractor Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a new Toolshub Excel utility that uses local `codex exec --model gpt-5.3-codex-spark` to extract unmarked source terms, extract existing target terms when present, deduplicate by source, and flag real target conflicts.

**Architecture:** Add a focused `tools/llm_term_extractor` package with external prompt files, a Codex subprocess adapter, pure aggregation/routing helpers, an Excel CLI processor, and a Tkinter GUI. Tests inject fake extraction and conflict-review callables so workbook behavior is covered without running Codex.

**Tech Stack:** Python 3, `openpyxl`, `tkinter`, `unittest`, local Codex CLI.

---

## File Structure

- Create `tools/llm_term_extractor/__init__.py`: package marker.
- Create `tools/llm_term_extractor/prompts/extract_terms_zh_target.md`: editable extraction prompt template.
- Create `tools/llm_term_extractor/prompts/conflict_review_zh_target.md`: editable conflict-review prompt template.
- Create `tools/llm_term_extractor/codex_term_review.py`: prompt loading/rendering, Codex command construction, JSON extraction, response parsing, retry-aware batch helpers.
- Create `tools/llm_term_extractor/extract_llm_terms.py`: CLI, Excel row loading, batching, aggregation, history TB loading, workbook output.
- Create `tools/llm_term_extractor/extract_llm_terms_gui.py`: standalone Tkinter GUI frame.
- Create `tools/llm_term_extractor/README.md`: user documentation.
- Create `tests/test_llm_term_extractor.py`: core, CLI, workbook, and Codex-adapter tests.
- Modify `toolshub_gui.py`: import and add the LLM term extractor GUI tab.
- Modify `README.md`: list the new tool.
- Modify `docs/cli-usage.md`: document non-interactive usage and command templates.
- Modify `tests/test_gui_excel_selection.py`: add GUI metadata tests if current no-display fake-var patterns cover the new frame cleanly.

### Task 1: Prompt Files And Codex Adapter

**Files:**
- Create: `tools/llm_term_extractor/__init__.py`
- Create: `tools/llm_term_extractor/prompts/extract_terms_zh_target.md`
- Create: `tools/llm_term_extractor/prompts/conflict_review_zh_target.md`
- Create: `tools/llm_term_extractor/codex_term_review.py`
- Test: `tests/test_llm_term_extractor.py`

- [ ] **Step 1: Write failing tests for prompt rendering, JSON parsing, and Codex command construction**

Add this initial test module:

```python
from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from subprocess import CompletedProcess
from unittest.mock import patch

from tools.llm_term_extractor.codex_term_review import (
    ConflictGroup,
    InputBatchRow,
    build_codex_command,
    extract_json_object,
    parse_conflict_response,
    parse_extraction_response,
    render_conflict_prompt,
    render_extraction_prompt,
    run_codex_prompt,
)


class CodexTermReviewTests(unittest.TestCase):
    def test_render_extraction_prompt_includes_mode_rows_and_schema(self) -> None:
        template = "规则正文\n{{MODE}}\n{{ROWS_JSON}}\n{{OUTPUT_SCHEMA}}"
        prompt = render_extraction_prompt(
            template,
            mode="mixed",
            rows=[
                InputBatchRow(row_index=2, source_text="获得心花礼盒", target_text="Obtenez le Coffret Coeur-Fleur"),
                InputBatchRow(row_index=3, source_text="花艺等级提升", target_text=""),
            ],
        )

        self.assertIn("mixed", prompt)
        self.assertIn('"row_index": 2', prompt)
        self.assertIn("心花礼盒", prompt)
        self.assertIn('"target_term"', prompt)

    def test_parse_extraction_response_reads_plain_and_fenced_json(self) -> None:
        text = """```json
        {"rows":[{"row_index":2,"terms":[{"source_term":"花艺","target_term":"Art Floral","term_type":"system_or_concept","confidence":"high","note":"固定概念"}]}]}
        ```"""

        rows = parse_extraction_response(text)

        self.assertEqual(rows[0].row_index, 2)
        self.assertEqual(rows[0].terms[0].source_term, "花艺")
        self.assertEqual(rows[0].terms[0].target_term, "Art Floral")

    def test_parse_conflict_response_reads_decisions(self) -> None:
        text = '{"results":[{"source_term":"花艺","decision":"conflict","category":"实质译名差异","confidence":"high","note":"需要统一"}]}'

        decisions = parse_conflict_response(text)

        self.assertEqual(decisions["花艺"].decision, "conflict")
        self.assertEqual(decisions["花艺"].category, "实质译名差异")

    def test_build_codex_command_uses_spark_model_and_reasoning_effort(self) -> None:
        command = build_codex_command(
            output_path=Path("/tmp/codex-output.txt"),
            model="gpt-5.3-codex-spark",
            reasoning_effort="high",
        )

        self.assertIn("--model", command)
        self.assertIn("gpt-5.3-codex-spark", command)
        self.assertIn('model_reasoning_effort="high"', command)
        self.assertIn("--output-last-message", command)

    def test_run_codex_prompt_writes_prompt_to_stdin_and_reads_output_file(self) -> None:
        def fake_run(command, input, **kwargs):
            self.assertEqual(input, "prompt text")
            output_path = Path(command[command.index("--output-last-message") + 1])
            output_path.write_text('{"rows":[]}', encoding="utf-8")
            return CompletedProcess(command, 0, "", "")

        with tempfile.TemporaryDirectory() as tmp_dir:
            output_path = Path(tmp_dir) / "codex-output.txt"
            with patch("tools.llm_term_extractor.codex_term_review.subprocess.run", side_effect=fake_run):
                output = run_codex_prompt(
                    "prompt text",
                    output_path=output_path,
                    model="gpt-5.3-codex-spark",
                    reasoning_effort="high",
                )

        self.assertEqual(output, '{"rows":[]}')
```

- [ ] **Step 2: Run tests to verify they fail**

Run:

```bash
python3 -m unittest tests.test_llm_term_extractor.CodexTermReviewTests
```

Expected: FAIL with `ModuleNotFoundError: No module named 'tools.llm_term_extractor'`.

- [ ] **Step 3: Add prompt templates**

Create `tools/llm_term_extractor/prompts/extract_terms_zh_target.md`:

```markdown
你是本地化术语提取助手。

任务：从 Excel source 文案中判断哪些片段应该进入术语表。是否是术语只由 source 决定。

术语口径：
- 收固定名、活动名、玩法名、道具名、货币名、系统名、称号、固定短语、跨上下文需要统一译法的表达。
- 不收普通 UI 状态词、普通操作词、普通形容词、普通完整句子。
- “是否容易翻译不一致”优先于“是不是名词”。
- target 为空时，只收 source_term。
- target 非空时，从同行 target 中抽取已有对应译法；不要推荐译法，不要改译文。

当前模式：
{{MODE}}

待处理行：
{{ROWS_JSON}}

输出要求：
{{OUTPUT_SCHEMA}}

只返回严格 JSON，不要 Markdown，不要解释。
```

Create `tools/llm_term_extractor/prompts/conflict_review_zh_target.md`:

```markdown
你是本地化术语冲突复核助手。

任务：判断同一个 source_term 的多个 target_term 观察值是否是实质冲突。

忽略：
- 大小写变化
- 明显单复数变化
- 普通语法形态变化
- 标点或 mark 差异

标记为 conflict 或 review：
- 明显不同定稿
- 不同官方风格译名
- 概念不同
- 需要项目语境才能决定是否统一

待复核组：
{{GROUPS_JSON}}

输出要求：
{{OUTPUT_SCHEMA}}

只返回严格 JSON，不要 Markdown，不要解释。
```

- [ ] **Step 4: Implement `codex_term_review.py` minimally**

Create `tools/llm_term_extractor/codex_term_review.py`:

```python
#!/usr/bin/env python3
"""Codex prompt rendering and response parsing for the LLM term extractor."""

from __future__ import annotations

import json
import re
import subprocess
from dataclasses import dataclass
from pathlib import Path
from typing import Any

DEFAULT_CODEX_MODEL = "gpt-5.3-codex-spark"
DEFAULT_REASONING_EFFORT = "high"
DEFAULT_TIMEOUT_SECONDS = 600


@dataclass(frozen=True)
class InputBatchRow:
    row_index: int
    source_text: str
    target_text: str = ""


@dataclass(frozen=True)
class ExtractedLlmTerm:
    source_term: str
    target_term: str
    term_type: str
    confidence: str
    note: str


@dataclass(frozen=True)
class RowExtraction:
    row_index: int
    terms: tuple[ExtractedLlmTerm, ...]


@dataclass(frozen=True)
class ConflictGroup:
    source_term: str
    target_terms: tuple[str, ...]
    rows: tuple[dict[str, object], ...]


@dataclass(frozen=True)
class ConflictDecision:
    source_term: str
    decision: str
    category: str
    confidence: str
    note: str


EXTRACTION_OUTPUT_SCHEMA = """{
  "rows": [
    {
      "row_index": 2,
      "terms": [
        {
          "source_term": "花艺",
          "target_term": "Art Floral",
          "term_type": "system_or_concept",
          "confidence": "high",
          "note": "固定系统/概念名，跨上下文需要一致"
        }
      ]
    }
  ]
}"""

CONFLICT_OUTPUT_SCHEMA = """{
  "results": [
    {
      "source_term": "花艺",
      "decision": "conflict",
      "category": "实质译名差异",
      "confidence": "high",
      "note": "建议人工确认是否按语境区分"
    }
  ]
}"""


def load_prompt_template(path: str | Path) -> str:
    return Path(path).expanduser().resolve().read_text(encoding="utf-8")


def render_extraction_prompt(template: str, *, mode: str, rows: list[InputBatchRow]) -> str:
    rows_payload = [
        {"row_index": row.row_index, "source": row.source_text, "target": row.target_text}
        for row in rows
    ]
    return (
        template.replace("{{MODE}}", mode)
        .replace("{{ROWS_JSON}}", json.dumps(rows_payload, ensure_ascii=False, indent=2))
        .replace("{{OUTPUT_SCHEMA}}", EXTRACTION_OUTPUT_SCHEMA)
    )


def render_conflict_prompt(template: str, *, groups: list[ConflictGroup]) -> str:
    groups_payload = [
        {
            "source_term": group.source_term,
            "target_terms": list(group.target_terms),
            "rows": list(group.rows),
        }
        for group in groups
    ]
    return (
        template.replace("{{GROUPS_JSON}}", json.dumps(groups_payload, ensure_ascii=False, indent=2))
        .replace("{{OUTPUT_SCHEMA}}", CONFLICT_OUTPUT_SCHEMA)
    )


def extract_json_object(text: str) -> dict[str, Any]:
    stripped = text.strip()
    fenced_match = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", stripped, flags=re.DOTALL)
    if fenced_match:
        stripped = fenced_match.group(1).strip()
    if not stripped.startswith("{"):
        start = stripped.find("{")
        end = stripped.rfind("}")
        if start == -1 or end == -1 or end < start:
            raise ValueError("Codex 输出中没有 JSON 对象。")
        stripped = stripped[start : end + 1]
    parsed = json.loads(stripped, strict=False)
    if not isinstance(parsed, dict):
        raise ValueError("Codex 输出 JSON 不是对象。")
    return parsed


def parse_extraction_response(text: str) -> list[RowExtraction]:
    parsed = extract_json_object(text)
    rows = parsed.get("rows")
    if not isinstance(rows, list):
        raise ValueError("Codex 提取输出缺少 rows 数组。")

    parsed_rows: list[RowExtraction] = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        row_index = int(row.get("row_index", 0) or 0)
        terms = row.get("terms", [])
        if row_index < 1 or not isinstance(terms, list):
            continue
        parsed_terms: list[ExtractedLlmTerm] = []
        for term in terms:
            if not isinstance(term, dict):
                continue
            source_term = str(term.get("source_term") or "").strip()
            if not source_term:
                continue
            parsed_terms.append(
                ExtractedLlmTerm(
                    source_term=source_term,
                    target_term=str(term.get("target_term") or "").strip(),
                    term_type=str(term.get("term_type") or "").strip(),
                    confidence=str(term.get("confidence") or "").strip(),
                    note=str(term.get("note") or "").strip(),
                )
            )
        parsed_rows.append(RowExtraction(row_index=row_index, terms=tuple(parsed_terms)))
    return parsed_rows


def parse_conflict_response(text: str) -> dict[str, ConflictDecision]:
    parsed = extract_json_object(text)
    results = parsed.get("results")
    if not isinstance(results, list):
        raise ValueError("Codex 冲突输出缺少 results 数组。")

    decisions: dict[str, ConflictDecision] = {}
    for item in results:
        if not isinstance(item, dict):
            continue
        source_term = str(item.get("source_term") or "").strip()
        if not source_term:
            continue
        decisions[source_term] = ConflictDecision(
            source_term=source_term,
            decision=str(item.get("decision") or "review").strip() or "review",
            category=str(item.get("category") or "需人工确认").strip() or "需人工确认",
            confidence=str(item.get("confidence") or "low").strip() or "low",
            note=str(item.get("note") or "").strip(),
        )
    return decisions


def build_codex_command(
    *,
    output_path: Path,
    model: str = DEFAULT_CODEX_MODEL,
    reasoning_effort: str = DEFAULT_REASONING_EFFORT,
) -> list[str]:
    return [
        "codex",
        "--ask-for-approval",
        "never",
        "--config",
        f'model_reasoning_effort="{reasoning_effort}"',
        "exec",
        "--ephemeral",
        "--ignore-rules",
        "--skip-git-repo-check",
        "--sandbox",
        "read-only",
        "--output-last-message",
        str(output_path),
        "--model",
        model,
        "-",
    ]


def run_codex_prompt(
    prompt: str,
    *,
    output_path: Path,
    model: str = DEFAULT_CODEX_MODEL,
    reasoning_effort: str = DEFAULT_REASONING_EFFORT,
    timeout_seconds: int = DEFAULT_TIMEOUT_SECONDS,
) -> str:
    command = build_codex_command(
        output_path=output_path,
        model=model,
        reasoning_effort=reasoning_effort,
    )
    completed = subprocess.run(
        command,
        input=prompt,
        capture_output=True,
        text=True,
        timeout=timeout_seconds,
        check=False,
    )
    if completed.returncode != 0:
        error_text = completed.stderr.strip() or completed.stdout.strip()
        raise RuntimeError(f"Codex 术语提取失败: {error_text}")
    if output_path.exists():
        return output_path.read_text(encoding="utf-8")
    return completed.stdout
```

- [ ] **Step 5: Run tests to verify they pass**

Run:

```bash
python3 -m unittest tests.test_llm_term_extractor.CodexTermReviewTests
```

Expected: all `CodexTermReviewTests` pass.

- [ ] **Step 6: Commit**

```bash
git add tools/llm_term_extractor/__init__.py \
  tools/llm_term_extractor/prompts/extract_terms_zh_target.md \
  tools/llm_term_extractor/prompts/conflict_review_zh_target.md \
  tools/llm_term_extractor/codex_term_review.py \
  tests/test_llm_term_extractor.py
git commit -m "feat: add LLM term extractor Codex adapter"
```

### Task 2: Aggregation And Workbook Output

**Files:**
- Create: `tools/llm_term_extractor/extract_llm_terms.py`
- Modify: `tests/test_llm_term_extractor.py`

- [ ] **Step 1: Write failing workbook tests with fake extraction and conflict review**

Append these tests:

```python
from openpyxl import Workbook, load_workbook

from tools.llm_term_extractor.extract_llm_terms import process_excel
from tools.llm_term_extractor.codex_term_review import ConflictDecision, ExtractedLlmTerm, RowExtraction


class LlmTermWorkbookTests(unittest.TestCase):
    def test_process_excel_handles_source_only_and_source_target_rows(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "Data"
            worksheet["A1"] = "source"
            worksheet["B1"] = "target"
            worksheet["A2"] = "花艺等级提升"
            worksheet["B2"] = "Amélioration Art Floral"
            worksheet["A3"] = "获得心花礼盒"
            worksheet["B3"] = ""
            workbook.save(input_path)

            def fake_extractor(rows):
                return [
                    RowExtraction(
                        row_index=2,
                        terms=(ExtractedLlmTerm("花艺", "Art Floral", "system_or_concept", "high", "固定概念"),),
                    ),
                    RowExtraction(
                        row_index=3,
                        terms=(ExtractedLlmTerm("心花礼盒", "", "item", "high", "道具名"),),
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

            self.assertEqual(summary.term_count, 2)
            self.assertEqual(summary.import_candidate_count, 1)
            self.assertEqual(summary.review_before_import_count, 1)

            result = load_workbook(summary.output_path)
            self.assertIn("Terms_Source_Dedup", result.sheetnames)
            self.assertIn("Extraction_Evidence", result.sheetnames)
            self.assertIn("Import_Candidate", result.sheetnames)
            self.assertIn("Review_Before_Import", result.sheetnames)

            terms = result["Terms_Source_Dedup"]
            self.assertEqual(terms["A2"].value, "花艺")
            self.assertEqual(terms["B2"].value, "Art Floral")
            self.assertEqual(terms["A3"].value, "心花礼盒")
            self.assertEqual(terms["B3"].value, "")

    def test_process_excel_routes_real_conflicts_to_review_sheets(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "Data"
            worksheet["A1"] = "source"
            worksheet["B1"] = "target"
            worksheet["A2"] = "花艺玩法"
            worksheet["B2"] = "Art Floral"
            worksheet["A3"] = "花艺作品"
            worksheet["B3"] = "composition florale"
            workbook.save(input_path)

            def fake_extractor(rows):
                return [
                    RowExtraction(2, (ExtractedLlmTerm("花艺", "Art Floral", "system_or_concept", "high", "固定概念"),)),
                    RowExtraction(3, (ExtractedLlmTerm("花艺", "composition florale", "system_or_concept", "high", "固定概念"),)),
                ]

            def fake_conflict_reviewer(groups):
                return {
                    "花艺": ConflictDecision(
                        source_term="花艺",
                        decision="conflict",
                        category="实质译名差异",
                        confidence="high",
                        note="Art Floral 和 composition florale 需要人工确认",
                    )
                }

            summary = process_excel(
                input_file=input_path,
                source_column="A",
                target_column="B",
                sheet="Data",
                start_row=2,
                batch_extractor=fake_extractor,
                conflict_reviewer=fake_conflict_reviewer,
            )

            self.assertEqual(summary.conflict_count, 1)
            self.assertEqual(summary.import_candidate_count, 0)

            result = load_workbook(summary.output_path)
            conflicts = result["Conflicts_To_Review"]
            self.assertEqual(conflicts["A2"].value, "花艺")
            self.assertIn("Art Floral", conflicts["B2"].value)
            self.assertIn("composition florale", conflicts["B2"].value)
            self.assertEqual(conflicts["D2"].value, "实质译名差异")
```

- [ ] **Step 2: Run tests to verify they fail**

Run:

```bash
python3 -m unittest tests.test_llm_term_extractor.LlmTermWorkbookTests
```

Expected: FAIL with import error for `tools.llm_term_extractor.extract_llm_terms`.

- [ ] **Step 3: Implement core data model and output workbook**

Create `tools/llm_term_extractor/extract_llm_terms.py` with these core definitions:

```python
#!/usr/bin/env python3
"""Extract unmarked terms from Excel source/target text with Codex assistance."""

from __future__ import annotations

import argparse
import json
import re
import sys
import tempfile
from collections import defaultdict
from collections.abc import Callable, Iterable
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter

from tools.llm_term_extractor.codex_term_review import (
    DEFAULT_CODEX_MODEL,
    DEFAULT_REASONING_EFFORT,
    ConflictDecision,
    ConflictGroup,
    ExtractedLlmTerm,
    InputBatchRow,
    RowExtraction,
)

TERMS_SHEET_NAME = "Terms_Source_Dedup"
EVIDENCE_SHEET_NAME = "Extraction_Evidence"
CONFLICTS_SHEET_NAME = "Conflicts_To_Review"
IMPORT_SHEET_NAME = "Import_Candidate"
REVIEW_SHEET_NAME = "Review_Before_Import"
HISTORY_SHEET_NAME = "Already_In_History"
SUMMARY_SHEET_NAME = "Summary"

BatchExtractor = Callable[[list[InputBatchRow]], list[RowExtraction]]
ConflictReviewer = Callable[[list[ConflictGroup]], dict[str, ConflictDecision]]


@dataclass(frozen=True)
class SourceRow:
    row_index: int
    source_text: str
    target_text: str


@dataclass(frozen=True)
class ExtractionObservation:
    row_index: int
    source_text: str
    target_text: str
    source_term: str
    target_term: str
    term_type: str
    confidence: str
    note: str


@dataclass
class AggregatedTerm:
    source_term: str
    source_key: str
    target_terms: list[str] = field(default_factory=list)
    row_indexes: list[int] = field(default_factory=list)
    source_examples: list[str] = field(default_factory=list)
    target_examples: list[str] = field(default_factory=list)
    term_types: set[str] = field(default_factory=set)
    confidences: set[str] = field(default_factory=set)
    notes: list[str] = field(default_factory=list)
    history_target: str = ""
    conflict_decision: ConflictDecision | None = None


@dataclass(frozen=True)
class LlmTermExtractionSummary:
    output_path: Path
    worksheet_title: str
    source_column: str
    target_column: str
    start_row: int
    scanned_row_count: int
    batch_count: int
    term_count: int
    evidence_count: int
    conflict_count: int
    import_candidate_count: int
    review_before_import_count: int
    already_in_history_count: int
```

Add these pure helpers:

```python
def normalize_column(column_name: str) -> str:
    normalized = column_name.strip().upper()
    column_index_from_string(normalized)
    return normalized


def normalize_term_key(value: str) -> str:
    return re.sub(r"\s+", " ", value.strip()).casefold()


def unique_join(values: Iterable[str], separator: str = "、") -> str:
    seen: set[str] = set()
    unique_values: list[str] = []
    for value in values:
        normalized = value.strip()
        if not normalized or normalized in seen:
            continue
        seen.add(normalized)
        unique_values.append(normalized)
    return separator.join(unique_values)


def build_default_output_path(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}_llm_terms{input_path.suffix}")


def iter_batches(rows: list[SourceRow], batch_size: int) -> Iterable[list[SourceRow]]:
    if batch_size < 1:
        raise ValueError("batch-size 必须大于等于 1。")
    for start_index in range(0, len(rows), batch_size):
        yield rows[start_index : start_index + batch_size]


def batch_mode(rows: list[InputBatchRow]) -> str:
    has_target = [bool(row.target_text.strip()) for row in rows]
    if all(has_target):
        return "source_target"
    if any(has_target):
        return "mixed"
    return "source_only"
```

Add aggregation:

```python
def collect_observations(
    rows_by_index: dict[int, SourceRow],
    extractions: list[RowExtraction],
) -> list[ExtractionObservation]:
    observations: list[ExtractionObservation] = []
    for row_extraction in extractions:
        source_row = rows_by_index.get(row_extraction.row_index)
        if source_row is None:
            continue
        for term in row_extraction.terms:
            if not term.source_term.strip():
                continue
            observations.append(
                ExtractionObservation(
                    row_index=row_extraction.row_index,
                    source_text=source_row.source_text,
                    target_text=source_row.target_text,
                    source_term=term.source_term.strip(),
                    target_term=term.target_term.strip(),
                    term_type=term.term_type.strip(),
                    confidence=term.confidence.strip(),
                    note=term.note.strip(),
                )
            )
    return observations


def aggregate_observations(
    observations: list[ExtractionObservation],
    history_mapping: dict[str, str] | None = None,
) -> dict[str, AggregatedTerm]:
    history_mapping = history_mapping or {}
    aggregated: dict[str, AggregatedTerm] = {}
    for observation in observations:
        source_key = normalize_term_key(observation.source_term)
        if not source_key:
            continue
        term = aggregated.setdefault(
            source_key,
            AggregatedTerm(
                source_term=observation.source_term,
                source_key=source_key,
                history_target=history_mapping.get(source_key, ""),
            ),
        )
        if observation.target_term and observation.target_term not in term.target_terms:
            term.target_terms.append(observation.target_term)
        if observation.row_index not in term.row_indexes:
            term.row_indexes.append(observation.row_index)
        if observation.source_text and len(term.source_examples) < 3:
            term.source_examples.append(observation.source_text)
        if observation.target_text and len(term.target_examples) < 3:
            term.target_examples.append(observation.target_text)
        if observation.term_type:
            term.term_types.add(observation.term_type)
        if observation.confidence:
            term.confidences.add(observation.confidence)
        if observation.note and observation.note not in term.notes:
            term.notes.append(observation.note)
    return aggregated
```

Add workbook writing with headers:

```python
def rebuild_output_sheet(workbook, current_sheet_name: str, sheet_name: str):
    if current_sheet_name == sheet_name:
        raise ValueError(f"数据工作表名称不能为 {sheet_name}")
    if sheet_name in workbook.sheetnames:
        del workbook[sheet_name]
    return workbook.create_sheet(title=sheet_name)


def write_output_sheets(
    workbook,
    worksheet_title: str,
    observations: list[ExtractionObservation],
    aggregated_terms: dict[str, AggregatedTerm],
    summary: dict[str, object],
) -> tuple[int, int, int, int]:
    terms_sheet = rebuild_output_sheet(workbook, worksheet_title, TERMS_SHEET_NAME)
    terms_sheet.append([
        "source_term",
        "target_terms_observed",
        "count",
        "rows",
        "term_types",
        "confidence",
        "history_target",
        "conflict_decision",
        "conflict_note",
        "notes",
    ])

    evidence_sheet = rebuild_output_sheet(workbook, worksheet_title, EVIDENCE_SHEET_NAME)
    evidence_sheet.append([
        "row_index",
        "source_term",
        "target_term",
        "term_type",
        "confidence",
        "note",
        "source_text",
        "target_text",
    ])

    conflicts_sheet = rebuild_output_sheet(workbook, worksheet_title, CONFLICTS_SHEET_NAME)
    conflicts_sheet.append([
        "source_term",
        "target_terms_observed",
        "rows",
        "category",
        "confidence",
        "note",
        "source_examples",
        "target_examples",
    ])

    import_sheet = rebuild_output_sheet(workbook, worksheet_title, IMPORT_SHEET_NAME)
    import_sheet.append(["source_term", "target_term", "rows", "count", "notes"])

    review_sheet = rebuild_output_sheet(workbook, worksheet_title, REVIEW_SHEET_NAME)
    review_sheet.append(["source_term", "target_terms_observed", "rows", "reason", "notes"])

    history_sheet = rebuild_output_sheet(workbook, worksheet_title, HISTORY_SHEET_NAME)
    history_sheet.append(["source_term", "history_target", "rows", "target_terms_observed"])

    for observation in observations:
        evidence_sheet.append([
            observation.row_index,
            observation.source_term,
            observation.target_term,
            observation.term_type,
            observation.confidence,
            observation.note,
            observation.source_text,
            observation.target_text,
        ])

    conflict_count = 0
    import_count = 0
    review_count = 0
    history_count = 0
    for term in sorted(aggregated_terms.values(), key=lambda item: item.source_term):
        target_terms = unique_join(term.target_terms)
        rows = unique_join(str(row_index) for row_index in term.row_indexes)
        notes = unique_join(term.notes, separator=" | ")
        decision = term.conflict_decision.decision if term.conflict_decision else ""
        decision_note = term.conflict_decision.note if term.conflict_decision else ""
        terms_sheet.append([
            term.source_term,
            target_terms,
            len(term.row_indexes),
            rows,
            unique_join(sorted(term.term_types)),
            unique_join(sorted(term.confidences)),
            term.history_target,
            decision,
            decision_note,
            notes,
        ])

        if term.history_target:
            history_count += 1
            history_sheet.append([term.source_term, term.history_target, rows, target_terms])
            continue

        if term.conflict_decision and term.conflict_decision.decision in {"conflict", "review"}:
            conflict_count += 1
            review_count += 1
            conflicts_sheet.append([
                term.source_term,
                target_terms,
                rows,
                term.conflict_decision.category,
                term.conflict_decision.confidence,
                term.conflict_decision.note,
                unique_join(term.source_examples, separator=" | "),
                unique_join(term.target_examples, separator=" | "),
            ])
            review_sheet.append([term.source_term, target_terms, rows, term.conflict_decision.decision, notes])
        elif len(term.target_terms) == 1:
            import_count += 1
            import_sheet.append([term.source_term, term.target_terms[0], rows, len(term.row_indexes), notes])
        else:
            review_count += 1
            reason = "target缺失" if not term.target_terms else "多译法需确认"
            review_sheet.append([term.source_term, target_terms, rows, reason, notes])

    summary_sheet = rebuild_output_sheet(workbook, worksheet_title, SUMMARY_SHEET_NAME)
    summary_sheet.append(["metric", "value"])
    for key, value in summary.items():
        summary_sheet.append([key, value])
    summary_sheet.append(["conflict_count", conflict_count])
    summary_sheet.append(["import_candidate_count", import_count])
    summary_sheet.append(["review_before_import_count", review_count])
    summary_sheet.append(["already_in_history_count", history_count])

    return conflict_count, import_count, review_count, history_count
```

- [ ] **Step 4: Run tests and fix narrow failures**

Run:

```bash
python3 -m unittest tests.test_llm_term_extractor.LlmTermWorkbookTests
```

Expected: tests pass after implementing `process_excel` in the next task; if these tests fail because `process_excel` is still missing, proceed directly to Task 3 Step 1 and keep the failing output as the RED state.

### Task 3: CLI Orchestration, Codex Retry, And History TB

**Files:**
- Modify: `tools/llm_term_extractor/extract_llm_terms.py`
- Modify: `tests/test_llm_term_extractor.py`

- [ ] **Step 1: Write failing tests for `process_excel`, history routing, and CLI defaults**

Append:

```python
class LlmTermCliAndHistoryTests(unittest.TestCase):
    def test_process_excel_routes_history_matches(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "input.xlsx"
            history_path = Path(tmp_dir) / "history.xlsx"

            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "Data"
            worksheet["A1"] = "source"
            worksheet["B1"] = "target"
            worksheet["A2"] = "花艺等级提升"
            worksheet["B2"] = "Amélioration Art Floral"
            workbook.save(input_path)

            history = Workbook()
            history_sheet = history.active
            history_sheet.title = "术语表"
            history_sheet["A1"] = "source"
            history_sheet["B1"] = "target"
            history_sheet["A2"] = "花艺"
            history_sheet["B2"] = "Art Floral"
            history.save(history_path)

            def fake_extractor(rows):
                return [
                    RowExtraction(2, (ExtractedLlmTerm("花艺", "Art Floral", "system_or_concept", "high", "固定概念"),))
                ]

            summary = process_excel(
                input_file=input_path,
                source_column="A",
                target_column="B",
                sheet="Data",
                history_tb_file=history_path,
                batch_extractor=fake_extractor,
            )

            self.assertEqual(summary.already_in_history_count, 1)
            self.assertEqual(summary.import_candidate_count, 0)

            result = load_workbook(summary.output_path)
            history_output = result["Already_In_History"]
            self.assertEqual(history_output["A2"].value, "花艺")
            self.assertEqual(history_output["B2"].value, "Art Floral")

    def test_default_output_path_uses_llm_terms_suffix(self) -> None:
        from tools.llm_term_extractor.extract_llm_terms import build_default_output_path

        self.assertEqual(
            build_default_output_path(Path("/tmp/task.xlsx")),
            Path("/tmp/task_llm_terms.xlsx"),
        )
```

- [ ] **Step 2: Run tests to verify expected failure**

Run:

```bash
python3 -m unittest tests.test_llm_term_extractor
```

Expected: FAIL because `process_excel` and history helpers are incomplete.

- [ ] **Step 3: Implement row loading, history loading, default Codex-backed callables, and `process_excel`**

Add to `extract_llm_terms.py`:

```python
def read_source_rows(
    worksheet,
    *,
    source_column: str,
    target_column: str | None,
    start_row: int,
) -> list[SourceRow]:
    rows: list[SourceRow] = []
    for row_index in range(start_row, worksheet.max_row + 1):
        source_value = worksheet[f"{source_column}{row_index}"].value
        target_value = worksheet[f"{target_column}{row_index}"].value if target_column else ""
        source_text = "" if source_value is None else str(source_value)
        target_text = "" if target_value is None else str(target_value)
        if not source_text.strip() and not target_text.strip():
            continue
        rows.append(SourceRow(row_index=row_index, source_text=source_text, target_text=target_text))
    return rows


def load_history_tb_mapping(
    history_tb_file: str | Path,
    *,
    sheet: str | None = None,
    source_column: str | None = None,
    target_column: str | None = None,
    start_row: int = 2,
) -> dict[str, str]:
    path = Path(history_tb_file).expanduser().resolve()
    workbook = load_workbook(path, read_only=True)
    try:
        worksheet = workbook[sheet] if sheet else workbook["术语表"] if "术语表" in workbook.sheetnames else workbook.active
        header = next(worksheet.iter_rows(min_row=1, max_row=1, values_only=True), ())
        if not source_column or not target_column:
            normalized_headers = ["" if value is None else str(value).strip().casefold().replace(" ", "") for value in header]
            source_candidates = {"source", "source术语", "source术语（无mark）", "source术语(无mark)"}
            target_candidates = {"target", "target术语", "target术语（无mark）", "target术语(无mark)"}
            for index, value in enumerate(normalized_headers, start=1):
                if not source_column and value in source_candidates:
                    source_column = get_column_letter(index)
                if not target_column and value in target_candidates:
                    target_column = get_column_letter(index)
            nonempty = [get_column_letter(index) for index, value in enumerate(normalized_headers, start=1) if value]
            if (not source_column or not target_column) and len(nonempty) == 2:
                source_column = source_column or nonempty[0]
                target_column = target_column or nonempty[1]
        if not source_column or not target_column:
            raise ValueError("历史 TB 缺少 source/target 列。")

        source_index = column_index_from_string(normalize_column(source_column))
        target_index = column_index_from_string(normalize_column(target_column))
        max_index = max(source_index, target_index)
        mapping: dict[str, str] = {}
        for row in worksheet.iter_rows(min_row=start_row, max_col=max_index, values_only=True):
            source_value = row[source_index - 1] if len(row) >= source_index else None
            target_value = row[target_index - 1] if len(row) >= target_index else None
            source_text = "" if source_value is None else str(source_value).strip()
            target_text = "" if target_value is None else str(target_value).strip()
            if source_text and target_text:
                mapping.setdefault(normalize_term_key(source_text), target_text)
        return mapping
    finally:
        workbook.close()
```

Add Codex-backed factories:

```python
def default_prompt_path(name: str) -> Path:
    return Path(__file__).resolve().parent / "prompts" / name


def build_codex_batch_extractor(
    *,
    prompt_file: str | Path | None,
    model: str,
    reasoning_effort: str,
    dump_prompts_dir: str | Path | None,
    raw_output_path: Path | None,
) -> BatchExtractor:
    from tools.llm_term_extractor.codex_term_review import (
        load_prompt_template,
        parse_extraction_response,
        render_extraction_prompt,
        run_codex_prompt,
    )

    template = load_prompt_template(prompt_file or default_prompt_path("extract_terms_zh_target.md"))
    batch_counter = {"value": 0}

    def extractor(rows: list[InputBatchRow]) -> list[RowExtraction]:
        batch_counter["value"] += 1
        prompt = render_extraction_prompt(template, mode=batch_mode(rows), rows=rows)
        if dump_prompts_dir:
            prompt_dir = Path(dump_prompts_dir).expanduser().resolve()
            prompt_dir.mkdir(parents=True, exist_ok=True)
            (prompt_dir / f"extract-batch-{batch_counter['value']:04d}.md").write_text(prompt, encoding="utf-8")
        with tempfile.TemporaryDirectory(prefix="llm-term-extract-") as tmp_dir:
            output_text = run_codex_prompt(
                prompt,
                output_path=Path(tmp_dir) / "codex-output.txt",
                model=model,
                reasoning_effort=reasoning_effort,
            )
        if raw_output_path:
            raw_output_path.parent.mkdir(parents=True, exist_ok=True)
            with raw_output_path.open("a", encoding="utf-8") as raw_file:
                raw_file.write(json.dumps({"kind": "extract", "batch": batch_counter["value"], "output": output_text}, ensure_ascii=False) + "\n")
        try:
            return parse_extraction_response(output_text)
        except ValueError:
            retry_prompt = prompt + "\n\n上一次输出不是有效 JSON。请只返回符合 schema 的严格 JSON。"
            with tempfile.TemporaryDirectory(prefix="llm-term-extract-retry-") as tmp_dir:
                retry_text = run_codex_prompt(
                    retry_prompt,
                    output_path=Path(tmp_dir) / "codex-output.txt",
                    model=model,
                    reasoning_effort=reasoning_effort,
                )
            return parse_extraction_response(retry_text)

    return extractor
```

Add conflict grouping and `process_excel`:

```python
def build_conflict_groups(aggregated_terms: dict[str, AggregatedTerm]) -> list[ConflictGroup]:
    groups: list[ConflictGroup] = []
    for term in aggregated_terms.values():
        if len(term.target_terms) <= 1:
            continue
        groups.append(
            ConflictGroup(
                source_term=term.source_term,
                target_terms=tuple(term.target_terms),
                rows=tuple(
                    {
                        "row_index": row_index,
                        "source_examples": term.source_examples,
                        "target_examples": term.target_examples,
                    }
                    for row_index in term.row_indexes
                ),
            )
        )
    return groups


def apply_conflict_decisions(
    aggregated_terms: dict[str, AggregatedTerm],
    decisions: dict[str, ConflictDecision],
) -> None:
    for term in aggregated_terms.values():
        decision = decisions.get(term.source_term)
        if decision is not None:
            term.conflict_decision = decision


def process_excel(
    input_file: str | Path,
    source_column: str,
    target_column: str | None = None,
    sheet: str | None = None,
    start_row: int = 2,
    batch_size: int = 50,
    codex_model: str = DEFAULT_CODEX_MODEL,
    codex_reasoning_effort: str = DEFAULT_REASONING_EFFORT,
    extract_prompt_file: str | Path | None = None,
    conflict_prompt_file: str | Path | None = None,
    dump_prompts_dir: str | Path | None = None,
    keep_raw_codex_output: bool = False,
    history_tb_file: str | Path | None = None,
    history_sheet: str | None = None,
    history_source_column: str | None = None,
    history_target_column: str | None = None,
    history_start_row: int = 2,
    output_file: str | Path | None = None,
    batch_extractor: BatchExtractor | None = None,
    conflict_reviewer: ConflictReviewer | None = None,
) -> LlmTermExtractionSummary:
    input_path = Path(input_file).expanduser().resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"输入文件不存在: {input_path}")
    source_column = normalize_column(source_column)
    target_column = normalize_column(target_column) if target_column else None
    output_path = Path(output_file).expanduser().resolve() if output_file else build_default_output_path(input_path)
    raw_output_path = output_path.with_name(f"{output_path.stem}_codex_raw.jsonl") if keep_raw_codex_output else None

    history_mapping = (
        load_history_tb_mapping(
            history_tb_file,
            sheet=history_sheet,
            source_column=history_source_column,
            target_column=history_target_column,
            start_row=history_start_row,
        )
        if history_tb_file
        else {}
    )

    workbook = load_workbook(input_path)
    worksheet = workbook[sheet] if sheet else workbook.active
    source_rows = read_source_rows(
        worksheet,
        source_column=source_column,
        target_column=target_column,
        start_row=start_row,
    )
    rows_by_index = {row.row_index: row for row in source_rows}

    if batch_extractor is None:
        batch_extractor = build_codex_batch_extractor(
            prompt_file=extract_prompt_file,
            model=codex_model,
            reasoning_effort=codex_reasoning_effort,
            dump_prompts_dir=dump_prompts_dir,
            raw_output_path=raw_output_path,
        )

    all_extractions: list[RowExtraction] = []
    batch_count = 0
    for batch in iter_batches(source_rows, batch_size):
        batch_count += 1
        input_batch = [
            InputBatchRow(row_index=row.row_index, source_text=row.source_text, target_text=row.target_text)
            for row in batch
        ]
        all_extractions.extend(batch_extractor(input_batch))

    observations = collect_observations(rows_by_index, all_extractions)
    aggregated_terms = aggregate_observations(observations, history_mapping)

    conflict_groups = build_conflict_groups(aggregated_terms)
    if conflict_groups:
        if conflict_reviewer is None:
            conflict_reviewer = build_codex_conflict_reviewer(
                prompt_file=conflict_prompt_file,
                model=codex_model,
                reasoning_effort=codex_reasoning_effort,
                dump_prompts_dir=dump_prompts_dir,
                raw_output_path=raw_output_path,
            )
        apply_conflict_decisions(aggregated_terms, conflict_reviewer(conflict_groups))

    summary_values = {
        "worksheet": worksheet.title,
        "source_column": source_column,
        "target_column": target_column or "",
        "start_row": start_row,
        "scanned_row_count": len(source_rows),
        "batch_count": batch_count,
        "codex_model": codex_model,
        "codex_reasoning_effort": codex_reasoning_effort,
        "term_count": len(aggregated_terms),
        "evidence_count": len(observations),
    }
    conflict_count, import_count, review_count, history_count = write_output_sheets(
        workbook,
        worksheet.title,
        observations,
        aggregated_terms,
        summary_values,
    )
    workbook.save(output_path)

    return LlmTermExtractionSummary(
        output_path=output_path,
        worksheet_title=worksheet.title,
        source_column=source_column,
        target_column=target_column or "",
        start_row=start_row,
        scanned_row_count=len(source_rows),
        batch_count=batch_count,
        term_count=len(aggregated_terms),
        evidence_count=len(observations),
        conflict_count=conflict_count,
        import_candidate_count=import_count,
        review_before_import_count=review_count,
        already_in_history_count=history_count,
    )
```

- [ ] **Step 4: Implement `build_codex_conflict_reviewer`**

Add:

```python
def build_codex_conflict_reviewer(
    *,
    prompt_file: str | Path | None,
    model: str,
    reasoning_effort: str,
    dump_prompts_dir: str | Path | None,
    raw_output_path: Path | None,
) -> ConflictReviewer:
    from tools.llm_term_extractor.codex_term_review import (
        load_prompt_template,
        parse_conflict_response,
        render_conflict_prompt,
        run_codex_prompt,
    )

    template = load_prompt_template(prompt_file or default_prompt_path("conflict_review_zh_target.md"))

    def reviewer(groups: list[ConflictGroup]) -> dict[str, ConflictDecision]:
        prompt = render_conflict_prompt(template, groups=groups)
        if dump_prompts_dir:
            prompt_dir = Path(dump_prompts_dir).expanduser().resolve()
            prompt_dir.mkdir(parents=True, exist_ok=True)
            (prompt_dir / "conflict-review.md").write_text(prompt, encoding="utf-8")
        with tempfile.TemporaryDirectory(prefix="llm-term-conflict-") as tmp_dir:
            output_text = run_codex_prompt(
                prompt,
                output_path=Path(tmp_dir) / "codex-output.txt",
                model=model,
                reasoning_effort=reasoning_effort,
            )
        if raw_output_path:
            raw_output_path.parent.mkdir(parents=True, exist_ok=True)
            with raw_output_path.open("a", encoding="utf-8") as raw_file:
                raw_file.write(json.dumps({"kind": "conflict", "output": output_text}, ensure_ascii=False) + "\n")
        try:
            return parse_conflict_response(output_text)
        except ValueError:
            retry_prompt = prompt + "\n\n上一次输出不是有效 JSON。请只返回符合 schema 的严格 JSON。"
            with tempfile.TemporaryDirectory(prefix="llm-term-conflict-retry-") as tmp_dir:
                retry_text = run_codex_prompt(
                    retry_prompt,
                    output_path=Path(tmp_dir) / "codex-output.txt",
                    model=model,
                    reasoning_effort=reasoning_effort,
                )
            return parse_conflict_response(retry_text)

    return reviewer
```

- [ ] **Step 5: Add CLI argument parsing and `main`**

Add:

```python
def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="使用 Codex 从 Excel source/target 文案中提取未标记术语。")
    parser.add_argument("input_file", nargs="?", help="输入 Excel 文件路径")
    parser.add_argument("-s", "--sheet", help="工作表名称，不填则使用活动工作表")
    parser.add_argument("-c", "--source-column", help="source 列，例如 A")
    parser.add_argument("-t", "--target-column", help="target 列，例如 B；不填则只做 source 术语收集")
    parser.add_argument("--start-row", type=int, default=2, help="开始处理的行号，默认 2")
    parser.add_argument("--batch-size", type=int, default=50, help="每批发送给 Codex 的行数，默认 50")
    parser.add_argument("--codex-model", default=DEFAULT_CODEX_MODEL, help="Codex 模型，默认 gpt-5.3-codex-spark")
    parser.add_argument("--codex-reasoning-effort", default=DEFAULT_REASONING_EFFORT, choices=("low", "medium", "high", "xhigh"))
    parser.add_argument("--extract-prompt-file", help="自定义术语提取 prompt 文件")
    parser.add_argument("--conflict-prompt-file", help="自定义冲突复核 prompt 文件")
    parser.add_argument("--dump-prompts-dir", help="保存每批完整 prompt 的目录")
    parser.add_argument("--keep-raw-codex-output", action="store_true", help="保留 Codex 原始输出 JSONL")
    parser.add_argument("--history-tb", help="历史 TB Excel 文件路径")
    parser.add_argument("--history-sheet", help="历史 TB 工作表名称")
    parser.add_argument("--history-source-column", help="历史 TB source 列")
    parser.add_argument("--history-target-column", help="历史 TB target 列")
    parser.add_argument("--history-start-row", type=int, default=2, help="历史 TB 开始行，默认 2")
    parser.add_argument("-o", "--output", help="输出 Excel 文件路径")
    return parser.parse_args()


def prompt_if_missing(args: argparse.Namespace) -> argparse.Namespace:
    interactive_mode = sys.stdin.isatty()
    if not args.input_file and not interactive_mode:
        raise ValueError("缺少输入文件路径，请传入 input_file 参数。")
    if not args.input_file:
        args.input_file = input("请输入 Excel 文件路径: ").strip()
    if not args.source_column and not interactive_mode:
        raise ValueError("缺少 source 列，请使用 -c 或 --source-column 指定。")
    if not args.source_column:
        args.source_column = input("请输入 source 列（例如 A）: ").strip().upper()
    return args


def main() -> None:
    args = prompt_if_missing(parse_args())
    summary = process_excel(
        input_file=args.input_file,
        source_column=args.source_column,
        target_column=args.target_column,
        sheet=args.sheet,
        start_row=args.start_row,
        batch_size=args.batch_size,
        codex_model=args.codex_model,
        codex_reasoning_effort=args.codex_reasoning_effort,
        extract_prompt_file=args.extract_prompt_file,
        conflict_prompt_file=args.conflict_prompt_file,
        dump_prompts_dir=args.dump_prompts_dir,
        keep_raw_codex_output=args.keep_raw_codex_output,
        history_tb_file=args.history_tb,
        history_sheet=args.history_sheet,
        history_source_column=args.history_source_column,
        history_target_column=args.history_target_column,
        history_start_row=args.history_start_row,
        output_file=args.output,
    )
    print("处理完成。")
    print(f"工作表: {summary.worksheet_title}")
    print(f"source 列: {summary.source_column}")
    print(f"target 列: {summary.target_column or '未指定'}")
    print(f"扫描行数: {summary.scanned_row_count}")
    print(f"batch 数: {summary.batch_count}")
    print(f"术语数: {summary.term_count}")
    print(f"冲突数: {summary.conflict_count}")
    print(f"输出文件: {summary.output_path}")


if __name__ == "__main__":
    main()
```

- [ ] **Step 6: Run focused tests**

Run:

```bash
python3 -m unittest tests.test_llm_term_extractor
```

Expected: all LLM term extractor tests pass.

- [ ] **Step 7: Commit**

```bash
git add tools/llm_term_extractor/extract_llm_terms.py tests/test_llm_term_extractor.py
git commit -m "feat: add LLM term extraction workbook processor"
```

### Task 4: GUI And Toolshub Integration

**Files:**
- Create: `tools/llm_term_extractor/extract_llm_terms_gui.py`
- Modify: `toolshub_gui.py`
- Modify: `tests/test_gui_excel_selection.py`

- [ ] **Step 1: Write GUI metadata tests if no-display patterns fit**

Add tests following the fake Tk variable style already used in `tests/test_gui_excel_selection.py`:

```python
def test_llm_term_extractor_refresh_populates_sheet_choices_and_detects_columns(self) -> None:
    from tools.llm_term_extractor.extract_llm_terms_gui import LlmTermExtractorApp

    with tempfile.TemporaryDirectory() as tmp_dir:
        workbook_path = Path(tmp_dir) / "input.xlsx"
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Data"
        worksheet["A1"] = "source"
        worksheet["B1"] = "target"
        workbook.save(workbook_path)

        app = LlmTermExtractorApp.__new__(LlmTermExtractorApp)
        app.input_file_var = FakeVar(str(workbook_path))
        app.sheet_var = FakeVar("")
        app.source_column_var = FakeVar("")
        app.target_column_var = FakeVar("")
        app.sheet_combobox = FakeCombobox()

        app.refresh_sheet_choices(show_error=False)

        self.assertEqual(app.sheet_var.get(), "Data")
        self.assertEqual(app.source_column_var.get(), "A")
        self.assertEqual(app.target_column_var.get(), "B")
```

If the existing fake classes are local to tests and not reusable, move them to module scope inside the same test file before adding this test.

- [ ] **Step 2: Run GUI test to verify failure**

Run:

```bash
python3 -m unittest tests.test_gui_excel_selection
```

Expected: FAIL because `extract_llm_terms_gui.py` does not exist.

- [ ] **Step 3: Implement `LlmTermExtractorApp`**

Create `tools/llm_term_extractor/extract_llm_terms_gui.py`:

```python
#!/usr/bin/env python3
"""Desktop UI for the LLM term extractor."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from tools.excel_metadata import detect_source_target_columns, list_workbook_sheets
from tools.llm_term_extractor.codex_term_review import DEFAULT_CODEX_MODEL, DEFAULT_REASONING_EFFORT
from tools.llm_term_extractor.extract_llm_terms import build_default_output_path, process_excel


class LlmTermExtractorApp(ttk.Frame):
    def __init__(self, master: tk.Misc) -> None:
        super().__init__(master, padding=16)
        self.input_file_var = tk.StringVar()
        self.output_file_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.source_column_var = tk.StringVar(value="A")
        self.target_column_var = tk.StringVar(value="B")
        self.start_row_var = tk.StringVar(value="2")
        self.batch_size_var = tk.StringVar(value="50")
        self.codex_model_var = tk.StringVar(value=DEFAULT_CODEX_MODEL)
        self.codex_reasoning_effort_var = tk.StringVar(value=DEFAULT_REASONING_EFFORT)
        self.extract_prompt_file_var = tk.StringVar()
        self.conflict_prompt_file_var = tk.StringVar()
        self.history_tb_file_var = tk.StringVar()
        self.keep_raw_output_var = tk.BooleanVar(value=False)
        self._build_ui()

    def _build_ui(self) -> None:
        labels = [
            ("输入 Excel", self.input_file_var, self.choose_input_file),
            ("输出 Excel", self.output_file_var, self.choose_output_file),
            ("提取 Prompt", self.extract_prompt_file_var, self.choose_extract_prompt_file),
            ("冲突 Prompt", self.conflict_prompt_file_var, self.choose_conflict_prompt_file),
            ("历史 TB Excel", self.history_tb_file_var, self.choose_history_tb_file),
        ]
        for row_index, (label, variable, command) in enumerate(labels):
            ttk.Label(self, text=label).grid(row=row_index, column=0, sticky="w", pady=(0, 8))
            ttk.Entry(self, textvariable=variable, width=46).grid(row=row_index, column=1, sticky="ew", pady=(0, 8))
            ttk.Button(self, text="选择", command=command).grid(row=row_index, column=2, padx=(8, 0), pady=(0, 8))

        ttk.Label(self, text="工作表名").grid(row=5, column=0, sticky="w", pady=(0, 8))
        self.sheet_combobox = ttk.Combobox(self, textvariable=self.sheet_var, width=20, state="readonly")
        self.sheet_combobox.grid(row=5, column=1, sticky="w", pady=(0, 8))
        self.sheet_combobox.bind("<<ComboboxSelected>>", self.handle_sheet_selected)

        field_rows = [
            ("Source 列", self.source_column_var),
            ("Target 列（可空）", self.target_column_var),
            ("开始行", self.start_row_var),
            ("Batch 行数", self.batch_size_var),
            ("Codex 模型", self.codex_model_var),
            ("Reasoning effort", self.codex_reasoning_effort_var),
        ]
        for offset, (label, variable) in enumerate(field_rows, start=6):
            ttk.Label(self, text=label).grid(row=offset, column=0, sticky="w", pady=(0, 8))
            ttk.Entry(self, textvariable=variable, width=24).grid(row=offset, column=1, sticky="w", pady=(0, 8))

        ttk.Checkbutton(self, text="保留 Codex 原始输出 JSONL", variable=self.keep_raw_output_var).grid(
            row=12, column=0, columnspan=3, sticky="w", pady=(0, 12)
        )
        ttk.Button(self, text="开始提取", command=self.run_extraction).grid(row=13, column=0, columnspan=3, sticky="ew")
        self.columnconfigure(1, weight=1)

    def choose_input_file(self) -> None:
        file_path = filedialog.askopenfilename(title="选择 Excel 文件", filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")])
        if not file_path:
            return
        self.input_file_var.set(file_path)
        if not self.output_file_var.get().strip():
            self.output_file_var.set(str(build_default_output_path(Path(file_path))))
        self.refresh_sheet_choices()

    def choose_output_file(self) -> None:
        file_path = filedialog.asksaveasfilename(title="选择输出 Excel 文件", defaultextension=".xlsx")
        if file_path:
            self.output_file_var.set(file_path)

    def choose_extract_prompt_file(self) -> None:
        file_path = filedialog.askopenfilename(title="选择提取 Prompt", filetypes=[("Markdown", "*.md"), ("所有文件", "*.*")])
        if file_path:
            self.extract_prompt_file_var.set(file_path)

    def choose_conflict_prompt_file(self) -> None:
        file_path = filedialog.askopenfilename(title="选择冲突 Prompt", filetypes=[("Markdown", "*.md"), ("所有文件", "*.*")])
        if file_path:
            self.conflict_prompt_file_var.set(file_path)

    def choose_history_tb_file(self) -> None:
        file_path = filedialog.askopenfilename(title="选择历史 TB Excel", filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")])
        if file_path:
            self.history_tb_file_var.set(file_path)

    def refresh_sheet_choices(self, show_error: bool = True) -> None:
        input_file = self.input_file_var.get().strip()
        if not input_file:
            self.sheet_combobox["values"] = ()
            self.sheet_var.set("")
            return
        try:
            choices = list_workbook_sheets(input_file)
        except Exception as exc:
            self.sheet_combobox["values"] = ()
            self.sheet_var.set("")
            if show_error:
                messagebox.showerror("读取失败", str(exc))
            return
        self.sheet_combobox["values"] = choices.sheet_names
        selected_sheet = self.sheet_var.get().strip()
        if selected_sheet not in choices.sheet_names:
            self.sheet_var.set(choices.default_sheet or (choices.sheet_names[0] if choices.sheet_names else ""))
        self.handle_sheet_selected(show_error=show_error)

    def handle_sheet_selected(self, event: object | None = None, show_error: bool = True) -> None:
        input_file = self.input_file_var.get().strip()
        sheet = self.sheet_var.get().strip()
        if not input_file or not sheet:
            return
        try:
            detected = detect_source_target_columns(input_file, sheet=sheet)
        except Exception as exc:
            if show_error:
                messagebox.showwarning("列识别失败", str(exc))
            return
        if detected.source_column:
            self.source_column_var.set(detected.source_column)
        if detected.target_column:
            self.target_column_var.set(detected.target_column)

    def run_extraction(self) -> None:
        try:
            summary = process_excel(
                input_file=self.input_file_var.get().strip(),
                source_column=self.source_column_var.get().strip(),
                target_column=self.target_column_var.get().strip() or None,
                sheet=self.sheet_var.get().strip() or None,
                start_row=int(self.start_row_var.get().strip() or "2"),
                batch_size=int(self.batch_size_var.get().strip() or "50"),
                codex_model=self.codex_model_var.get().strip() or DEFAULT_CODEX_MODEL,
                codex_reasoning_effort=self.codex_reasoning_effort_var.get().strip() or DEFAULT_REASONING_EFFORT,
                extract_prompt_file=self.extract_prompt_file_var.get().strip() or None,
                conflict_prompt_file=self.conflict_prompt_file_var.get().strip() or None,
                history_tb_file=self.history_tb_file_var.get().strip() or None,
                keep_raw_codex_output=self.keep_raw_output_var.get(),
                output_file=self.output_file_var.get().strip() or None,
            )
        except Exception as exc:
            messagebox.showerror("处理失败", str(exc))
            return
        messagebox.showinfo(
            "处理完成",
            f"术语数: {summary.term_count}\n冲突数: {summary.conflict_count}\n输出: {summary.output_path}",
        )


def main() -> None:
    root = tk.Tk()
    root.title("LLM 术语提取")
    LlmTermExtractorApp(root).grid(row=0, column=0, sticky="nsew")
    root.mainloop()


if __name__ == "__main__":
    main()
```

- [ ] **Step 4: Register the Toolshub tab**

Modify `toolshub_gui.py`:

```python
from tools.llm_term_extractor.extract_llm_terms_gui import LlmTermExtractorApp
```

Instantiate and add after the workflow tab:

```python
llm_term_tab = LlmTermExtractorApp(notebook)
notebook.add(llm_term_tab, text="LLM术语提取")
```

Update the subtitle text to include `LLM术语提取`.

- [ ] **Step 5: Run GUI tests**

Run:

```bash
python3 -m unittest tests.test_gui_excel_selection
```

Expected: all GUI metadata tests pass.

- [ ] **Step 6: Commit**

```bash
git add tools/llm_term_extractor/extract_llm_terms_gui.py toolshub_gui.py tests/test_gui_excel_selection.py
git commit -m "feat: add LLM term extractor GUI"
```

### Task 5: Documentation And Final Verification

**Files:**
- Create: `tools/llm_term_extractor/README.md`
- Modify: `README.md`
- Modify: `docs/cli-usage.md`

- [ ] **Step 1: Write tool README**

Create `tools/llm_term_extractor/README.md`:

```markdown
# LLM 术语提取工具

用于从 Excel 的 source/target 文案中提取未标记术语。工具通过本机 `codex exec` 调用 `gpt-5.3-codex-spark`。

## 两种模式

- target 为空或未指定 target 列：只收集 source 术语。
- target 非空：收集 source-target 已有术语对，并检查同一 source 是否出现实质不同 target。

## Prompt

默认 prompt 位于：

- `tools/llm_term_extractor/prompts/extract_terms_zh_target.md`
- `tools/llm_term_extractor/prompts/conflict_review_zh_target.md`

可以用 `--extract-prompt-file` 和 `--conflict-prompt-file` 覆盖。

## CLI

```bash
python3 tools/llm_term_extractor/extract_llm_terms.py input.xlsx \
  -s Sheet1 \
  -c A \
  -t B \
  --start-row 2 \
  --batch-size 50 \
  --codex-model gpt-5.3-codex-spark \
  -o output_llm_terms.xlsx
```

只做 source 术语收集：

```bash
python3 tools/llm_term_extractor/extract_llm_terms.py input.xlsx \
  -s Sheet1 \
  -c A \
  --start-row 2 \
  -o output_source_terms.xlsx
```

## 输出工作表

- `Terms_Source_Dedup`
- `Extraction_Evidence`
- `Conflicts_To_Review`
- `Import_Candidate`
- `Review_Before_Import`
- `Already_In_History`
- `Summary`
```

- [ ] **Step 2: Update root README**

Add a new tool section after `术语对检查`:

```markdown
### 2. LLM 术语提取

- 目录：`tools/llm_term_extractor`
- 用途：从没有 mark 的 Excel `source` 文案中使用 Codex 提取术语；`target` 有内容时抽取已有译法并复核冲突
- 默认模型：`gpt-5.3-codex-spark`
- Prompt：默认从 `tools/llm_term_extractor/prompts/` 读取，可通过 CLI 覆盖
- 输出方式：生成新的结果 Excel，不覆盖原文件
- CLI：

```bash
python3 tools/llm_term_extractor/extract_llm_terms.py input.xlsx \
  -s Sheet1 \
  -c A \
  -t B \
  --batch-size 50 \
  -o output_llm_terms.xlsx
```

详情见 `tools/llm_term_extractor/README.md`。
```

Renumber later sections or keep existing numbering style if the README already tolerates approximate numbering.

- [ ] **Step 3: Update CLI usage docs**

Add to `docs/cli-usage.md`:

```markdown
### LLM 术语提取

推荐完整调用：

```bash
python3 tools/llm_term_extractor/extract_llm_terms.py ./input.xlsx \
  -s Sheet1 \
  -c A \
  -t B \
  --start-row 2 \
  --batch-size 50 \
  --codex-model gpt-5.3-codex-spark \
  --codex-reasoning-effort high \
  -o ./artifacts/input_llm_terms.xlsx
```

只收集 source 术语时不传 `-t/--target-column`：

```bash
python3 tools/llm_term_extractor/extract_llm_terms.py ./input.xlsx \
  -s Sheet1 \
  -c A \
  --start-row 2 \
  -o ./artifacts/input_source_terms.xlsx
```
```

- [ ] **Step 4: Run focused verification**

Run:

```bash
python3 -m unittest tests.test_llm_term_extractor tests.test_gui_excel_selection
```

Expected: all listed tests pass.

- [ ] **Step 5: Run related regression tests**

Run:

```bash
python3 -m unittest tests.test_term_pair_checker tests.test_term_glossary_checker tests.test_workflow_runner tests.test_gui_excel_selection tests.test_llm_term_extractor
```

Expected: all listed tests pass.

- [ ] **Step 6: Run a no-Codex smoke test through injected tests only**

Run:

```bash
python3 -m unittest tests.test_llm_term_extractor.LlmTermWorkbookTests
```

Expected: workbook output sheets are created from fake LLM responses, proving Excel writing works without Codex.

- [ ] **Step 7: Commit**

```bash
git add tools/llm_term_extractor/README.md README.md docs/cli-usage.md
git commit -m "docs: document LLM term extractor"
```

## Self-Review

- Spec coverage: Task 1 covers external prompts, prompt rendering, Codex command construction, and JSON parsing. Task 2 covers source-only/source-target aggregation and output sheets. Task 3 covers history TB, Codex retry entry points, raw output retention, CLI defaults, and mixed-mode processing. Task 4 covers GUI and Toolshub registration. Task 5 covers docs and verification.
- Deferred-detail scan: the plan contains concrete file paths, function names, commands, expected failures, and implementation snippets for each code-producing task.
- Type consistency: `InputBatchRow`, `ExtractedLlmTerm`, `RowExtraction`, `ConflictGroup`, `ConflictDecision`, `ExtractionObservation`, `AggregatedTerm`, and `LlmTermExtractionSummary` are introduced before later tasks use them.
