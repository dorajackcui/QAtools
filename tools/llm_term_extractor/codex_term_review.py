"""Prompt rendering and Codex subprocess helpers for LLM term extraction."""

from __future__ import annotations

import json
import re
import subprocess
from dataclasses import asdict, dataclass, is_dataclass
from pathlib import Path
from typing import Any


DEFAULT_CODEX_MODEL = "gpt-5.3-codex-spark"
DEFAULT_REASONING_EFFORT = "high"
DEFAULT_TIMEOUT_SECONDS = 600


@dataclass(frozen=True)
class InputBatchRow:
    row_id: str
    source_text: str
    target_text: str = ""


@dataclass(frozen=True)
class ExtractedLlmTerm:
    source_term: str
    target_term: str = ""
    category: str = ""
    note: str = ""


@dataclass(frozen=True)
class RowExtraction:
    row_id: str
    terms: tuple[ExtractedLlmTerm, ...]


@dataclass(frozen=True)
class ConflictGroup:
    group_id: str
    source_term: str
    target_terms: tuple[str, ...]
    examples: tuple[str, ...] = ()


@dataclass(frozen=True)
class ConflictDecision:
    group_id: str
    decision: str
    canonical_target: str = ""
    reason: str = ""


EXTRACTION_OUTPUT_SCHEMA = {
    "rows": [
        {
            "row_id": "input row id",
            "terms": [
                {
                    "source_term": "term text from source",
                    "target_term": "existing target expression when target text exists, otherwise empty string",
                    "category": "name | event | gameplay | item | currency | system | title | fixed_phrase | cross_context_anchor | other",
                    "note": "short reason",
                }
            ],
        }
    ]
}

CONFLICT_OUTPUT_SCHEMA = {
    "decisions": [
        {
            "group_id": "input group id",
            "decision": "same | conflict | review",
            "canonical_target": "preferred existing target if clear, otherwise empty string",
            "reason": "short reason",
        }
    ]
}


def _cell_text(value: object) -> str:
    return "" if value is None else str(value).strip()


def _json_payload(items: Any) -> str:
    return json.dumps(_to_jsonable(items), ensure_ascii=False, indent=2)


def _to_jsonable(value: Any) -> Any:
    if is_dataclass(value) and not isinstance(value, type):
        return {key: _to_jsonable(item) for key, item in asdict(value).items()}
    if isinstance(value, dict):
        return {key: _to_jsonable(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_to_jsonable(item) for item in value]
    return value


def load_prompt_template(path: str | Path) -> str:
    return Path(path).read_text(encoding="utf-8")


def render_extraction_prompt(template: str, *, mode: str, rows: list[InputBatchRow]) -> str:
    return (
        template.replace("{{MODE}}", mode)
        .replace("{{ROWS_JSON}}", _json_payload({"rows": rows}))
        .replace("{{OUTPUT_SCHEMA}}", _json_payload(EXTRACTION_OUTPUT_SCHEMA))
    )


def render_conflict_prompt(template: str, *, groups: list[ConflictGroup]) -> str:
    return template.replace("{{GROUPS_JSON}}", _json_payload({"groups": groups})).replace(
        "{{OUTPUT_SCHEMA}}",
        _json_payload(CONFLICT_OUTPUT_SCHEMA),
    )


def extract_json_object(text: str) -> dict[str, Any]:
    stripped = text.strip()
    fenced_match = re.search(r"```(?:json)?\s*(.*?)\s*```", stripped, flags=re.DOTALL)
    if fenced_match:
        stripped = fenced_match.group(1).strip()
    if not stripped.startswith("{"):
        start = stripped.find("{")
        end = stripped.rfind("}")
        if start == -1 or end == -1 or end < start:
            raise ValueError("Codex output did not contain a JSON object.")
        stripped = stripped[start : end + 1]
    parsed = json.loads(stripped, strict=False)
    if not isinstance(parsed, dict):
        raise ValueError("Codex output JSON is not an object.")
    return parsed


def parse_extraction_response(text: str) -> list[RowExtraction]:
    parsed = extract_json_object(text)
    rows = parsed.get("rows")
    if not isinstance(rows, list):
        raise ValueError("Codex extraction output is missing a rows array.")

    parsed_rows: list[RowExtraction] = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        terms = row.get("terms")
        if not isinstance(terms, list):
            terms = []
        parsed_terms = tuple(
            ExtractedLlmTerm(
                source_term=_cell_text(term.get("source_term") or term.get("source") or term.get("term")),
                target_term=_cell_text(
                    term.get("target_term")
                    or term.get("target")
                    or term.get("target_expression")
                    or term.get("existing_target_expression")
                ),
                category=_cell_text(term.get("category") or term.get("type")),
                note=_cell_text(term.get("note") or term.get("reason")),
            )
            for term in terms
            if isinstance(term, dict)
        )
        parsed_rows.append(RowExtraction(row_id=_cell_text(row.get("row_id") or row.get("id")), terms=parsed_terms))
    return parsed_rows


def parse_conflict_response(text: str) -> list[ConflictDecision]:
    parsed = extract_json_object(text)
    decisions = parsed.get("decisions")
    if not isinstance(decisions, list):
        raise ValueError("Codex conflict output is missing a decisions array.")

    return [
        ConflictDecision(
            group_id=_cell_text(item.get("group_id") or item.get("id")),
            decision=_cell_text(item.get("decision")),
            canonical_target=_cell_text(item.get("canonical_target") or item.get("preferred_target")),
            reason=_cell_text(item.get("reason") or item.get("note")),
        )
        for item in decisions
        if isinstance(item, dict)
    ]


def build_codex_command(output_path: str | Path, model: str, reasoning_effort: str) -> list[str]:
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
    output_path: str | Path,
    model: str,
    reasoning_effort: str,
    timeout_seconds: int,
) -> str:
    completed = subprocess.run(
        build_codex_command(output_path, model, reasoning_effort),
        input=prompt,
        capture_output=True,
        text=True,
        timeout=timeout_seconds,
        check=False,
    )
    if completed.returncode != 0:
        raise RuntimeError(
            "Codex term review failed "
            f"with exit code {completed.returncode}.\n"
            f"stderr:\n{completed.stderr}\n"
            f"stdout:\n{completed.stdout}"
        )

    output_file = Path(output_path)
    return output_file.read_text(encoding="utf-8") if output_file.exists() else completed.stdout
