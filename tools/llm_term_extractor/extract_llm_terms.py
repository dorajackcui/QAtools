#!/usr/bin/env python3
"""Workbook aggregation for LLM-assisted term extraction."""

from __future__ import annotations

import re
from collections.abc import Callable, Iterable
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

from tools.llm_term_extractor.codex_term_review import (
    ConflictDecision,
    ConflictGroup,
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

BatchExtractor = Callable[[list[InputBatchRow]], Iterable[RowExtraction]]
ConflictReviewer = Callable[[list[ConflictGroup]], Any]


@dataclass(frozen=True)
class SourceRow:
    row_index: int
    source_text: str
    target_text: str

    @property
    def row_id(self) -> str:
        return str(self.row_index)


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
    term_types: list[str] = field(default_factory=list)
    confidences: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    conflict_decision: Any | None = None


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


def normalize_column(column_name: str) -> str:
    normalized = column_name.strip().upper()
    column_index_from_string(normalized)
    return normalized


def normalize_term_key(value: str) -> str:
    return re.sub(r"\s+", " ", value.strip()).casefold()


def build_default_output_path(input_file: str | Path) -> Path:
    input_path = Path(input_file).expanduser().absolute()
    return input_path.with_name(f"{input_path.stem}_llm_terms.xlsx")


def unique_join(values: Iterable[object], separator: str = "、") -> str:
    seen: set[str] = set()
    unique_values: list[str] = []
    for value in values:
        text = str(value).strip()
        if not text or text in seen:
            continue
        seen.add(text)
        unique_values.append(text)
    return separator.join(unique_values)


def iter_batches(rows: list[SourceRow], batch_size: int) -> Iterable[list[SourceRow]]:
    if batch_size < 1:
        raise ValueError("batch_size must be at least 1.")
    for start_index in range(0, len(rows), batch_size):
        yield rows[start_index : start_index + batch_size]


def _cell_text(value: object) -> str:
    return "" if value is None else str(value).strip()


def _read_source_rows(
    worksheet: Any,
    source_column: str,
    target_column: str,
    start_row: int,
) -> list[SourceRow]:
    rows: list[SourceRow] = []
    for row_index in range(start_row, worksheet.max_row + 1):
        source_text = _cell_text(worksheet[f"{source_column}{row_index}"].value)
        target_text = (
            _cell_text(worksheet[f"{target_column}{row_index}"].value)
            if target_column
            else ""
        )
        if not source_text and not target_text:
            continue
        rows.append(
            SourceRow(
                row_index=row_index,
                source_text=source_text,
                target_text=target_text,
            )
        )
    return rows


def _to_input_batch(rows: list[SourceRow]) -> list[InputBatchRow]:
    return [
        InputBatchRow(
            row_id=row.row_id,
            source_text=row.source_text,
            target_text=row.target_text,
        )
        for row in rows
    ]


def _value_from(value: object, *names: str) -> str:
    if isinstance(value, dict):
        for name in names:
            text = _cell_text(value.get(name))
            if text:
                return text
        return ""
    for name in names:
        text = _cell_text(getattr(value, name, ""))
        if text:
            return text
    return ""


def _row_extraction_id(row_extraction: object) -> str:
    row_id = _value_from(row_extraction, "row_id", "id")
    if row_id:
        return row_id
    row_index = _value_from(row_extraction, "row_index")
    return row_index


def collect_observations(
    rows_by_id: dict[str, SourceRow],
    extractions: Iterable[RowExtraction],
) -> list[ExtractionObservation]:
    observations: list[ExtractionObservation] = []
    for row_extraction in extractions:
        source_row = rows_by_id.get(_row_extraction_id(row_extraction))
        if source_row is None:
            continue
        terms = getattr(row_extraction, "terms", ())
        if isinstance(row_extraction, dict):
            terms = row_extraction.get("terms", ())
        for term in terms:
            source_term = _value_from(term, "source_term", "source", "term")
            if not source_term:
                continue
            observations.append(
                ExtractionObservation(
                    row_index=source_row.row_index,
                    source_text=source_row.source_text,
                    target_text=source_row.target_text,
                    source_term=source_term,
                    target_term=_value_from(term, "target_term", "target", "target_expression"),
                    term_type=_value_from(term, "term_type", "category", "type"),
                    confidence=_value_from(term, "confidence"),
                    note=_value_from(term, "note", "reason"),
                )
            )
    return observations


def _append_unique(values: list[str], value: str) -> None:
    if value and value not in values:
        values.append(value)


def aggregate_observations(observations: list[ExtractionObservation]) -> dict[str, AggregatedTerm]:
    aggregated: dict[str, AggregatedTerm] = {}
    for observation in observations:
        source_key = normalize_term_key(observation.source_term)
        if not source_key:
            continue
        term = aggregated.setdefault(
            source_key,
            AggregatedTerm(source_term=observation.source_term, source_key=source_key),
        )
        _append_unique(term.target_terms, observation.target_term)
        if observation.row_index not in term.row_indexes:
            term.row_indexes.append(observation.row_index)
        _append_unique(term.source_examples, observation.source_text)
        _append_unique(term.target_examples, observation.target_text)
        _append_unique(term.term_types, observation.term_type)
        _append_unique(term.confidences, observation.confidence)
        _append_unique(term.notes, observation.note)
    return aggregated


def _ordered_terms(aggregated_terms: dict[str, AggregatedTerm]) -> list[AggregatedTerm]:
    return sorted(
        aggregated_terms.values(),
        key=lambda term: (
            min(term.row_indexes) if term.row_indexes else 0,
            normalize_term_key(term.source_term),
        ),
    )


def _build_conflict_groups(aggregated_terms: dict[str, AggregatedTerm]) -> list[ConflictGroup]:
    return [
        ConflictGroup(
            group_id=term.source_key,
            source_term=term.source_term,
            target_terms=tuple(term.target_terms),
            examples=tuple(term.source_examples),
        )
        for term in _ordered_terms(aggregated_terms)
        if len(term.target_terms) > 1
    ]


def _iter_decision_items(decisions: object) -> Iterable[tuple[str, object]]:
    if decisions is None:
        return []
    if isinstance(decisions, dict):
        return [(str(key), value) for key, value in decisions.items()]
    return [("", decision) for decision in decisions]


def _apply_conflict_decisions(
    aggregated_terms: dict[str, AggregatedTerm],
    groups: list[ConflictGroup],
    conflict_reviewer: ConflictReviewer | None,
) -> None:
    if not groups or conflict_reviewer is None:
        return

    decisions = conflict_reviewer(groups)
    group_to_key = {group.group_id: group.group_id for group in groups}
    group_to_key.update({group.source_term: group.group_id for group in groups})

    for fallback_key, decision in _iter_decision_items(decisions):
        decision_group_id = _value_from(decision, "group_id", "source_term", "source")
        lookup_key = decision_group_id or fallback_key
        source_key = group_to_key.get(lookup_key) or normalize_term_key(lookup_key)
        term = aggregated_terms.get(source_key)
        if term is not None:
            term.conflict_decision = decision


def _decision_value(decision: object | None, *names: str) -> str:
    if decision is None:
        return ""
    return _value_from(decision, *names)


def _decision_status(decision: object | None) -> str:
    return _decision_value(decision, "decision").casefold()


def _effective_target(term: AggregatedTerm) -> str:
    canonical_target = _decision_value(term.conflict_decision, "canonical_target", "preferred_target")
    if canonical_target:
        return canonical_target
    if len(term.target_terms) == 1:
        return term.target_terms[0]
    return ""


def _rebuild_output_sheet(workbook: Any, data_sheet_name: str, sheet_name: str) -> Any:
    if data_sheet_name == sheet_name:
        raise ValueError(f"Data worksheet cannot be named {sheet_name}.")
    if sheet_name in workbook.sheetnames:
        del workbook[sheet_name]
    return workbook.create_sheet(title=sheet_name)


def _write_output_sheets(
    workbook: Any,
    worksheet_title: str,
    observations: list[ExtractionObservation],
    aggregated_terms: dict[str, AggregatedTerm],
    summary_values: dict[str, object],
) -> tuple[int, int, int, int]:
    terms_sheet = _rebuild_output_sheet(workbook, worksheet_title, TERMS_SHEET_NAME)
    terms_sheet.append(
        [
            "source_term",
            "target_terms_observed",
            "row_count",
            "rows",
            "source_examples",
            "target_examples",
            "term_types",
            "confidences",
            "notes",
            "decision",
            "decision_reason",
        ]
    )

    evidence_sheet = _rebuild_output_sheet(workbook, worksheet_title, EVIDENCE_SHEET_NAME)
    evidence_sheet.append(
        [
            "row_index",
            "source_term",
            "target_term",
            "term_type",
            "confidence",
            "note",
            "source_text",
            "target_text",
        ]
    )

    conflicts_sheet = _rebuild_output_sheet(workbook, worksheet_title, CONFLICTS_SHEET_NAME)
    conflicts_sheet.append(
        [
            "source_term",
            "target_terms_observed",
            "rows",
            "decision",
            "canonical_target",
            "reason",
            "source_examples",
            "target_examples",
            "notes",
        ]
    )

    import_sheet = _rebuild_output_sheet(workbook, worksheet_title, IMPORT_SHEET_NAME)
    import_sheet.append(["source_term", "target_term", "rows", "row_count", "notes"])

    review_sheet = _rebuild_output_sheet(workbook, worksheet_title, REVIEW_SHEET_NAME)
    review_sheet.append(
        [
            "source_term",
            "target_terms_observed",
            "rows",
            "reason",
            "decision",
            "decision_reason",
            "notes",
        ]
    )

    history_sheet = _rebuild_output_sheet(workbook, worksheet_title, HISTORY_SHEET_NAME)
    history_sheet.append(["source_term", "history_target", "rows", "target_terms_observed"])

    for observation in observations:
        evidence_sheet.append(
            [
                observation.row_index,
                observation.source_term,
                observation.target_term,
                observation.term_type,
                observation.confidence,
                observation.note,
                observation.source_text,
                observation.target_text,
            ]
        )

    conflict_count = 0
    import_candidate_count = 0
    review_before_import_count = 0
    already_in_history_count = 0

    for term in _ordered_terms(aggregated_terms):
        target_terms = unique_join(term.target_terms)
        rows = unique_join(term.row_indexes)
        notes = unique_join(term.notes, separator=" | ")
        decision = _decision_status(term.conflict_decision)
        decision_reason = _decision_value(term.conflict_decision, "reason", "note")
        canonical_target = _decision_value(
            term.conflict_decision,
            "canonical_target",
            "preferred_target",
        )

        terms_sheet.append(
            [
                term.source_term,
                target_terms,
                len(term.row_indexes),
                rows,
                unique_join(term.source_examples, separator=" | "),
                unique_join(term.target_examples, separator=" | "),
                unique_join(term.term_types),
                unique_join(term.confidences),
                notes,
                decision,
                decision_reason,
            ]
        )

        if decision in {"conflict", "review"}:
            conflict_count += 1
            review_before_import_count += 1
            conflicts_sheet.append(
                [
                    term.source_term,
                    target_terms,
                    rows,
                    decision,
                    canonical_target,
                    decision_reason,
                    unique_join(term.source_examples, separator=" | "),
                    unique_join(term.target_examples, separator=" | "),
                    notes,
                ]
            )
            review_sheet.append(
                [
                    term.source_term,
                    target_terms,
                    rows,
                    decision,
                    decision,
                    decision_reason,
                    notes,
                ]
            )
            continue

        if len(term.target_terms) > 1 and term.conflict_decision is None:
            conflict_count += 1
            review_before_import_count += 1
            conflicts_sheet.append(
                [
                    term.source_term,
                    target_terms,
                    rows,
                    "review",
                    "",
                    "多译法需确认",
                    unique_join(term.source_examples, separator=" | "),
                    unique_join(term.target_examples, separator=" | "),
                    notes,
                ]
            )
            review_sheet.append(
                [
                    term.source_term,
                    target_terms,
                    rows,
                    "多译法需确认",
                    "",
                    "",
                    notes,
                ]
            )
            continue

        effective_target = _effective_target(term)
        if effective_target and (decision != "same" or len(term.target_terms) == 1 or canonical_target):
            import_candidate_count += 1
            import_sheet.append(
                [
                    term.source_term,
                    effective_target,
                    rows,
                    len(term.row_indexes),
                    notes,
                ]
            )
            continue

        review_before_import_count += 1
        reason = "target缺失" if not term.target_terms else "多译法需确认"
        review_sheet.append([term.source_term, target_terms, rows, reason, decision, decision_reason, notes])

    summary_sheet = _rebuild_output_sheet(workbook, worksheet_title, SUMMARY_SHEET_NAME)
    summary_sheet.append(["metric", "value"])
    for key, value in summary_values.items():
        summary_sheet.append([key, str(value) if isinstance(value, Path) else value])
    summary_sheet.append(["conflict_count", conflict_count])
    summary_sheet.append(["import_candidate_count", import_candidate_count])
    summary_sheet.append(["review_before_import_count", review_before_import_count])
    summary_sheet.append(["already_in_history_count", already_in_history_count])

    return (
        conflict_count,
        import_candidate_count,
        review_before_import_count,
        already_in_history_count,
    )


def process_excel(
    input_file: str | Path,
    source_column: str,
    target_column: str | None = None,
    sheet: str | None = None,
    start_row: int = 2,
    batch_size: int = 50,
    output_file: str | Path | None = None,
    batch_extractor: BatchExtractor | None = None,
    conflict_reviewer: ConflictReviewer | None = None,
    history_tb_file: str | Path | None = None,
) -> LlmTermExtractionSummary:
    if start_row < 1:
        raise ValueError("start_row must be at least 1.")
    if batch_extractor is None:
        raise NotImplementedError("Default Codex extraction is not wired until Task 3.")

    input_path = Path(input_file).expanduser().absolute()
    if not input_path.exists():
        raise FileNotFoundError(f"Input file does not exist: {input_path}")

    source_column = normalize_column(source_column)
    normalized_target_column = normalize_column(target_column) if target_column else ""
    output_path = (
        Path(output_file).expanduser().absolute()
        if output_file
        else build_default_output_path(input_path)
    )

    workbook = load_workbook(input_path)
    worksheet = workbook[sheet] if sheet else workbook.active
    source_rows = _read_source_rows(
        worksheet=worksheet,
        source_column=source_column,
        target_column=normalized_target_column,
        start_row=start_row,
    )

    observations: list[ExtractionObservation] = []
    batch_count = 0
    for batch in iter_batches(source_rows, batch_size):
        batch_count += 1
        rows_by_id = {row.row_id: row for row in batch}
        extractions = batch_extractor(_to_input_batch(batch))
        observations.extend(collect_observations(rows_by_id, extractions or []))

    aggregated_terms = aggregate_observations(observations)
    conflict_groups = _build_conflict_groups(aggregated_terms)
    _apply_conflict_decisions(aggregated_terms, conflict_groups, conflict_reviewer)

    summary_values: dict[str, object] = {
        "output_path": output_path,
        "worksheet_title": worksheet.title,
        "source_column": source_column,
        "target_column": normalized_target_column,
        "start_row": start_row,
        "scanned_row_count": len(source_rows),
        "batch_count": batch_count,
        "term_count": len(aggregated_terms),
        "evidence_count": len(observations),
    }
    (
        conflict_count,
        import_candidate_count,
        review_before_import_count,
        already_in_history_count,
    ) = _write_output_sheets(
        workbook=workbook,
        worksheet_title=worksheet.title,
        observations=observations,
        aggregated_terms=aggregated_terms,
        summary_values=summary_values,
    )
    workbook.save(output_path)

    return LlmTermExtractionSummary(
        output_path=output_path,
        worksheet_title=worksheet.title,
        source_column=source_column,
        target_column=normalized_target_column,
        start_row=start_row,
        scanned_row_count=len(source_rows),
        batch_count=batch_count,
        term_count=len(aggregated_terms),
        evidence_count=len(observations),
        conflict_count=conflict_count,
        import_candidate_count=import_candidate_count,
        review_before_import_count=review_before_import_count,
        already_in_history_count=already_in_history_count,
    )
