#!/usr/bin/env python3
"""Shared term matching helpers for Excel terminology tools."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable

try:
    import ahocorasick
except ImportError:  # pragma: no cover - exercised via fallback path
    ahocorasick = None


SUPPORTED_MATCH_MODES = ("hybrid-boundary", "substring")


@dataclass(frozen=True)
class TermMappingEntry:
    source_term: str
    target_term: str
    normalized_source: str
    normalized_target: str


@dataclass(frozen=True)
class MatchSpan:
    start: int
    end: int
    entry: TermMappingEntry


def normalize_text(value: object, case_sensitive: bool) -> str:
    text = "" if value is None else str(value)
    return text if case_sensitive else text.casefold()


def is_ascii_word_char(character: str) -> bool:
    return character.isascii() and (character.isalnum() or character == "_")


def needs_left_boundary(term: str) -> bool:
    return bool(term) and term[0].isascii() and term[0].isalnum()


def needs_right_boundary(term: str) -> bool:
    return bool(term) and term[-1].isascii() and term[-1].isalnum()


def span_matches_mode(text: str, start: int, end: int, term: str, match_mode: str) -> bool:
    if match_mode == "substring":
        return True
    if match_mode != "hybrid-boundary":
        raise ValueError(f"不支持的匹配模式: {match_mode}")

    if needs_left_boundary(term) and start > 0 and is_ascii_word_char(text[start - 1]):
        return False
    if needs_right_boundary(term) and end < len(text) and is_ascii_word_char(text[end]):
        return False
    return True


def text_contains_term(text: str, term: str, match_mode: str) -> bool:
    if not term:
        return False

    search_start = 0
    while True:
        start_index = text.find(term, search_start)
        if start_index < 0:
            return False
        end_index = start_index + len(term)
        if span_matches_mode(text, start_index, end_index, term, match_mode):
            return True
        search_start = start_index + 1


def build_matcher(entries: Iterable[TermMappingEntry]):
    entries = list(entries)
    if ahocorasick is None:
        return entries

    automaton = ahocorasick.Automaton()
    for entry in entries:
        automaton.add_word(entry.normalized_source, entry)
    automaton.make_automaton()
    return automaton


def iter_match_spans(text: str, matcher, match_mode: str) -> list[MatchSpan]:
    spans: list[MatchSpan] = []

    if ahocorasick is not None and isinstance(matcher, ahocorasick.Automaton):
        for end_index, entry in matcher.iter(text):
            start_index = end_index - len(entry.normalized_source) + 1
            end_position = end_index + 1
            if span_matches_mode(text, start_index, end_position, entry.normalized_source, match_mode):
                spans.append(MatchSpan(start=start_index, end=end_position, entry=entry))
        return spans

    for entry in matcher:
        search_start = 0
        while True:
            start_index = text.find(entry.normalized_source, search_start)
            if start_index < 0:
                break
            end_index = start_index + len(entry.normalized_source)
            if span_matches_mode(text, start_index, end_index, entry.normalized_source, match_mode):
                spans.append(MatchSpan(start=start_index, end=end_index, entry=entry))
            search_start = start_index + 1
    return spans


def select_longest_non_overlapping_matches(spans: Iterable[MatchSpan]) -> list[TermMappingEntry]:
    sorted_spans = sorted(
        spans,
        key=lambda item: (-(item.end - item.start), item.start, item.entry.normalized_source),
    )
    accepted: list[MatchSpan] = []
    for span in sorted_spans:
        if any(not (span.end <= existing.start or span.start >= existing.end) for existing in accepted):
            continue
        accepted.append(span)

    accepted.sort(key=lambda item: (item.start, item.end))
    unique_entries: list[TermMappingEntry] = []
    seen_sources: set[str] = set()
    for span in accepted:
        normalized_source = span.entry.normalized_source
        if normalized_source in seen_sources:
            continue
        seen_sources.add(normalized_source)
        unique_entries.append(span.entry)
    return unique_entries


def find_row_terms(
    source_text: object,
    matcher,
    case_sensitive: bool,
    match_mode: str = "hybrid-boundary",
) -> list[TermMappingEntry]:
    normalized_source_text = normalize_text(source_text, case_sensitive=case_sensitive)
    if not normalized_source_text:
        return []

    spans = iter_match_spans(normalized_source_text, matcher, match_mode=match_mode)
    return select_longest_non_overlapping_matches(spans)
