#!/usr/bin/env python3
"""Shared term matching helpers for Excel terminology tools."""

from __future__ import annotations

import functools
import re
from dataclasses import dataclass
from itertools import product
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
    is_variant: bool = False


SIMPLE_PLURAL_SUFFIX = "s"
WORD_PATTERN = re.compile(r"\w+", re.UNICODE)
LITERAL_WHITESPACE_ESCAPES = (
    ("\\r\\n", " "),
    ("\\n", " "),
    ("\\r", " "),
    ("\\t", " "),
)
WHITESPACE_PATTERN = re.compile(r"[\t\r\n\f\v\u00a0]+")
TARGET_VARIANT_LIMIT = 256


def normalize_text(value: object, case_sensitive: bool) -> str:
    text = "" if value is None else str(value)
    for literal_escape, replacement in LITERAL_WHITESPACE_ESCAPES:
        text = text.replace(literal_escape, replacement)
    text = WHITESPACE_PATTERN.sub(" ", text)
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


def simple_s_plural_variant(term: str) -> str:
    if not term or term.endswith(SIMPLE_PLURAL_SUFFIX):
        return ""
    if not (term[-1].isascii() and term[-1].isalpha()):
        return ""
    return f"{term}{SIMPLE_PLURAL_SUFFIX}"


def simple_plural_token_variant(token: str) -> str:
    if not token or token.endswith(SIMPLE_PLURAL_SUFFIX):
        return ""
    if not (token[-1].isascii() and token[-1].isalpha()):
        return ""
    if len(token) > 1 and token.endswith("y"):
        return f"{token[:-1]}ies"
    return f"{token}{SIMPLE_PLURAL_SUFFIX}"


def iter_simple_s_variants(term: str) -> tuple[str, ...]:
    plural_term = simple_s_plural_variant(term)
    if not plural_term:
        return (term,)
    return (term, plural_term)


def _has_ascii_alpha(term: str) -> bool:
    return any(c.isascii() and c.isalpha() for c in term)


def iter_source_match_variants(term: str) -> tuple[str, ...]:
    if not _has_ascii_alpha(term):
        return (term,)
    variants = list(iter_simple_s_variants(term))
    seen = set(variants)
    matches = list(WORD_PATTERN.finditer(term))
    for match in matches:
        token_variant = simple_plural_token_variant(match.group(0))
        if not token_variant:
            continue
        variant = f"{term[:match.start()]}{token_variant}{term[match.end():]}"
        if variant in seen:
            continue
        seen.add(variant)
        variants.append(variant)
    return tuple(variants)


def iter_source_match_variant_pairs(term: str) -> tuple[tuple[str, bool], ...]:
    return tuple((variant, variant != term) for variant in iter_source_match_variants(term))


def s_plural_token_variants(token: str) -> tuple[str, ...]:
    if not token or not (token[-1].isascii() and token[-1].isalpha()):
        return (token,)
    if len(token) <= 2:
        return (token,)

    variants: list[str] = [token]
    seen = {token}

    def add_variant(variant: str) -> None:
        if variant and variant not in seen:
            seen.add(variant)
            variants.append(variant)

    if token.endswith(SIMPLE_PLURAL_SUFFIX):
        if len(token) > 3 and token.endswith("ies"):
            add_variant(f"{token[:-3]}y")
        elif len(token) > 3 and not token.endswith(("is", "os", "ss", "us")):
            add_variant(token[:-1])
    else:
        add_variant(f"{token}{SIMPLE_PLURAL_SUFFIX}")
        if len(token) > 1 and token.endswith("y"):
            add_variant(f"{token[:-1]}ies")

    return tuple(variants)


@functools.lru_cache(maxsize=1024)
def iter_target_s_plural_variants(term: str) -> tuple[str, ...]:
    matches = list(WORD_PATTERN.finditer(term))
    if not matches:
        return (term,)

    literal_parts: list[str] = []
    variant_lists: list[tuple[str, ...]] = []
    last_index = 0
    for match in matches:
        literal_parts.append(term[last_index : match.start()])
        variant_lists.append(s_plural_token_variants(match.group(0)))
        last_index = match.end()
    literal_parts.append(term[last_index:])

    variants: list[str] = []
    seen: set[str] = set()
    for token_values in product(*variant_lists):
        parts: list[str] = []
        for index, token_value in enumerate(token_values):
            parts.append(literal_parts[index])
            parts.append(token_value)
        parts.append(literal_parts[-1])
        variant = "".join(parts)
        if variant in seen:
            continue
        seen.add(variant)
        variants.append(variant)
        if len(variants) >= TARGET_VARIANT_LIMIT:
            break
    return tuple(variants)


def text_contains_target_s_plural_variant(text: str, term: str, match_mode: str) -> bool:
    return any(
        text_contains_term(text, variant, match_mode=match_mode)
        for variant in iter_target_s_plural_variants(term)
    )


def text_contains_source_match_variant(text: str, term: str, match_mode: str) -> bool:
    return any(
        text_contains_term(text, variant, match_mode=match_mode)
        for variant in iter_source_match_variants(term)
    )


def term_has_expected_target(
    source_text: str,
    target_text: str,
    entry: TermMappingEntry,
    match_mode: str,
    allow_target_plural_variants: bool = False,
) -> bool:
    if not text_contains_source_match_variant(
        source_text,
        entry.normalized_source,
        match_mode=match_mode,
    ):
        return False

    if text_contains_term(target_text, entry.normalized_target, match_mode=match_mode):
        return True
    if not allow_target_plural_variants:
        return False
    return text_contains_target_s_plural_variant(
        target_text,
        entry.normalized_target,
        match_mode=match_mode,
    )


def build_matcher(entries: Iterable[TermMappingEntry]):
    entries = list(entries)
    if ahocorasick is None:
        return entries

    match_values_by_variant: dict[str, list[tuple[str, TermMappingEntry, bool]]] = {}
    for entry in entries:
        for source_variant, is_variant in iter_source_match_variant_pairs(entry.normalized_source):
            match_values_by_variant.setdefault(source_variant, []).append((source_variant, entry, is_variant))

    automaton = ahocorasick.Automaton()
    for source_variant, match_values in match_values_by_variant.items():
        automaton.add_word(source_variant, tuple(match_values))
    automaton.make_automaton()
    return automaton


def iter_match_spans(text: str, matcher, match_mode: str) -> list[MatchSpan]:
    spans: list[MatchSpan] = []

    if ahocorasick is not None and isinstance(matcher, ahocorasick.Automaton):
        for end_index, match_values in matcher.iter(text):
            for source_variant, entry, is_variant in match_values:
                start_index = end_index - len(source_variant) + 1
                end_position = end_index + 1
                if span_matches_mode(text, start_index, end_position, source_variant, match_mode):
                    spans.append(MatchSpan(start=start_index, end=end_position, entry=entry, is_variant=is_variant))
        return spans

    for entry in matcher:
        for source_variant, is_variant in iter_source_match_variant_pairs(entry.normalized_source):
            search_start = 0
            while True:
                start_index = text.find(source_variant, search_start)
                if start_index < 0:
                    break
                end_index = start_index + len(source_variant)
                if span_matches_mode(text, start_index, end_index, source_variant, match_mode):
                    spans.append(MatchSpan(start=start_index, end=end_index, entry=entry, is_variant=is_variant))
                search_start = start_index + 1
    return spans


def select_longest_non_overlapping_matches(spans: Iterable[MatchSpan]) -> list[TermMappingEntry]:
    sorted_spans = sorted(
        spans,
        key=lambda item: (-(item.end - item.start), item.start, item.is_variant, item.entry.normalized_source),
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
