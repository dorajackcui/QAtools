from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from difflib import SequenceMatcher
import re

from .tag_engine import PROTECTED_TOKEN_RE
from .template_engine import PLACEHOLDER_RE, is_candidate_template, parse_template


@dataclass(frozen=True)
class SimilarStringCluster:
    group_id: str
    source_pattern: str
    member_indexes: tuple[int, ...]


@dataclass(frozen=True)
class _PreparedSource:
    normalized: str
    signal: str


def cluster_similar_strings(
    sources: list[str],
    *,
    min_group_size: int = 3,
    min_confidence: float = 0.7,
) -> list[SimilarStringCluster]:
    """Group structurally similar strings without merging or translating them."""
    if min_group_size < 2:
        raise ValueError("min_group_size must be at least 2")

    candidates: list[tuple[tuple[int, ...], str]] = []
    assigned: set[int] = set()
    template_groups: dict[str, list[int]] = defaultdict(list)
    for index, source in enumerate(sources):
        match = parse_template(source)
        if is_candidate_template(match):
            template_groups[match.template].append(index)
    for pattern, indexes in template_groups.items():
        members = tuple(dict.fromkeys(indexes))
        if len(members) >= min_group_size:
            candidates.append((members, pattern))
            assigned.update(members)

    available = [index for index in range(len(sources)) if index not in assigned]
    prepared = [_prepare(source) for source in sources]
    graph: dict[int, set[int]] = defaultdict(set)
    for position, left_index in enumerate(available):
        for right_index in available[position + 1 :]:
            if not _prepared_structurally_similar(
                prepared[left_index],
                prepared[right_index],
                min_confidence=min_confidence,
            ):
                continue
            graph[left_index].add(right_index)
            graph[right_index].add(left_index)

    for members in _complete_link_groups(
        graph,
        available,
        min_group_size=min_group_size,
    ):
        candidates.append(
            (
                members,
                _shared_pattern([sources[index] for index in members]),
            )
        )

    candidates.sort(key=lambda item: (min(item[0]), item[1]))
    return [
        SimilarStringCluster(
            group_id=f"G{index:03d}",
            source_pattern=pattern,
            member_indexes=members,
        )
        for index, (members, pattern) in enumerate(candidates, start=1)
    ]


def _structurally_similar(
    left: str,
    right: str,
    *,
    min_confidence: float,
) -> bool:
    return _prepared_structurally_similar(
        _prepare(left),
        _prepare(right),
        min_confidence=min_confidence,
    )


def _prepared_structurally_similar(
    left: _PreparedSource,
    right: _PreparedSource,
    *,
    min_confidence: float,
) -> bool:
    if (
        not left.normalized
        or not right.normalized
        or left.normalized == right.normalized
    ):
        return False

    shorter_signal = min(len(left.signal), len(right.signal))
    longer_signal = max(len(left.signal), len(right.signal))
    if not shorter_signal or not longer_signal:
        return False

    minimum_short_coverage = max(0.55, min_confidence - 0.15)
    minimum_long_coverage = max(0.35, min_confidence - 0.35)
    maximum_sequence_ratio = (2 * shorter_signal) / (
        shorter_signal + longer_signal
    )
    if (
        maximum_sequence_ratio < min_confidence
        and shorter_signal / longer_signal < minimum_long_coverage
    ):
        return False

    matcher = SequenceMatcher(
        None,
        left.normalized,
        right.normalized,
        autojunk=False,
    )
    longest = max(matcher.get_matching_blocks(), key=lambda block: block.size)
    shared = left.normalized[longest.a : longest.a + longest.size]
    shared_signal = _signal(shared)
    shared_short_coverage = len(shared_signal) / shorter_signal
    shared_long_coverage = len(shared_signal) / longer_signal
    return (
        len(shared_signal) >= 4
        and (
            matcher.ratio() >= min_confidence
            or (
                shared_short_coverage >= minimum_short_coverage
                and shared_long_coverage >= minimum_long_coverage
            )
        )
    )


def _complete_link_groups(
    graph: dict[int, set[int]],
    available: list[int],
    *,
    min_group_size: int,
) -> list[tuple[int, ...]]:
    """Build disjoint groups whose members are all mutually similar."""
    remaining = set(available)
    groups: list[tuple[int, ...]] = []
    for seed in available:
        if seed not in remaining:
            continue
        neighbors = graph.get(seed, set()) & remaining
        if len(neighbors) + 1 < min_group_size:
            continue

        members = [seed]
        ordered_neighbors = sorted(
            neighbors,
            key=lambda index: (
                -len(graph.get(index, set()) & neighbors),
                index,
            ),
        )
        for candidate in ordered_neighbors:
            if all(
                candidate in graph.get(member, set())
                for member in members
            ):
                members.append(candidate)

        if len(members) < min_group_size:
            continue
        group = tuple(sorted(members))
        groups.append(group)
        remaining.difference_update(group)
    return groups


def _prepare(text: str) -> _PreparedSource:
    normalized = _normalize(text)
    return _PreparedSource(normalized=normalized, signal=_signal(normalized))


def _normalize(text: str) -> str:
    without_tokens = PROTECTED_TOKEN_RE.sub(" ", str(text))
    without_tokens = PLACEHOLDER_RE.sub(" ", without_tokens)
    return re.sub(r"\s+", " ", without_tokens.strip().lower())


def _signal(text: str) -> str:
    return re.sub(r"[\W_]+", "", text, flags=re.UNICODE)


def _shared_pattern(sources: list[str]) -> str:
    if not sources:
        return ""
    prefix = sources[0]
    for source in sources[1:]:
        length = 0
        for left_char, right_char in zip(prefix, source):
            if left_char != right_char:
                break
            length += 1
        prefix = prefix[:length]

    suffix = sources[0][len(prefix) :]
    for source in sources[1:]:
        candidate = source[len(prefix) :]
        length = 0
        for left_char, right_char in zip(reversed(suffix), reversed(candidate)):
            if left_char != right_char:
                break
            length += 1
        suffix = suffix[len(suffix) - length :] if length else ""
    return f"{prefix}{{variant}}{suffix}"


__all__ = ["SimilarStringCluster", "cluster_similar_strings"]
