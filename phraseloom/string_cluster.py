from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from difflib import SequenceMatcher
import re

from .template_engine import is_candidate_template, parse_template


@dataclass(frozen=True)
class SimilarStringCluster:
    group_id: str
    source_pattern: str
    member_indexes: tuple[int, ...]


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

    graph: dict[int, set[int]] = defaultdict(set)
    available = [index for index in range(len(sources)) if index not in assigned]
    for position, left_index in enumerate(available):
        for right_index in available[position + 1 :]:
            if not _structurally_similar(
                sources[left_index],
                sources[right_index],
                min_confidence=min_confidence,
            ):
                continue
            graph[left_index].add(right_index)
            graph[right_index].add(left_index)

    visited: set[int] = set()
    for start in available:
        if start in visited or start not in graph:
            continue
        stack = [start]
        component: list[int] = []
        while stack:
            index = stack.pop()
            if index in visited:
                continue
            visited.add(index)
            component.append(index)
            stack.extend(graph[index] - visited)
        members = tuple(sorted(component))
        if len(members) >= min_group_size:
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
    normalized_left = _normalize(left)
    normalized_right = _normalize(right)
    if not normalized_left or normalized_left == normalized_right:
        return False

    matcher = SequenceMatcher(
        None,
        normalized_left,
        normalized_right,
        autojunk=False,
    )
    longest = max(matcher.get_matching_blocks(), key=lambda block: block.size)
    shared = normalized_left[longest.a : longest.a + longest.size]
    shared_signal = _signal(shared)
    shorter_signal = min(
        len(_signal(normalized_left)),
        len(_signal(normalized_right)),
    )
    shared_ratio = len(shared_signal) / shorter_signal if shorter_signal else 0
    return (
        len(shared_signal) >= 4
        and shorter_signal > 0
        and (
            matcher.ratio() >= min_confidence
            or shared_ratio >= max(0.55, min_confidence - 0.15)
        )
    )


def _normalize(text: str) -> str:
    return re.sub(r"\s+", " ", str(text).strip().lower())


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
