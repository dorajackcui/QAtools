from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass


TAG_PLACEHOLDER_PREFIX = "t"
TAG_OPEN = "op"
TAG_CLOSE = "cl"
TAG_SELF = "sf"

TAG_PLACEHOLDER_RE = re.compile(r"\{t([1-9]\d*)_(op|cl|sf)\}")
_TAG_PLACEHOLDER_FULL_RE = re.compile(r"^\{t([1-9]\d*)_(op|cl|sf)\}$")
_TAG_SPAN_RE = re.compile(
    r"</[A-Za-z][A-Za-z0-9:._-]*\s*>"
    r"|</>"
    r"|<[A-Za-z][A-Za-z0-9:._-]*(?:\s+[^<>]*)?/\s*>"
    r"|<[A-Za-z][A-Za-z0-9:._-]*(?:\s+[^<>]*)?>"
    r"|\[/\]"
    r"|\[[A-Za-z][A-Za-z0-9:._-]*(?:=[^\[\]]+)?\]"
)


@dataclass(frozen=True)
class TagToken:
    index: int
    kind: str
    placeholder: str
    raw: str


@dataclass(frozen=True)
class TagExtraction:
    text: str
    tags: tuple[TagToken, ...]
    warnings: tuple[str, ...] = ()


@dataclass(frozen=True)
class TagValidation:
    warnings: tuple[str, ...]


def make_tag_placeholder(index: int, kind: str) -> str:
    if index < 1:
        raise ValueError("tag placeholder index must be >= 1")
    if kind not in {TAG_OPEN, TAG_CLOSE, TAG_SELF}:
        raise ValueError(f"unknown tag placeholder kind: {kind}")
    return f"{{{TAG_PLACEHOLDER_PREFIX}{index}_{kind}}}"


def parse_tag_placeholder(placeholder: str) -> tuple[int, str] | None:
    found = _TAG_PLACEHOLDER_FULL_RE.match(placeholder)
    if not found:
        return None
    return int(found.group(1)), found.group(2)


def is_tag_placeholder(placeholder: str) -> bool:
    return parse_tag_placeholder(placeholder) is not None


def extract_tags(text: str) -> TagExtraction:
    source = "" if text is None else str(text)
    warnings = [
        f"reserved_tag_placeholder: {found.group(0)}"
        for found in TAG_PLACEHOLDER_RE.finditer(source)
    ]
    chunks: list[str] = []
    tags: list[TagToken] = []
    stack: list[tuple[int, str | None]] = []
    next_index = 1
    pos = 0

    for found in _TAG_SPAN_RE.finditer(source):
        raw = found.group(0)
        chunks.append(source[pos : found.start()])
        kind, name = _classify_raw_tag(raw)

        if kind == TAG_OPEN:
            index = next_index
            next_index += 1
            stack.append((index, name))
            placeholder = make_tag_placeholder(index, TAG_OPEN)
        elif kind == TAG_SELF:
            index = next_index
            next_index += 1
            placeholder = make_tag_placeholder(index, TAG_SELF)
        else:
            matched = _pop_matching_open(stack, name)
            if matched is None:
                chunks.append(raw)
                warnings.append(f"unpaired_close_tag: {raw}")
                pos = found.end()
                continue
            index = matched
            placeholder = make_tag_placeholder(index, TAG_CLOSE)

        tags.append(TagToken(index, kind, placeholder, raw))
        chunks.append(placeholder)
        pos = found.end()

    chunks.append(source[pos:])

    for index, _name in stack:
        warnings.append(f"unclosed_open_tag: {make_tag_placeholder(index, TAG_OPEN)}")

    return TagExtraction("".join(chunks), tuple(tags), tuple(warnings))


def is_tag_only_segment(source: str) -> bool:
    text = "" if source is None else str(source)
    stripped = text.strip()
    if not stripped:
        return False
    remainder = TAG_PLACEHOLDER_RE.sub("", stripped)
    return remainder.strip() == ""


def restore_tags(text: str, tags: tuple[TagToken, ...]) -> str:
    result = "" if text is None else str(text)
    by_placeholder = {tag.placeholder: tag.raw for tag in tags}
    for placeholder, raw in by_placeholder.items():
        result = result.replace(placeholder, raw)
    return result


def validate_tag_placeholders(text: str, tags: tuple[TagToken, ...]) -> TagValidation:
    target_counts = Counter(TAG_PLACEHOLDER_RE.findall("" if text is None else str(text)))
    source_counts = Counter((str(tag.index), tag.kind) for tag in tags)
    warnings: list[str] = []

    for key in sorted(source_counts, key=_placeholder_sort_key):
        missing = source_counts[key] - target_counts[key]
        warnings.extend(
            f"tag_mismatch: missing {make_tag_placeholder(int(key[0]), key[1])}"
            for _ in range(missing)
        )

    for key in sorted(target_counts, key=_placeholder_sort_key):
        extra = target_counts[key] - source_counts[key]
        warnings.extend(
            f"tag_mismatch: extra {make_tag_placeholder(int(key[0]), key[1])}"
            for _ in range(extra)
        )

    return TagValidation(tuple(warnings))


def serialize_known_tags(text: str, tags: tuple[TagToken, ...]) -> TagExtraction:
    source = "" if text is None else str(text)
    if source == "":
        return TagExtraction("", ())

    result = source
    found_tags: list[TagToken] = []
    warnings: list[str] = []

    for tag in sorted(tags, key=lambda item: len(item.raw), reverse=True):
        if tag.raw in result:
            result = result.replace(tag.raw, tag.placeholder)
            found_tags.append(tag)
        else:
            warnings.append(f"source_tag_not_found: {tag.raw}")

    return TagExtraction(result, tuple(found_tags), tuple(warnings))


def _classify_raw_tag(raw: str) -> tuple[str, str | None]:
    if raw.startswith("</"):
        return TAG_CLOSE, None if raw == "</>" else raw[2:-1].strip().lower()
    if raw == "[/]":
        return TAG_CLOSE, None
    if raw.startswith("<") and re.search(r"/\s*>$", raw):
        return TAG_SELF, _angle_tag_name(raw)
    if raw.startswith("<"):
        return TAG_OPEN, _angle_tag_name(raw)
    return TAG_OPEN, _bbcode_tag_name(raw)


def _angle_tag_name(raw: str) -> str:
    match = re.match(r"</?\s*([A-Za-z][A-Za-z0-9:._-]*)", raw)
    return match.group(1).lower() if match else ""


def _bbcode_tag_name(raw: str) -> str:
    match = re.match(r"\[([A-Za-z][A-Za-z0-9:._-]*)", raw)
    return match.group(1).lower() if match else ""


def _pop_matching_open(stack: list[tuple[int, str | None]], name: str | None) -> int | None:
    if not stack:
        return None
    if name is None:
        return stack.pop()[0]
    for offset in range(len(stack) - 1, -1, -1):
        index, open_name = stack[offset]
        if open_name == name:
            del stack[offset]
            return index
    return None


def _placeholder_sort_key(key: tuple[str, str]) -> tuple[int, int]:
    kind_order = {TAG_OPEN: 0, TAG_CLOSE: 1, TAG_SELF: 2}
    return int(key[0]), kind_order[key[1]]


__all__ = [
    "TAG_PLACEHOLDER_PREFIX",
    "TAG_OPEN",
    "TAG_CLOSE",
    "TAG_SELF",
    "TAG_PLACEHOLDER_RE",
    "TagToken",
    "TagExtraction",
    "TagValidation",
    "make_tag_placeholder",
    "parse_tag_placeholder",
    "is_tag_placeholder",
    "extract_tags",
    "is_tag_only_segment",
    "restore_tags",
    "validate_tag_placeholders",
    "serialize_known_tags",
]
