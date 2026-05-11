from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass


TAG_OPEN = "op"
TAG_CLOSE = "cl"
TAG_SELF = "sf"
RAW_PLACEHOLDER = "ph"
PROTECTED_SINGLE = "single"

PROTECTED_TOKEN_RE = re.compile(r"\{([1-9]\d*)>|<([1-9]\d*)\}|\{([1-9]\d*)\}")
_PROTECTED_TOKEN_FULL_RE = re.compile(
    r"^\{([1-9]\d*)>$|^<([1-9]\d*)\}$|^\{([1-9]\d*)\}$"
)
TAG_PLACEHOLDER_RE = PROTECTED_TOKEN_RE
_TAG_SPAN_RE = re.compile(
    r"</[A-Za-z][A-Za-z0-9:._-]*\s*>"
    r"|</>"
    r"|<[A-Za-z][A-Za-z0-9:._-]*(?:\s+[^<>]*)?/\s*>"
    r"|<[A-Za-z][A-Za-z0-9:._-]*(?:\s+[^<>]*)?>"
    r"|\[/\]"
    r"|\[/[A-Za-z][A-Za-z0-9:._-]*\]"
    r"|\[[A-Za-z][A-Za-z0-9:._-]*(?:=[^\[\]]+)?\]"
)
_RAW_BRACE_RE = re.compile(r"\{[^{}]+\}")
_PROTECTED_SPAN_RE = re.compile(_TAG_SPAN_RE.pattern + r"|" + _RAW_BRACE_RE.pattern)


@dataclass(frozen=True)
class TagToken:
    index: int
    kind: str
    placeholder: str
    raw: str
    partner_index: int | None = None


@dataclass(frozen=True)
class TagExtraction:
    text: str
    tags: tuple[TagToken, ...]
    warnings: tuple[str, ...] = ()


@dataclass(frozen=True)
class TagValidation:
    warnings: tuple[str, ...]


def make_protected_token(index: int, kind: str) -> str:
    if index < 1:
        raise ValueError("protected token index must be >= 1")
    if kind == TAG_OPEN:
        return f"{{{index}>"
    if kind == TAG_CLOSE:
        return f"<{index}}}"
    if kind in {TAG_SELF, RAW_PLACEHOLDER, PROTECTED_SINGLE}:
        return f"{{{index}}}"
    raise ValueError(f"unknown protected token kind: {kind}")


def parse_protected_token(token: str) -> tuple[int, str] | None:
    found = _PROTECTED_TOKEN_FULL_RE.match(token)
    if not found:
        return None
    if found.group(1):
        return int(found.group(1)), TAG_OPEN
    if found.group(2):
        return int(found.group(2)), TAG_CLOSE
    return int(found.group(3)), PROTECTED_SINGLE


def is_protected_token(token: str) -> bool:
    return parse_protected_token(token) is not None


def make_tag_placeholder(index: int, kind: str) -> str:
    return make_protected_token(index, kind)


def parse_tag_placeholder(placeholder: str) -> tuple[int, str] | None:
    return parse_protected_token(placeholder)


def is_tag_placeholder(placeholder: str) -> bool:
    return is_protected_token(placeholder)


def extract_tags(text: str) -> TagExtraction:
    source = "" if text is None else str(text)
    warnings: list[str] = []
    spans = list(_PROTECTED_SPAN_RE.finditer(source))
    matched_plain_bbcode_starts = _matched_plain_bbcode_open_starts(spans)
    chunks: list[str] = []
    tags: list[TagToken] = []
    stack: list[tuple[int, str | None]] = []
    next_index = 1
    pos = 0

    for found in spans:
        raw = found.group(0)
        chunks.append(source[pos : found.start()])
        if _RAW_BRACE_RE.fullmatch(raw):
            index = next_index
            next_index += 1
            placeholder = make_protected_token(index, RAW_PLACEHOLDER)
            tags.append(TagToken(index, RAW_PLACEHOLDER, placeholder, raw))
            chunks.append(placeholder)
            pos = found.end()
            continue

        if _is_unmatched_plain_bbcode_open(raw, found.start(), matched_plain_bbcode_starts):
            chunks.append(raw)
            pos = found.end()
            continue
        kind, name = _classify_raw_tag(raw)

        if kind == TAG_OPEN:
            index = next_index
            next_index += 1
            stack.append((index, name))
            placeholder = make_protected_token(index, TAG_OPEN)
            token = TagToken(index, TAG_OPEN, placeholder, raw)
        elif kind == TAG_SELF:
            index = next_index
            next_index += 1
            placeholder = make_protected_token(index, TAG_SELF)
            token = TagToken(index, TAG_SELF, placeholder, raw)
        else:
            matched = _pop_matching_open(stack, name)
            if matched is None:
                chunks.append(raw)
                warnings.append(f"unpaired close tag: {raw}")
                pos = found.end()
                continue
            index = next_index
            next_index += 1
            placeholder = make_protected_token(index, TAG_CLOSE)
            token = TagToken(index, TAG_CLOSE, placeholder, raw, matched)

        tags.append(token)
        chunks.append(placeholder)
        pos = found.end()

    chunks.append(source[pos:])

    for index, _name in stack:
        warnings.append(
            f"open tag has no close partner: {make_protected_token(index, TAG_OPEN)}"
        )

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
    target_counts = Counter(
        parsed
        for found in TAG_PLACEHOLDER_RE.finditer("" if text is None else str(text))
        if (parsed := parse_protected_token(found.group(0))) is not None
    )
    source_counts = Counter((tag.index, tag.kind) for tag in tags)
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

    for tag in tags:
        if tag.raw in result:
            result = result.replace(tag.raw, tag.placeholder, 1)
            found_tags.append(tag)
        else:
            warnings.append(f"source_tag_not_found: {tag.raw}")

    return TagExtraction(result, tuple(found_tags), tuple(warnings))


def _classify_raw_tag(raw: str) -> tuple[str, str | None]:
    if raw.startswith("</"):
        return TAG_CLOSE, None if raw == "</>" else raw[2:-1].strip().lower()
    if raw == "[/]":
        return TAG_CLOSE, None
    if raw.startswith("[/"):
        return TAG_CLOSE, raw[2:-1].strip().lower()
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
    index, open_name = stack[-1]
    if open_name == name:
        stack.pop()
        return index
    return None


def _matched_plain_bbcode_open_starts(spans: list[re.Match[str]]) -> set[int]:
    later_closes_by_name: dict[str, int] = {}
    matched_starts: set[int] = set()

    for found in reversed(spans):
        raw = found.group(0)
        if _is_named_bbcode_close(raw):
            name = raw[2:-1].strip().lower()
            later_closes_by_name[name] = later_closes_by_name.get(name, 0) + 1
        elif _is_plain_bbcode_open(raw):
            name = _bbcode_tag_name(raw)
            if later_closes_by_name.get(name, 0) > 0:
                matched_starts.add(found.start())
                later_closes_by_name[name] -= 1

    return matched_starts


def _is_unmatched_plain_bbcode_open(
    raw: str, start: int, matched_plain_bbcode_starts: set[int]
) -> bool:
    return _is_plain_bbcode_open(raw) and start not in matched_plain_bbcode_starts


def _is_plain_bbcode_open(raw: str) -> bool:
    return re.fullmatch(r"\[[A-Za-z][A-Za-z0-9:._-]*\]", raw) is not None


def _is_named_bbcode_close(raw: str) -> bool:
    return re.fullmatch(r"\[/[A-Za-z][A-Za-z0-9:._-]*\]", raw) is not None


def _placeholder_sort_key(key: tuple[int, str]) -> tuple[int, int]:
    kind_order = {TAG_OPEN: 0, TAG_CLOSE: 1, TAG_SELF: 2, PROTECTED_SINGLE: 2}
    return int(key[0]), kind_order[key[1]]


__all__ = [
    "RAW_PLACEHOLDER",
    "PROTECTED_SINGLE",
    "TAG_OPEN",
    "TAG_CLOSE",
    "TAG_SELF",
    "PROTECTED_TOKEN_RE",
    "TAG_PLACEHOLDER_RE",
    "TagToken",
    "TagExtraction",
    "TagValidation",
    "make_protected_token",
    "parse_protected_token",
    "is_protected_token",
    "make_tag_placeholder",
    "parse_tag_placeholder",
    "is_tag_placeholder",
    "extract_tags",
    "is_tag_only_segment",
    "restore_tags",
    "validate_tag_placeholders",
    "serialize_known_tags",
]
