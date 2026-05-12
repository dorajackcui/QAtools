from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass

from .tag_rules import TagRules, default_tag_rules


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
    r"|<[A-Za-z][A-Za-z0-9:._-]*(?:=[^<>/\s][^<>/]*)?(?:\s+[^<>]*)?/\s*>"
    r"|<[A-Za-z][A-Za-z0-9:._-]*(?:=[^<>\s][^<>]*)?(?:\s+[^<>]*)?>"
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


def extract_tags(text: str, rules: TagRules | None = None) -> TagExtraction:
    source = "" if text is None else str(text)
    active_rules = default_tag_rules() if rules is None else rules
    warnings: list[str] = []
    spans = list(_PROTECTED_SPAN_RE.finditer(source))
    matched_plain_bbcode_starts = _matched_plain_bbcode_open_starts(spans, active_rules)
    matched_optional_angle_starts = _matched_optional_pair_angle_open_starts(
        spans, active_rules
    )
    chunks: list[str] = []
    tags: list[TagToken] = []
    stack: list[tuple[int, str | None]] = []
    next_index = 1
    pos = 0

    for found in spans:
        raw = found.group(0)
        chunks.append(source[pos : found.start()])
        if _RAW_BRACE_RE.fullmatch(raw):
            if not active_rules.protect_raw_braces:
                chunks.append(raw)
                pos = found.end()
                continue
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
        if not _is_raw_tag_allowed(raw, kind, name, active_rules):
            chunks.append(raw)
            pos = found.end()
            continue
        if (
            kind == TAG_OPEN
            and raw.startswith("<")
            and _angle_open_should_be_single(
                name, found.start(), matched_optional_angle_starts, active_rules
            )
        ):
            kind = TAG_SELF

        if kind == TAG_OPEN:
            index = next_index
            next_index += 1
            stack.append((index, _canonical_raw_tag_name(raw, name, active_rules)))
            placeholder = make_protected_token(index, TAG_OPEN)
            token = TagToken(index, TAG_OPEN, placeholder, raw)
        elif kind == TAG_SELF:
            index = next_index
            next_index += 1
            placeholder = make_protected_token(index, TAG_SELF)
            token = TagToken(index, TAG_SELF, placeholder, raw)
        else:
            matched = _pop_matching_open(
                stack, _canonical_raw_tag_name(raw, name, active_rules)
            )
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
    remainder = PROTECTED_TOKEN_RE.sub("", stripped)
    return remainder.strip() == ""


def restore_tags(text: str, tags: tuple[TagToken, ...]) -> str:
    result = "" if text is None else str(text)
    by_placeholder = {tag.placeholder: tag.raw for tag in tags}
    for placeholder, raw in by_placeholder.items():
        result = result.replace(placeholder, raw)
    return result


def validate_tag_placeholders(text: str, tags: tuple[TagToken, ...]) -> TagValidation:
    source_counts = Counter(tag.placeholder for tag in tags)
    target_counts = Counter(
        found.group(0) for found in PROTECTED_TOKEN_RE.finditer("" if text is None else str(text))
    )
    warnings: list[str] = []

    for placeholder in sorted(source_counts, key=_protected_token_sort_key):
        missing = source_counts[placeholder] - target_counts[placeholder]
        warnings.extend(
            f"protected_token_mismatch: missing {placeholder}" for _ in range(missing)
        )

    for placeholder in sorted(target_counts, key=_protected_token_sort_key):
        extra = target_counts[placeholder] - source_counts[placeholder]
        warnings.extend(
            f"protected_token_mismatch: extra {placeholder}" for _ in range(extra)
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
            warnings.append(f"source_protected_span_not_found: {tag.raw}")

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


def _is_raw_tag_allowed(raw: str, kind: str, name: str | None, rules: TagRules) -> bool:
    if raw.startswith("<"):
        if raw == "</>":
            return True
        return name is not None and rules.allows_angle(name)
    if raw == "[/]":
        return True
    return name is not None and rules.allows_bbcode(name)


def _angle_open_should_be_single(
    name: str | None,
    start: int,
    matched_optional_angle_starts: set[int],
    rules: TagRules,
) -> bool:
    if name is None:
        return False
    if rules.is_angle_single(name):
        return True
    return rules.is_angle_optional_pair(name) and start not in matched_optional_angle_starts


def _canonical_raw_tag_name(
    raw: str, name: str | None, rules: TagRules
) -> str | None:
    if name is None:
        return None
    if raw.startswith("<"):
        return rules.canonical_angle(name)
    return name.lower()


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


def _matched_plain_bbcode_open_starts(
    spans: list[re.Match[str]], rules: TagRules
) -> set[int]:
    later_closes_by_name: dict[str, int] = {}
    matched_starts: set[int] = set()

    for found in reversed(spans):
        raw = found.group(0)
        if _is_named_bbcode_close(raw):
            name = raw[2:-1].strip().lower()
            if not rules.allows_bbcode(name):
                continue
            later_closes_by_name[name] = later_closes_by_name.get(name, 0) + 1
        elif _is_plain_bbcode_open(raw):
            name = _bbcode_tag_name(raw)
            if rules.allows_bbcode(name) and later_closes_by_name.get(name, 0) > 0:
                matched_starts.add(found.start())
                later_closes_by_name[name] -= 1

    return matched_starts


def _matched_optional_pair_angle_open_starts(
    spans: list[re.Match[str]], rules: TagRules
) -> set[int]:
    later_closes_by_name: dict[str, int] = {}
    matched_starts: set[int] = set()

    for found in reversed(spans):
        raw = found.group(0)
        if not raw.startswith("<"):
            continue

        kind, name = _classify_raw_tag(raw)
        if not _is_raw_tag_allowed(raw, kind, name, rules):
            continue
        canonical_name = _canonical_raw_tag_name(raw, name, rules)

        if kind == TAG_CLOSE and canonical_name is not None:
            later_closes_by_name[canonical_name] = (
                later_closes_by_name.get(canonical_name, 0) + 1
            )
        elif (
            kind == TAG_OPEN
            and name is not None
            and rules.is_angle_optional_pair(name)
            and later_closes_by_name.get(canonical_name or "", 0) > 0
        ):
            matched_starts.add(found.start())
            later_closes_by_name[canonical_name or ""] -= 1

    return matched_starts


def _is_unmatched_plain_bbcode_open(
    raw: str, start: int, matched_plain_bbcode_starts: set[int]
) -> bool:
    return _is_plain_bbcode_open(raw) and start not in matched_plain_bbcode_starts


def _is_plain_bbcode_open(raw: str) -> bool:
    return re.fullmatch(r"\[[A-Za-z][A-Za-z0-9:._-]*\]", raw) is not None


def _is_named_bbcode_close(raw: str) -> bool:
    return re.fullmatch(r"\[/[A-Za-z][A-Za-z0-9:._-]*\]", raw) is not None


def _protected_token_sort_key(placeholder: str) -> tuple[int, int]:
    parsed = parse_protected_token(placeholder)
    if parsed is None:
        return 10**9, 99
    index, kind = parsed
    kind_order = {TAG_OPEN: 0, PROTECTED_SINGLE: 1, TAG_CLOSE: 2}
    return index, kind_order[kind]


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
