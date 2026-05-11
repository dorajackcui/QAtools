from __future__ import annotations

import re
from collections import defaultdict

from .models import TemplateMatch
from .tag_engine import PROTECTED_TOKEN_RE

VAR_RE = re.compile(r"#[0-9A-Fa-f]{6}|\d+(?:[./:-]\d+)+|\d+(?:\.\d+)?")
PLACEHOLDER_RE = re.compile(r"\{[A-Za-z_][A-Za-z0-9_]*\}")


def _iter_protected_aware_spans(source: str):
    pos = 0
    for found in PROTECTED_TOKEN_RE.finditer(source):
        if found.start() > pos:
            yield source[pos : found.start()], False
        yield found.group(0), True
        pos = found.end()
    if pos < len(source):
        yield source[pos:], False


def _replace_outside_protected_tokens(source: str, old: str, new: str) -> tuple[str, bool]:
    changed = False
    chunks: list[str] = []
    for chunk, protected in _iter_protected_aware_spans(source):
        if protected:
            chunks.append(chunk)
            continue
        replaced = chunk.replace(old, new)
        if replaced != chunk:
            changed = True
        chunks.append(replaced)
    return "".join(chunks), changed


def parse_template(text: object) -> TemplateMatch:
    source = "" if text is None else str(text)
    chunks: list[str] = []
    values: dict[str, str] = {}
    counters: dict[str, int] = defaultdict(int)

    for span, protected in _iter_protected_aware_spans(source):
        if protected:
            chunks.append(span)
            continue
        pos = 0
        for found in VAR_RE.finditer(span):
            chunks.append(span[pos : found.start()])
            value = found.group(0)
            key = _variable_key(value, counters)
            chunks.append("{" + key + "}")
            values[key] = value
            pos = found.end()
        chunks.append(span[pos:])

    return TemplateMatch(source, "".join(chunks), values)


def _variable_key(value: str, counters: dict[str, int]) -> str:
    if value.startswith("#"):
        prefix = "color"
    elif re.fullmatch(r"\d+-\d+(?:-\d+)*", value):
        prefix = "stage"
    elif re.fullmatch(r"\d+(?:[./:]\d+)+", value):
        prefix = "seq"
    else:
        prefix = "num"
    counters[prefix] += 1
    return f"{prefix}{counters[prefix]}"


def infer_target_template(values: dict[str, str], target_text: object) -> str | None:
    target_template = "" if target_text is None else str(target_text)
    matched = False
    tokens: dict[str, str] = {}

    for index, (key, value) in enumerate(
        sorted(values.items(), key=lambda item: len(item[1]), reverse=True)
    ):
        if not value:
            continue
        token = "\x00" + _letters_token(index) + "\x00"
        target_template, changed = _replace_outside_protected_tokens(
            target_template, value, token
        )
        if changed:
            tokens[token] = "{" + key + "}"
            matched = True

    for token, placeholder in tokens.items():
        target_template = target_template.replace(token, placeholder)

    return target_template if matched else None


def _letters_token(index: int) -> str:
    letters = []
    value = index
    while True:
        letters.append(chr(ord("A") + (value % 26)))
        value = value // 26 - 1
        if value < 0:
            return "".join(reversed(letters))


def apply_target_template(target_template: str, values: dict[str, str]) -> str:
    result = target_template
    for key, value in values.items():
        result = result.replace("{" + key + "}", value)
    return result


def is_non_translatable_segment(source: str) -> bool:
    text = source.strip()
    if not text:
        return True
    return not re.search(
        r"[A-Za-z\u3400-\u9fff\u3040-\u30ff\uac00-\ud7af]", text
    )


def is_candidate_template(match: TemplateMatch) -> bool:
    if not match.values:
        return False
    literal = PLACEHOLDER_RE.sub("", match.template)
    literal = re.sub(r"\s+", "", literal)
    return len(literal) >= 2


__all__ = [
    "PLACEHOLDER_RE",
    "VAR_RE",
    "apply_target_template",
    "infer_target_template",
    "is_candidate_template",
    "is_non_translatable_segment",
    "parse_template",
]
