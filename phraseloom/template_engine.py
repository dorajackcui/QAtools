from __future__ import annotations

import re
from collections import defaultdict

from .models import TemplateMatch
from .tag_engine import is_tag_placeholder

VAR_RE = re.compile(
    r"\{[^{}]+\}|#[0-9A-Fa-f]{6}|\d+(?:[./:-]\d+)+|\d+(?:\.\d+)?"
)
PLACEHOLDER_RE = re.compile(r"\{[A-Za-z_][A-Za-z0-9_]*\}")
NAMED_PLACEHOLDER_RE = re.compile(r"^\{([A-Za-z_][A-Za-z0-9_]*)\}$")


def parse_template(text: object) -> TemplateMatch:
    source = "" if text is None else str(text)
    chunks: list[str] = []
    values: dict[str, str] = {}
    pos = 0
    counters: dict[str, int] = defaultdict(int)

    for found in VAR_RE.finditer(source):
        chunks.append(source[pos : found.start()])
        value = found.group(0)
        if is_tag_placeholder(value):
            chunks.append(value)
            pos = found.end()
            continue
        key = _variable_key(value, counters)
        chunks.append("{" + key + "}")
        values[key] = value
        pos = found.end()

    chunks.append(source[pos:])
    return TemplateMatch(source, "".join(chunks), values)


def _variable_key(value: str, counters: dict[str, int]) -> str:
    named = NAMED_PLACEHOLDER_RE.match(value)
    if named:
        return named.group(1)
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
        if value in target_template:
            target_template = target_template.replace(value, token)
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
    "NAMED_PLACEHOLDER_RE",
    "PLACEHOLDER_RE",
    "VAR_RE",
    "apply_target_template",
    "infer_target_template",
    "is_candidate_template",
    "is_non_translatable_segment",
    "parse_template",
]
