"""Shared comparison keys for the three consistency checks.

Only trim surrounding whitespace and enclosing wrapper pairs. The normalized
text stays a contiguous slice of the original, so substring offsets are exact.
"""

import unicodedata

WRAPPER_PAIRS = {
    '"': '"', "'": "'", "“": "”", "‘": "’", "「": "」", "『": "』",
    "«": "»", "‹": "›", "＂": "＂", "＇": "＇",
    "【": "】", "[": "]", "(": ")", "（": "）", "《": "》", "〈": "〉",
    "〔": "〕", "〖": "〗", "［": "］",
}


def _encloses_text(text: str, start: int, end: int) -> bool:
    opening = text[start]
    closing = WRAPPER_PAIRS.get(opening)
    if closing is None or text[end - 1] != closing:
        return False
    depth = 1
    escaped = False
    for index in range(start + 1, end):
        char = text[index]
        if escaped:
            escaped = False
            continue
        if char == "\\":
            escaped = True
            continue
        if opening != closing and char == opening:
            depth += 1
        elif char == closing:
            # Preserve Latin contractions, but do not read adjacent Chinese
            # quotations (e.g. '甲'和'乙') as apostrophes inside one word.
            if (
                closing in {"'", "’", "＇"}
                and index + 1 < end
                and unicodedata.name(text[index - 1], "").startswith("LATIN ")
                and unicodedata.name(text[index + 1], "").startswith("LATIN ")
            ):
                continue
            depth -= 1
            if depth == 0:
                return index == end - 1
    return False


def normalized_text_with_offset(value: object) -> tuple[str, int]:
    """Return the comparison text and its zero-based start in the original."""
    text = "" if value is None else str(value)
    start, end = 0, len(text)
    while True:
        while start < end and text[start].isspace():
            start += 1
        while end > start and text[end - 1].isspace():
            end -= 1
        if end - start < 2 or not _encloses_text(text, start, end):
            return text[start:end], start
        start += 1
        end -= 1


def normalize_consistency_text(value: object) -> str:
    return normalized_text_with_offset(value)[0]
