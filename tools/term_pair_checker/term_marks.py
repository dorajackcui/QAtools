"""Term mark extraction and cleanup helpers for term pair checking."""

from __future__ import annotations

import json
import re
from collections.abc import Iterable
from dataclasses import dataclass
from pathlib import Path


SUPPORTED_MARKS = ("【】", "[]")
DEFAULT_MARK_STYLES = ("【】", "[]")
SINGLE_ASCII_LETTER_PATTERN = re.compile(r"^[A-Za-z]$")
NUMERIC_TAG_BOUNDARY_SPAN_PATTERN = re.compile(r"^\d+\}.*\{\d+$", re.DOTALL)
SQUARE_COLOR_TAG_PATTERN = re.compile(r"^(?:color\s*=.+|/color)$", re.IGNORECASE)
MARK_PATTERNS: dict[str, tuple[re.Pattern[str], ...]] = {
    "【】": (re.compile(r"【([^【】]+)】"),),
    "[]": (
        re.compile(r"\[([^\[\]]+)\]"),
        re.compile(r"［([^［］]+)］"),
    ),
}


@dataclass(frozen=True)
class ExtractedTerm:
    display_text: str
    plain_text: str
    start: int
    end: int


def normalize_mark_styles(
    mark_styles: Iterable[str] | None = None,
    mark_style: str | None = None,
) -> tuple[str, ...]:
    if mark_styles is None:
        raw_mark_styles: list[str] = [mark_style] if mark_style else list(DEFAULT_MARK_STYLES)
    elif isinstance(mark_styles, str):
        raw_mark_styles = [mark_styles]
    else:
        raw_mark_styles = [style for style in mark_styles if style]

    invalid_mark_styles = [style for style in raw_mark_styles if style not in SUPPORTED_MARKS]
    if invalid_mark_styles:
        raise ValueError(f"不支持的 mark 类型: {'、'.join(invalid_mark_styles)}")

    normalized_mark_styles = tuple(style for style in SUPPORTED_MARKS if style in raw_mark_styles)
    if not normalized_mark_styles:
        raise ValueError("请至少选择一种 mark 类型。")
    return normalized_mark_styles


def extract_terms(
    text: object,
    mark_styles: Iterable[str] | None = None,
    mark_style: str | None = None,
    exclusion_patterns: Iterable[str] | None = None,
    exclusion_config_file: str | Path | None = None,
) -> list[str]:
    return [
        extracted_term.display_text
        for extracted_term in extract_term_details(
            text,
            mark_styles=mark_styles,
            mark_style=mark_style,
            exclusion_patterns=exclusion_patterns,
            exclusion_config_file=exclusion_config_file,
        )
    ]


def normalize_exclusion_patterns(exclusion_patterns: Iterable[str] | None) -> tuple[str, ...]:
    if exclusion_patterns is None:
        raw_patterns: list[str] = []
    elif isinstance(exclusion_patterns, str):
        raw_patterns = [exclusion_patterns]
    else:
        raw_patterns = [pattern.strip() for pattern in exclusion_patterns if pattern and pattern.strip()]
    return tuple(raw_patterns)


def load_exclusion_patterns_from_file(config_file: str | Path | None = None) -> tuple[str, ...]:
    if config_file is None:
        raise ValueError("请提供术语候选排除配置文件路径。")
    config_path = Path(config_file).expanduser().resolve()
    if not config_path.exists():
        raise FileNotFoundError(f"误判排除配置文件不存在: {config_path}")

    try:
        config_data = json.loads(config_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"误判排除配置文件不是有效 JSON: {config_path} ({exc})") from exc

    patterns = config_data.get("patterns")
    if not isinstance(patterns, list) or any(not isinstance(pattern, str) for pattern in patterns):
        raise ValueError(f"误判排除配置格式错误: {config_path}，需要 JSON 对象中的 patterns 字符串数组。")
    return normalize_exclusion_patterns(patterns)


def resolve_exclusion_patterns(
    exclusion_patterns: Iterable[str] | None,
    exclusion_config_file: str | Path | None = None,
) -> tuple[str, ...]:
    if exclusion_patterns is not None:
        return normalize_exclusion_patterns(exclusion_patterns)
    if exclusion_config_file is not None:
        return load_exclusion_patterns_from_file(exclusion_config_file)
    return ()


def compile_exclusion_patterns(
    exclusion_patterns: Iterable[str] | None,
    exclusion_config_file: str | Path | None = None,
) -> tuple[re.Pattern[str], ...]:
    normalized_patterns = resolve_exclusion_patterns(exclusion_patterns, exclusion_config_file)
    compiled_patterns: list[re.Pattern[str]] = []
    for pattern in normalized_patterns:
        try:
            compiled_patterns.append(re.compile(pattern, re.IGNORECASE))
        except re.error as exc:
            raise ValueError(f"误判排除正则无效: {pattern} ({exc})") from exc
    return tuple(compiled_patterns)


def should_exclude_term(
    display_text: str,
    plain_text: str,
    exclusion_regexes: Iterable[re.Pattern[str]],
) -> bool:
    if display_text.startswith(("[", "［")) and SQUARE_COLOR_TAG_PATTERN.fullmatch(plain_text):
        return True
    if SINGLE_ASCII_LETTER_PATTERN.fullmatch(plain_text):
        return True
    if not any(character.isalpha() for character in plain_text):
        return True
    if NUMERIC_TAG_BOUNDARY_SPAN_PATTERN.fullmatch(plain_text):
        return True
    return any(
        regex.search(plain_text) or regex.search(display_text)
        for regex in exclusion_regexes
    )


def extract_term_details(
    text: object,
    mark_styles: Iterable[str] | None = None,
    mark_style: str | None = None,
    exclusion_patterns: Iterable[str] | None = None,
    exclusion_config_file: str | Path | None = None,
) -> list[ExtractedTerm]:
    if text is None:
        return []

    normalized_mark_styles = normalize_mark_styles(mark_styles=mark_styles, mark_style=mark_style)
    exclusion_regexes = compile_exclusion_patterns(exclusion_patterns, exclusion_config_file)
    text_value = str(text)
    matches: list[ExtractedTerm] = []

    for current_mark_style in normalized_mark_styles:
        for pattern in MARK_PATTERNS[current_mark_style]:
            for match in pattern.finditer(text_value):
                display_text = match.group(0)
                plain_text = match.group(1).strip()
                if should_exclude_term(display_text, plain_text, exclusion_regexes):
                    continue
                matches.append(
                    ExtractedTerm(
                        display_text=display_text,
                        plain_text=plain_text,
                        start=match.start(),
                        end=match.end(),
                    )
                )

    matches.sort(key=lambda item: (item.start, item.end))
    return matches


def strip_supported_marks(
    text: object,
    mark_styles: Iterable[str] | None = None,
    exclusion_patterns: Iterable[str] | None = None,
    exclusion_config_file: str | Path | None = None,
) -> str:
    text_value = "" if text is None else str(text)
    if not text_value:
        return ""

    normalized_mark_styles = normalize_mark_styles(
        mark_styles=SUPPORTED_MARKS if mark_styles is None else mark_styles
    )
    extracted_terms = extract_term_details(
        text_value,
        mark_styles=normalized_mark_styles,
        exclusion_patterns=exclusion_patterns,
        exclusion_config_file=exclusion_config_file,
    )
    if not extracted_terms:
        return text_value

    parts: list[str] = []
    last_index = 0
    for term_index, extracted_term in enumerate(extracted_terms):
        parts.append(text_value[last_index : extracted_term.start])
        if (
            parts
            and parts[-1]
            and extracted_term.plain_text
            and extracted_term.plain_text[0].isascii()
            and extracted_term.plain_text[0].isalnum()
            and _is_ascii_word_char(parts[-1][-1])
        ):
            parts.append(" ")
        parts.append(extracted_term.plain_text)

        next_term = extracted_terms[term_index + 1] if term_index + 1 < len(extracted_terms) else None
        text_until_next_term = text_value[
            extracted_term.end : next_term.start if next_term is not None else len(text_value)
        ]
        next_output_character = (
            text_until_next_term[0]
            if text_until_next_term
            else next_term.plain_text[0]
            if next_term is not None and next_term.plain_text
            else ""
        )
        if (
            extracted_term.plain_text
            and extracted_term.plain_text[-1].isascii()
            and extracted_term.plain_text[-1].isalnum()
            and _is_ascii_word_char(next_output_character)
        ):
            parts.append(" ")
        last_index = extracted_term.end
    parts.append(text_value[last_index:])
    return "".join(parts)


def _is_ascii_word_char(character: str) -> bool:
    return bool(character) and character.isascii() and (character.isalnum() or character == "_")
