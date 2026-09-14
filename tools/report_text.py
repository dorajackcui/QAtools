"""Bound report descriptions without shortening source/target snapshots."""

from collections.abc import Sequence


VARIANT_TEXT_LIMIT = 120
VARIANT_COUNT_LIMIT = 10
CHECK_DESCRIPTION_LIMIT = 2000
EXCEL_CELL_TEXT_LIMIT = 32767
TRUNCATION_NOTICE = "…（已截断）"


def truncate_report_text(text: str, max_chars: int) -> str:
    if max_chars < len(TRUNCATION_NOTICE):
        raise ValueError("描述长度必须足够容纳截断提示。")
    if len(text) <= max_chars:
        return text
    return text[:max_chars - len(TRUNCATION_NOTICE)] + TRUNCATION_NOTICE


def format_consistency_variants(
    variants: Sequence[str],
    label: str,
    *,
    max_variant_chars: int = VARIANT_TEXT_LIMIT,
    max_variants: int = VARIANT_COUNT_LIMIT,
) -> str:
    """Display a bounded preview while retaining the full variant count."""
    if max_variants < 1:
        raise ValueError("展示的版本数必须至少为 1。")
    previews = [
        f"{index}: {truncate_report_text(text, max_variant_chars)}"
        for index, text in enumerate(variants[:max_variants], 1)
    ]
    if len(variants) > max_variants:
        previews.append(f"另 {len(variants) - max_variants} 种未展示。")
    return f"{len(variants)} 种{label}：" + "；".join(previews)
