from __future__ import annotations

from ._template_workflow import (
    fill_target_column_workbook,
    generate_tm_pairs,
    generate_workbook,
)

extract_tm_pairs = generate_tm_pairs
prepare_translation = generate_workbook
fill_translation = fill_target_column_workbook

__all__ = [
    "extract_tm_pairs",
    "fill_target_column_workbook",
    "fill_translation",
    "generate_tm_pairs",
    "generate_workbook",
    "prepare_translation",
]
