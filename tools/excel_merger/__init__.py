"""Merge the active worksheets from a folder of Excel workbooks."""

from .merge_active_sheets import (
    MergeSummary,
    build_default_output_path,
    merge_active_sheets,
)

__all__ = [
    "MergeSummary",
    "build_default_output_path",
    "merge_active_sheets",
]
