from __future__ import annotations

from phraseloom.cli import main
from phraseloom.excel_io import (
    _default_extract_output_path,
    _default_fill_output_path,
    _default_tm_output_path,
    _default_to_translate_output_path,
    _default_work_dir,
)
from phraseloom.interactive import _user_path
from phraseloom.template_engine import (
    apply_target_template,
    infer_target_template,
    parse_template,
)
from phraseloom.workflow import (
    fill_target_column_workbook,
    generate_tm_pairs,
    generate_workbook,
)

__all__ = [
    "apply_target_template",
    "fill_target_column_workbook",
    "generate_tm_pairs",
    "generate_workbook",
    "infer_target_template",
    "main",
    "parse_template",
    "_default_extract_output_path",
    "_default_fill_output_path",
    "_default_tm_output_path",
    "_default_to_translate_output_path",
    "_default_work_dir",
    "_user_path",
]


if __name__ == "__main__":
    raise SystemExit(main())
