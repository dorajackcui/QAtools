from __future__ import annotations

from .cli import (
    _main_extract,
    _main_fill,
    _main_legacy,
    _main_tm_extract,
    _parse_examples,
    _print_stats,
    _print_tm_stats,
    _print_top_level_help,
    main,
)
from .excel_io import *  # noqa: F403
from .interactive import (
    _interactive_extract,
    _interactive_fill,
    _interactive_tm_extract,
    _normalize_optional_column,
    _prompt_int,
    _prompt_text,
    _prompt_yes_no,
    _user_path,
    run_interactive,
)
from .template_engine import *  # noqa: F403
from .workflow import *  # noqa: F403

__all__ = [
    "main",
    "_parse_examples",
    "_main_tm_extract",
    "_main_extract",
    "_main_fill",
    "_main_legacy",
    "_print_stats",
    "_print_tm_stats",
    "_print_top_level_help",
    "run_interactive",
    "_interactive_tm_extract",
    "_interactive_extract",
    "_interactive_fill",
    "_prompt_text",
    "_prompt_int",
    "_prompt_yes_no",
    "_user_path",
    "_normalize_optional_column",
    "extract_tm_pairs",
    "fill_target_column_workbook",
    "fill_translation",
    "generate_tm_pairs",
    "generate_workbook",
    "prepare_translation",
    "_cell_value",
    "_default_extract_output_path",
    "_default_fill_output_path",
    "_default_legacy_output_path",
    "_default_tm_output_path",
    "_default_to_translate_output_path",
    "_default_work_dir",
    "_load_translated_units",
    "_load_unit_sheet",
    "_read_headers",
    "_read_source_rows",
    "_resolve_column",
    "_write_output_workbook",
    "_write_target_column_workbook",
    "_write_tm_workbook",
    "_write_to_translate_workbook",
    "NAMED_PLACEHOLDER_RE",
    "PLACEHOLDER_RE",
    "VAR_RE",
    "apply_target_template",
    "infer_target_template",
    "is_candidate_template",
    "is_non_translatable_segment",
    "parse_template",
]


if __name__ == "__main__":
    raise SystemExit(main())
