from __future__ import annotations

import unittest
from pathlib import Path
from unittest.mock import patch

from tools.excel_output import load_workbook_for_editing

from tools.chinese_target_checker.check_chinese_target import (
    build_default_output_path as build_chinese_target_output_path,
)
from tools.excel_line_splitter.split_excel_lines import (
    build_default_output_path as build_split_lines_output_path,
)
from tools.french_nbsp_restorer.restore_french_nbsp import (
    build_default_output_path as build_french_nbsp_output_path,
)
from tools.tag_placeholder_checker.check_tags_and_placeholders import (
    build_default_output_path as build_tag_check_output_path,
)
from tools.term_pair_checker.extract_terms_from_excel import (
    build_default_output_path as build_term_pair_output_path,
)
from tools.workflow.workflow_runner import (
    build_default_output_path as build_workflow_output_path,
)


class ExcelOutputPathTests(unittest.TestCase):
    def test_default_output_paths_put_tool_prefix_before_original_name(self) -> None:
        input_path = Path("/tmp/project/source.xlsx")

        cases = (
            (build_workflow_output_path, "workflow_check_source.xlsx"),
            (build_term_pair_output_path, "term_pair_check_source.xlsx"),
            (build_tag_check_output_path, "tag_check_source.xlsx"),
            (build_chinese_target_output_path, "target_chinese_check_source.xlsx"),
            (build_split_lines_output_path, "split_lines_source.xlsx"),
            (build_french_nbsp_output_path, "french_nbsp_restore_source.xlsx"),
        )

        for build_output_path, expected_name in cases:
            with self.subTest(expected_name=expected_name):
                self.assertEqual(build_output_path(input_path), input_path.with_name(expected_name))

    def test_default_output_paths_keep_excel_extension(self) -> None:
        input_path = Path("/tmp/project/source.xlsm")

        self.assertEqual(
            build_tag_check_output_path(input_path),
            input_path.with_name("tag_check_source.xlsm"),
        )

    def test_editable_workbook_loader_preserves_vba_for_xlsm_only(self) -> None:
        with patch("tools.excel_output.load_workbook") as load_workbook_mock:
            load_workbook_for_editing(Path("macro.xlsm"))
            load_workbook_mock.assert_called_once_with(
                Path("macro.xlsm"),
                keep_vba=True,
            )

        with patch("tools.excel_output.load_workbook") as load_workbook_mock:
            load_workbook_for_editing(Path("plain.xlsx"))
            load_workbook_mock.assert_called_once_with(
                Path("plain.xlsx"),
                keep_vba=False,
            )


if __name__ == "__main__":
    unittest.main()
