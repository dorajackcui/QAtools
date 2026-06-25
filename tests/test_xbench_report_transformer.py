from __future__ import annotations

import unittest

from tools.xbench_report_transformer.transform_xbench_report import (
    format_issue_text,
    parse_metadata,
    parse_qa_title,
)


class MetadataParsingTests(unittest.TestCase):
    def test_parse_two_metadata_lines_as_key_and_file_name(self) -> None:
        metadata = parse_metadata("LDLG_Text_ZH_q203101_Line_3\n磐城【配音】.xlsx")
        self.assertEqual(metadata.key, "LDLG_Text_ZH_q203101_Line_3")
        self.assertEqual(metadata.file_name, "磐城【配音】.xlsx")

    def test_parse_single_file_name_line_without_key(self) -> None:
        metadata = parse_metadata("磐城【配音】.xlsx")
        self.assertEqual(metadata.key, "")
        self.assertEqual(metadata.file_name, "磐城【配音】.xlsx")

    def test_parse_single_non_file_line_as_key(self) -> None:
        metadata = parse_metadata("LDLG_Text_ZH_q203101_Line_3")
        self.assertEqual(metadata.key, "LDLG_Text_ZH_q203101_Line_3")
        self.assertEqual(metadata.file_name, "")

    def test_parse_empty_metadata(self) -> None:
        metadata = parse_metadata(None)
        self.assertEqual(metadata.key, "")
        self.assertEqual(metadata.file_name, "")

    def test_parse_metadata_ignores_extra_lines_for_now(self) -> None:
        metadata = parse_metadata("Key_1\nfile.xlsx\nextra")
        self.assertEqual(metadata.key, "Key_1")
        self.assertEqual(metadata.file_name, "file.xlsx")


class QaTitleParsingTests(unittest.TestCase):
    def test_format_key_term_mismatch_title_as_term_pair_issue(self) -> None:
        issue = parse_qa_title("Key Term Mismatch (提示 / Avis)")
        self.assertEqual(issue.issue_type, "Key Term Mismatch")
        self.assertEqual(issue.source_term, "提示")
        self.assertEqual(issue.target_term, "Avis")
        self.assertEqual(format_issue_text(issue), "提示 -> Avis：Key Term Mismatch")

    def test_format_issue_preserves_quoted_terms(self) -> None:
        issue = parse_qa_title('Key Term Mismatch (“斑鸠” / "Colombe")')
        self.assertEqual(format_issue_text(issue), '“斑鸠” -> "Colombe"：Key Term Mismatch')

    def test_unparseable_title_is_used_as_issue_type(self) -> None:
        issue = parse_qa_title("Target same as Source")
        self.assertEqual(issue.issue_type, "Target same as Source")
        self.assertEqual(issue.source_term, "")
        self.assertEqual(issue.target_term, "")
        self.assertEqual(format_issue_text(issue), "Target same as Source")


if __name__ == "__main__":
    unittest.main()
