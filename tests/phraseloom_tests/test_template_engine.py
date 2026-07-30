import unittest

from phraseloom.tag_engine import extract_tags
from phraseloom.template_engine import (
    apply_target_template,
    is_candidate_template,
    is_non_translatable_segment,
    parse_template,
)


class TemplateEngineTests(unittest.TestCase):
    def test_parses_supported_variable_types(self) -> None:
        match = parse_template(
            "LV10 / Story 10-20 / Version 1.2.3 / Color #FF0000"
        )

        self.assertEqual(
            match.template,
            "LV{num1} / Story {stage1} / Version {seq1} / Color {color1}",
        )
        self.assertEqual(
            match.values,
            {
                "num1": "10",
                "stage1": "10-20",
                "seq1": "1.2.3",
                "color1": "#FF0000",
            },
        )

    def test_ignores_digits_inside_protected_tokens(self) -> None:
        match = parse_template("{1>Level 10<2} {3}")

        self.assertEqual(match.template, "{1>Level {num1}<2} {3}")
        self.assertEqual(match.values, {"num1": "10"})

    def test_raw_brace_placeholder_is_protected_before_parsing(self) -> None:
        protected = extract_tags("Damage {0}: 10").text
        match = parse_template(protected)

        self.assertEqual(protected, "Damage {1}: 10")
        self.assertEqual(match.template, "Damage {1}: {num1}")
        self.assertEqual(match.values, {"num1": "10"})

    def test_applies_target_template(self) -> None:
        self.assertEqual(
            apply_target_template(
                "Niveau {num1} / histoire {stage1}",
                {"num1": "11", "stage1": "21-30"},
            ),
            "Niveau 11 / histoire 21-30",
        )

    def test_classifies_template_and_non_translatable_segments(self) -> None:
        self.assertTrue(is_candidate_template(parse_template("Level 10")))
        self.assertFalse(is_candidate_template(parse_template("10")))
        self.assertTrue(is_non_translatable_segment("--- 123"))
        self.assertFalse(is_non_translatable_segment("Level 123"))


if __name__ == "__main__":
    unittest.main()
