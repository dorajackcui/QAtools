from __future__ import annotations

import unittest

from tools.term_matching import (
    TermMappingEntry,
    build_matcher,
    find_row_terms,
    normalize_text,
    s_plural_token_variants,
    term_has_expected_target,
    text_contains_term,
)


def make_entry(source_term: str, target_term: str) -> TermMappingEntry:
    return TermMappingEntry(
        source_term=source_term,
        target_term=target_term,
        normalized_source=normalize_text(source_term, case_sensitive=False),
        normalized_target=normalize_text(target_term, case_sensitive=False),
    )


class TermMatchingTests(unittest.TestCase):
    def find_sources(
        self,
        source_text: str,
        rows: list[tuple[str, str]],
        *,
        match_mode: str = "hybrid-boundary",
    ) -> list[str]:
        entries = [make_entry(source, target) for source, target in rows]
        matches = find_row_terms(
            source_text,
            build_matcher(entries),
            case_sensitive=False,
            match_mode=match_mode,
        )
        return [entry.source_term for entry in matches]

    def test_longest_overlapping_term_wins(self) -> None:
        self.assertEqual(
            self.find_sources("Use API key to sign in.", [("API", "接口"), ("API key", "接口密钥")]),
            ["API key"],
        )

    def test_explicit_plural_term_wins_over_generated_variant(self) -> None:
        self.assertEqual(
            self.find_sources("Collect shards.", [("shards", "ssss"), ("shard", "éclat")]),
            ["shards"],
        )

    def test_hybrid_boundary_rejects_embedded_ascii_terms(self) -> None:
        self.assertEqual(self.find_sources("training material", [("rain", "雨")]), [])
        self.assertEqual(self.find_sources("account setup", [("ACC", "ACC")]), [])
        self.assertEqual(self.find_sources("ACC_001 pending", [("ACC", "ACC")]), [])

    def test_hybrid_boundary_allows_punctuation_boundaries(self) -> None:
        self.assertEqual(self.find_sources("Use API-key.", [("API", "接口")]), ["API"])

    def test_chinese_terms_match_and_prefer_longest_overlap(self) -> None:
        self.assertEqual(
            self.find_sources(
                "金庭后门停泊处",
                [
                    ("金庭", "Golden Court"),
                    ("金庭后门", "Golden Court rear gate"),
                    ("金庭后门停泊处", "Golden Court rear dock"),
                ],
            ),
            ["金庭后门停泊处"],
        )

    def test_substring_mode_keeps_compatibility_behavior(self) -> None:
        self.assertEqual(
            self.find_sources("training material", [("rain", "雨")], match_mode="substring"),
            ["rain"],
        )

    def test_text_contains_term_uses_the_same_boundary_rules(self) -> None:
        self.assertFalse(text_contains_term("account setup", "acc", match_mode="hybrid-boundary"))
        self.assertTrue(text_contains_term("api-key", "api", match_mode="hybrid-boundary"))
        self.assertTrue(text_contains_term("account setup", "acc", match_mode="substring"))
        self.assertTrue(
            text_contains_term(
                normalize_text("Intro\\n\\nVigo-09", case_sensitive=False),
                normalize_text("Vigo-09", case_sensitive=False),
                match_mode="hybrid-boundary",
            )
        )

    def test_plural_variants_cover_ascii_y_ies_but_not_non_ascii(self) -> None:
        self.assertEqual(s_plural_token_variants("应用程序"), ("应用程序",))
        self.assertIn("policies", s_plural_token_variants("policy"))
        self.assertIn("policy", s_plural_token_variants("policies"))

    def test_optional_target_plural_matching_is_explicit(self) -> None:
        entry = make_entry("源", "company policy")
        source_text = normalize_text("源", case_sensitive=False)
        plural_target = normalize_text("company policies", case_sensitive=False)

        self.assertTrue(
            term_has_expected_target(
                source_text,
                plural_target,
                entry,
                match_mode="hybrid-boundary",
                allow_target_plural_variants=True,
            )
        )
        self.assertFalse(
            term_has_expected_target(
                source_text,
                plural_target,
                entry,
                match_mode="hybrid-boundary",
                allow_target_plural_variants=False,
            )
        )


if __name__ == "__main__":
    unittest.main()
