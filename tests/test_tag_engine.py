import unittest
from dataclasses import astuple


class TagEngineTests(unittest.TestCase):
    def test_placeholder_helpers_accept_only_tag_namespace(self):
        from phraseloom.tag_engine import (
            TAG_CLOSE,
            TAG_OPEN,
            TAG_SELF,
            is_tag_placeholder,
            make_tag_placeholder,
            parse_tag_placeholder,
        )

        self.assertEqual(TAG_OPEN, "op")
        self.assertEqual(TAG_CLOSE, "cl")
        self.assertEqual(TAG_SELF, "sf")
        self.assertEqual(make_tag_placeholder(1, TAG_OPEN), "{t1_op}")
        self.assertEqual(make_tag_placeholder(1, TAG_CLOSE), "{t1_cl}")
        self.assertEqual(make_tag_placeholder(2, TAG_SELF), "{t2_sf}")
        self.assertEqual(parse_tag_placeholder("{t1_op}"), (1, TAG_OPEN))
        self.assertEqual(parse_tag_placeholder("{t1_cl}"), (1, TAG_CLOSE))
        self.assertEqual(parse_tag_placeholder("{t2_sf}"), (2, TAG_SELF))
        self.assertTrue(is_tag_placeholder("{t1_op}"))
        self.assertTrue(is_tag_placeholder("{t1_cl}"))
        self.assertTrue(is_tag_placeholder("{t2_sf}"))
        self.assertFalse(is_tag_placeholder("{num1}"))
        self.assertFalse(is_tag_placeholder("{tag1_op}"))

    def test_extracts_angle_tags_and_preserves_token_fields(self):
        from phraseloom.tag_engine import TAG_CLOSE, TAG_OPEN, TAG_SELF, extract_tags

        result = extract_tags('<a href="shop">VIP10</a> <img src="coin.png"/>')

        self.assertEqual(result.text, "{t1_op}VIP10{t1_cl} {t2_sf}")
        self.assertEqual(
            tuple(astuple(tag) for tag in result.tags),
            (
                (1, TAG_OPEN, "{t1_op}", '<a href="shop">'),
                (1, TAG_CLOSE, "{t1_cl}", "</a>"),
                (2, TAG_SELF, "{t2_sf}", '<img src="coin.png"/>'),
            ),
        )
        self.assertEqual(result.warnings, ())

    def test_extracts_bbcode_tags(self):
        from phraseloom.tag_engine import extract_tags

        result = extract_tags("[color=#ff0]Bonus[/]")

        self.assertEqual(result.text, "{t1_op}Bonus{t1_cl}")
        self.assertEqual(result.tags[0].raw, "[color=#ff0]")
        self.assertEqual(result.tags[1].raw, "[/]")
        self.assertEqual(result.warnings, ())

    def test_extracts_named_bbcode_close_tags(self):
        from phraseloom.tag_engine import TAG_CLOSE, TAG_OPEN, extract_tags

        result = extract_tags("[b]Bold[/b]")

        self.assertEqual(result.text, "{t1_op}Bold{t1_cl}")
        self.assertEqual(
            tuple(astuple(tag) for tag in result.tags),
            (
                (1, TAG_OPEN, "{t1_op}", "[b]"),
                (1, TAG_CLOSE, "{t1_cl}", "[/b]"),
            ),
        )
        self.assertEqual(result.warnings, ())

    def test_unpaired_close_stays_raw_and_warns(self):
        from phraseloom.tag_engine import extract_tags

        result = extract_tags("</a>Text")

        self.assertEqual(result.text, "</a>Text")
        self.assertEqual(result.tags, ())
        self.assertTrue(any("unpaired close" in warning for warning in result.warnings))

    def test_unclosed_open_serializes_and_warns(self):
        from phraseloom.tag_engine import TAG_OPEN, extract_tags

        result = extract_tags('<a href="x">Text')

        self.assertEqual(result.text, "{t1_op}Text")
        self.assertEqual(len(result.tags), 1)
        self.assertEqual(result.tags[0].kind, TAG_OPEN)
        self.assertEqual(result.tags[0].raw, '<a href="x">')
        self.assertTrue(
            any("open tag has no close partner" in warning for warning in result.warnings)
        )

    def test_shorthand_angle_close_pairs_with_nearest_open(self):
        from phraseloom.tag_engine import extract_tags

        result = extract_tags("<a>Text</>")

        self.assertEqual(result.text, "{t1_op}Text{t1_cl}")
        self.assertEqual(result.tags[1].raw, "</>")
        self.assertEqual(result.warnings, ())

    def test_generated_placeholders_skip_reserved_raw_placeholders(self):
        from phraseloom.tag_engine import TAG_CLOSE, TAG_OPEN, extract_tags, restore_tags

        result = extract_tags("{t1_op}<a>x</a>")

        self.assertEqual(result.text, "{t1_op}{t2_op}x{t2_cl}")
        self.assertEqual(
            tuple(astuple(tag) for tag in result.tags),
            (
                (2, TAG_OPEN, "{t2_op}", "<a>"),
                (2, TAG_CLOSE, "{t2_cl}", "</a>"),
            ),
        )
        self.assertIn("reserved_tag_placeholder: {t1_op}", result.warnings)
        self.assertEqual(restore_tags(result.text, result.tags), "{t1_op}<a>x</a>")

    def test_misnested_named_tags_do_not_cross_pair(self):
        from phraseloom.tag_engine import extract_tags

        result = extract_tags("<a><b>x</a>y</b>")

        self.assertEqual(result.text, "{t1_op}{t2_op}x</a>y{t2_cl}")
        self.assertEqual(
            [tag.placeholder for tag in result.tags],
            ["{t1_op}", "{t2_op}", "{t2_cl}"],
        )
        self.assertTrue(
            any(
                "mismatch" in warning or "unpaired close" in warning
                for warning in result.warnings
            )
        )

    def test_unmatched_plain_bbcode_text_stays_raw(self):
        from phraseloom.tag_engine import extract_tags

        result = extract_tags("Press [OK] to continue")

        self.assertEqual(result.text, "Press [OK] to continue")
        self.assertEqual(result.tags, ())

    def test_suspicious_raw_text_stays_unchanged_and_warns_reserved_namespace(self):
        from phraseloom.tag_engine import extract_tags

        result = extract_tags("Use 3 < 5 and keep {t1_op}")

        self.assertEqual(result.text, "Use 3 < 5 and keep {t1_op}")
        self.assertEqual(result.tags, ())
        self.assertIn("reserved_tag_placeholder: {t1_op}", result.warnings)

    def test_identifies_tag_only_segments(self):
        from phraseloom.tag_engine import is_tag_only_segment

        self.assertTrue(is_tag_only_segment("{t1_sf}"))
        self.assertTrue(is_tag_only_segment("{t1_op}{t1_cl}"))
        self.assertTrue(is_tag_only_segment("{t1_op} {t1_cl}"))
        self.assertFalse(is_tag_only_segment("{t1_op}Click{t1_cl}"))
        self.assertFalse(is_tag_only_segment("{t1_sf} 100 coins"))

    def test_restore_tags_replaces_known_placeholders_only(self):
        from phraseloom.tag_engine import extract_tags, restore_tags

        extraction = extract_tags('<a href="shop">VIP10</a>')

        restored = restore_tags("{t1_op}VIP{t1_cl} {t9_sf}", extraction.tags)

        self.assertEqual(restored, '<a href="shop">VIP</a> {t9_sf}')

    def test_validate_tag_placeholders_reports_extra_counts(self):
        from phraseloom.tag_engine import extract_tags, validate_tag_placeholders

        extraction = extract_tags('<a href="shop">VIP10</a>')

        validation = validate_tag_placeholders(
            "{t1_op}VIP10{t1_cl} {t2_sf}", extraction.tags
        )

        self.assertEqual(validation.warnings, ("tag_mismatch: extra {t2_sf}",))

    def test_validate_tag_placeholders_reports_missing_counts(self):
        from phraseloom.tag_engine import extract_tags, validate_tag_placeholders

        extraction = extract_tags("<a>x</a>")

        validation = validate_tag_placeholders("{t1_op}x", extraction.tags)

        self.assertEqual(validation.warnings, ("tag_mismatch: missing {t1_cl}",))

    def test_serialize_known_tags_preserves_repeated_raw_tags_in_order(self):
        from phraseloom.tag_engine import TAG_SELF, TagToken, serialize_known_tags

        tags = (
            TagToken(1, TAG_SELF, "{t1_sf}", "<br/>"),
            TagToken(2, TAG_SELF, "{t2_sf}", "<br/>"),
        )

        result = serialize_known_tags("<br/> A <br/>", tags)

        self.assertEqual(result.text, "{t1_sf} A {t2_sf}")
        self.assertEqual(result.tags, tags)
        self.assertEqual(result.warnings, ())

    def test_serialize_known_tags_returns_empty_extraction_for_empty_target(self):
        from phraseloom.tag_engine import TagExtraction, extract_tags, serialize_known_tags

        extraction = extract_tags("<a>x</a>")

        result = serialize_known_tags("", extraction.tags)

        self.assertEqual(result, TagExtraction("", (), ()))


if __name__ == "__main__":
    unittest.main()
