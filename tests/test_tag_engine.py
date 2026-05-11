import unittest


class TagEngineTests(unittest.TestCase):
    def test_protected_token_helpers_accept_new_contract(self):
        from phraseloom.tag_engine import (
            RAW_PLACEHOLDER,
            TAG_CLOSE,
            TAG_OPEN,
            TAG_SELF,
            is_protected_token,
            is_tag_placeholder,
            make_protected_token,
            make_tag_placeholder,
            parse_protected_token,
        )

        self.assertEqual(TAG_OPEN, "op")
        self.assertEqual(TAG_CLOSE, "cl")
        self.assertEqual(TAG_SELF, "sf")
        self.assertEqual(RAW_PLACEHOLDER, "ph")
        self.assertEqual(make_protected_token(1, TAG_OPEN), "{1>")
        self.assertEqual(make_protected_token(2, TAG_CLOSE), "<2}")
        self.assertEqual(make_protected_token(3, TAG_SELF), "{3}")
        self.assertEqual(make_protected_token(4, RAW_PLACEHOLDER), "{4}")
        self.assertEqual(make_tag_placeholder(1, TAG_OPEN), "{1>")
        self.assertTrue(is_protected_token("{1>"))
        self.assertTrue(is_protected_token("<2}"))
        self.assertTrue(is_protected_token("{3}"))
        self.assertFalse(is_protected_token("{num1}"))
        self.assertFalse(is_protected_token("{t1_op}"))
        self.assertTrue(is_tag_placeholder("{3}"))
        self.assertEqual(parse_protected_token("{1>"), (1, TAG_OPEN))
        self.assertEqual(parse_protected_token("<2}"), (2, TAG_CLOSE))
        self.assertEqual(parse_protected_token("{3}"), (3, "single"))
        self.assertIsNone(parse_protected_token("{num1}"))
        self.assertIsNone(parse_protected_token("{t1_op}"))

    def test_protected_token_scans_accept_all_token_shapes(self):
        from phraseloom.tag_engine import extract_tags, validate_tag_placeholders

        extraction = extract_tags("{1><2}{3}")
        validation = validate_tag_placeholders("{1>", ())

        self.assertEqual(extraction.text, "{1><2}{1}")
        self.assertEqual(validation.warnings, ("tag_mismatch: extra {1>",))

    def test_extracts_angle_tags_and_preserves_token_fields(self):
        from phraseloom.tag_engine import TAG_CLOSE, TAG_OPEN, TAG_SELF, extract_tags

        result = extract_tags('<a href="shop">VIP10</a> <img src="coin.png"/>')

        self.assertEqual(result.text, "{1>VIP10<2} {3}")
        self.assertEqual(
            tuple(
                (tag.index, tag.kind, tag.placeholder, tag.raw, tag.partner_index)
                for tag in result.tags
            ),
            (
                (1, TAG_OPEN, "{1>", '<a href="shop">', None),
                (2, TAG_CLOSE, "<2}", "</a>", 1),
                (3, TAG_SELF, "{3}", '<img src="coin.png"/>', None),
            ),
        )
        self.assertEqual(result.warnings, ())

    def test_extracts_bbcode_tags(self):
        from phraseloom.tag_engine import extract_tags

        result = extract_tags("[color=#ff0]Bonus[/]")

        self.assertEqual(result.text, "{1>Bonus<2}")
        self.assertEqual(result.tags[0].raw, "[color=#ff0]")
        self.assertEqual(result.tags[1].raw, "[/]")
        self.assertEqual(result.tags[1].partner_index, 1)
        self.assertEqual(result.warnings, ())

    def test_extracts_named_bbcode_close_tags(self):
        from phraseloom.tag_engine import TAG_CLOSE, TAG_OPEN, extract_tags

        result = extract_tags("[b]Bold[/b]")

        self.assertEqual(result.text, "{1>Bold<2}")
        self.assertEqual(
            tuple(
                (tag.index, tag.kind, tag.placeholder, tag.raw, tag.partner_index)
                for tag in result.tags
            ),
            (
                (1, TAG_OPEN, "{1>", "[b]", None),
                (2, TAG_CLOSE, "<2}", "[/b]", 1),
            ),
        )
        self.assertEqual(result.warnings, ())

    def test_extracts_raw_brace_placeholders_as_single_tokens(self):
        from phraseloom.tag_engine import RAW_PLACEHOLDER, extract_tags

        result = extract_tags("Hit deals {0} damage to {player.name:N2}.")

        self.assertEqual(result.text, "Hit deals {1} damage to {2}.")
        self.assertEqual(
            tuple((tag.index, tag.kind, tag.placeholder, tag.raw) for tag in result.tags),
            (
                (1, RAW_PLACEHOLDER, "{1}", "{0}"),
                (2, RAW_PLACEHOLDER, "{2}", "{player.name:N2}"),
            ),
        )
        self.assertEqual(result.warnings, ())

    def test_extracts_tags_and_raw_braces_in_original_order(self):
        from phraseloom.tag_engine import RAW_PLACEHOLDER, TAG_CLOSE, TAG_OPEN, extract_tags

        result = extract_tags("[color=#1213] 击打造成{0}伤害[/]")

        self.assertEqual(result.text, "{1> 击打造成{2}伤害<3}")
        self.assertEqual(
            tuple(
                (tag.index, tag.kind, tag.placeholder, tag.raw, tag.partner_index)
                for tag in result.tags
            ),
            (
                (1, TAG_OPEN, "{1>", "[color=#1213]", None),
                (2, RAW_PLACEHOLDER, "{2}", "{0}", None),
                (3, TAG_CLOSE, "<3}", "[/]", 1),
            ),
        )

    def test_incomplete_raw_braces_stay_raw_without_warning(self):
        from phraseloom.tag_engine import extract_tags

        result = extract_tags("Use {abc")

        self.assertEqual(result.text, "Use {abc")
        self.assertEqual(result.tags, ())
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

        self.assertEqual(result.text, "{1>Text")
        self.assertEqual(len(result.tags), 1)
        self.assertEqual(result.tags[0].kind, TAG_OPEN)
        self.assertEqual(result.tags[0].raw, '<a href="x">')
        self.assertTrue(
            any("open tag has no close partner" in warning for warning in result.warnings)
        )

    def test_shorthand_angle_close_pairs_with_nearest_open(self):
        from phraseloom.tag_engine import extract_tags

        result = extract_tags("<a>Text</>")

        self.assertEqual(result.text, "{1>Text<2}")
        self.assertEqual(result.tags[1].raw, "</>")
        self.assertEqual(result.warnings, ())

    def test_extracts_existing_placeholders_as_raw_braces(self):
        from phraseloom.tag_engine import (
            RAW_PLACEHOLDER,
            TAG_CLOSE,
            TAG_OPEN,
            extract_tags,
            restore_tags,
        )

        result = extract_tags("{t1_op}<a>x</a>")

        self.assertEqual(result.text, "{1}{2>x<3}")
        self.assertEqual(
            tuple(
                (tag.index, tag.kind, tag.placeholder, tag.raw, tag.partner_index)
                for tag in result.tags
            ),
            (
                (1, RAW_PLACEHOLDER, "{1}", "{t1_op}", None),
                (2, TAG_OPEN, "{2>", "<a>", None),
                (3, TAG_CLOSE, "<3}", "</a>", 2),
            ),
        )
        self.assertEqual(result.warnings, ())
        self.assertEqual(restore_tags(result.text, result.tags), "{t1_op}<a>x</a>")

    def test_misnested_named_tags_do_not_cross_pair(self):
        from phraseloom.tag_engine import extract_tags

        result = extract_tags("<a><b>x</a>y</b>")

        self.assertEqual(result.text, "{1>{2>x</a>y<3}")
        self.assertEqual(
            [tag.placeholder for tag in result.tags],
            ["{1>", "{2>", "<3}"],
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

    def test_suspicious_raw_text_extracts_raw_brace_placeholder(self):
        from phraseloom.tag_engine import RAW_PLACEHOLDER, extract_tags

        result = extract_tags("Use 3 < 5 and keep {t1_op}")

        self.assertEqual(result.text, "Use 3 < 5 and keep {1}")
        self.assertEqual(
            tuple((tag.index, tag.kind, tag.placeholder, tag.raw) for tag in result.tags),
            ((1, RAW_PLACEHOLDER, "{1}", "{t1_op}"),),
        )
        self.assertEqual(result.warnings, ())

    def test_identifies_tag_only_segments(self):
        from phraseloom.tag_engine import is_tag_only_segment

        self.assertTrue(is_tag_only_segment("{3}"))
        self.assertTrue(is_tag_only_segment("{1><2}"))
        self.assertTrue(is_tag_only_segment("{1> <2}"))
        self.assertFalse(is_tag_only_segment("{1>Click<2}"))
        self.assertFalse(is_tag_only_segment("{1} 100 coins"))

    def test_restore_tags_replaces_known_placeholders_only(self):
        from phraseloom.tag_engine import extract_tags, restore_tags

        extraction = extract_tags('<a href="shop">VIP10</a>')

        restored = restore_tags("{1>VIP<2} {9}", extraction.tags)

        self.assertEqual(restored, '<a href="shop">VIP</a> {9}')

    def test_validate_tag_placeholders_reports_extra_counts(self):
        from phraseloom.tag_engine import extract_tags, validate_tag_placeholders

        extraction = extract_tags('<a href="shop">VIP10</a>')

        validation = validate_tag_placeholders(
            "{1>VIP10<2} {3} {4}", extraction.tags
        )

        self.assertEqual(validation.warnings, ("tag_mismatch: extra {4}",))

    def test_validate_tag_placeholders_reports_missing_counts(self):
        from phraseloom.tag_engine import extract_tags, validate_tag_placeholders

        extraction = extract_tags("<a>x</a>")

        validation = validate_tag_placeholders("{1>x", extraction.tags)

        self.assertEqual(validation.warnings, ("tag_mismatch: missing <2}",))

    def test_serialize_known_tags_preserves_repeated_raw_tags_in_order(self):
        from phraseloom.tag_engine import TAG_SELF, TagToken, serialize_known_tags

        tags = (
            TagToken(1, TAG_SELF, "{1}", "<br/>"),
            TagToken(2, TAG_SELF, "{2}", "<br/>"),
        )

        result = serialize_known_tags("<br/> A <br/>", tags)

        self.assertEqual(result.text, "{1} A {2}")
        self.assertEqual(result.tags, tags)
        self.assertEqual(result.warnings, ())

    def test_serialize_known_tags_returns_empty_extraction_for_empty_target(self):
        from phraseloom.tag_engine import TagExtraction, extract_tags, serialize_known_tags

        extraction = extract_tags("<a>x</a>")

        result = serialize_known_tags("", extraction.tags)

        self.assertEqual(result, TagExtraction("", (), ()))


if __name__ == "__main__":
    unittest.main()
