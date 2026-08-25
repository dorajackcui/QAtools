import unittest


class TagEngineTests(unittest.TestCase):
    def test_protected_token_helpers_accept_new_contract(self):
        from phraseloom.tag_engine import (
            RAW_MARKER,
            RAW_PLACEHOLDER,
            TAG_CLOSE,
            TAG_OPEN,
            TAG_SELF,
            is_protected_token,
            make_protected_token,
            parse_protected_token,
        )

        self.assertEqual(TAG_OPEN, "op")
        self.assertEqual(TAG_CLOSE, "cl")
        self.assertEqual(TAG_SELF, "sf")
        self.assertEqual(RAW_PLACEHOLDER, "ph")
        self.assertEqual(RAW_MARKER, "mk")
        self.assertEqual(make_protected_token(1, TAG_OPEN), "{1>")
        self.assertEqual(make_protected_token(2, TAG_CLOSE), "<2}")
        self.assertEqual(make_protected_token(3, TAG_SELF), "{3}")
        self.assertEqual(make_protected_token(4, RAW_PLACEHOLDER), "{4}")
        self.assertEqual(make_protected_token(5, RAW_MARKER), "{5}")
        self.assertTrue(is_protected_token("{1>"))
        self.assertTrue(is_protected_token("<2}"))
        self.assertTrue(is_protected_token("{3}"))
        self.assertFalse(is_protected_token("{num1}"))
        self.assertFalse(is_protected_token("{t1_op}"))
        self.assertEqual(parse_protected_token("{1>"), (1, TAG_OPEN))
        self.assertEqual(parse_protected_token("<2}"), (2, TAG_CLOSE))
        self.assertEqual(parse_protected_token("{3}"), (3, "single"))
        self.assertIsNone(parse_protected_token("{num1}"))
        self.assertIsNone(parse_protected_token("{t1_op}"))

    def test_protected_token_scans_accept_all_token_shapes(self):
        from phraseloom.tag_engine import (
            RAW_PLACEHOLDER,
            extract_tags,
            validate_tag_placeholders,
        )

        extraction = extract_tags("Keep {3} literal")
        validation = validate_tag_placeholders("{1><2}{3}", ())

        self.assertEqual(extraction.text, "Keep {1} literal")
        self.assertEqual(
            tuple((tag.index, tag.kind, tag.placeholder, tag.raw) for tag in extraction.tags),
            ((1, RAW_PLACEHOLDER, "{1}", "{3}"),),
        )
        self.assertEqual(
            validation.warnings,
            (
                "protected_token_mismatch: extra {1>",
                "protected_token_mismatch: extra <2}",
                "protected_token_mismatch: extra {3}",
            ),
        )

    def test_extracts_angle_tags_and_preserves_token_fields(self):
        from phraseloom.tag_engine import TAG_CLOSE, TAG_OPEN, TAG_SELF, extract_tags

        result = extract_tags('<color=#fff>VIP10</color> <img src="coin.png"/>')

        self.assertEqual(result.text, "{1>VIP10<2} {3}")
        self.assertEqual(
            tuple(
                (tag.index, tag.kind, tag.placeholder, tag.raw, tag.partner_index)
                for tag in result.tags
            ),
            (
                (1, TAG_OPEN, "{1>", "<color=#fff>", None),
                (2, TAG_CLOSE, "<2}", "</color>", 1),
                (3, TAG_SELF, "{3}", '<img src="coin.png"/>', None),
            ),
        )
        self.assertEqual(result.warnings, ())

    def test_default_rules_keep_unlisted_angle_label_raw(self):
        from phraseloom.tag_engine import RAW_PLACEHOLDER, extract_tags

        result = extract_tags("<Activate> HP increased by {a}%")

        self.assertEqual(result.text, "<Activate> HP increased by {1}%")
        self.assertEqual(
            tuple((tag.kind, tag.placeholder, tag.raw) for tag in result.tags),
            ((RAW_PLACEHOLDER, "{1}", "{a}"),),
        )
        self.assertEqual(result.warnings, ())

    def test_default_rules_keep_unlisted_angle_label_with_spaces_raw(self):
        from phraseloom.tag_engine import RAW_PLACEHOLDER, extract_tags

        result = extract_tags("<Weather Change> before turn {a}")

        self.assertEqual(result.text, "<Weather Change> before turn {1}")
        self.assertEqual(
            tuple((tag.kind, tag.placeholder, tag.raw) for tag in result.tags),
            ((RAW_PLACEHOLDER, "{1}", "{a}"),),
        )
        self.assertEqual(result.warnings, ())

    def test_default_rules_still_extract_allowed_color_pair(self):
        from phraseloom.tag_engine import RAW_PLACEHOLDER, TAG_CLOSE, TAG_OPEN, extract_tags

        result = extract_tags("<color=#123>HP {a}</>")

        self.assertEqual(result.text, "{1>HP {2}<3}")
        self.assertEqual(
            tuple((tag.kind, tag.placeholder, tag.raw, tag.partner_index) for tag in result.tags),
            (
                (TAG_OPEN, "{1>", "<color=#123>", None),
                (RAW_PLACEHOLDER, "{2}", "{a}", None),
                (TAG_CLOSE, "<3}", "</>", 1),
            ),
        )
        self.assertEqual(result.warnings, ())

    def test_default_rules_extract_span_and_hyperlink_pairs(self):
        from phraseloom.tag_engine import TAG_CLOSE, TAG_OPEN, extract_tags

        source = (
            '<span color="#BFF8FA" size="36">Title</>'
            '<hyperlink color="#E98845" action="Key39">Open</>'
        )

        result = extract_tags(source)

        self.assertEqual(result.text, "{1>Title<2}{3>Open<4}")
        self.assertEqual(
            tuple(
                (tag.kind, tag.placeholder, tag.raw, tag.partner_index)
                for tag in result.tags
            ),
            (
                (
                    TAG_OPEN,
                    "{1>",
                    '<span color="#BFF8FA" size="36">',
                    None,
                ),
                (TAG_CLOSE, "<2}", "</>", 1),
                (
                    TAG_OPEN,
                    "{3>",
                    '<hyperlink color="#E98845" action="Key39">',
                    None,
                ),
                (TAG_CLOSE, "<4}", "</>", 3),
            ),
        )
        self.assertEqual(result.warnings, ())

    def test_default_rules_leave_unlisted_named_close_raw(self):
        from phraseloom.tag_engine import extract_tags

        result = extract_tags("<foo>x</foo>")

        self.assertEqual(result.text, "<foo>x</foo>")
        self.assertEqual(result.tags, ())
        self.assertEqual(result.warnings, ())

    def test_custom_rules_can_leave_raw_braces_unprotected(self):
        from phraseloom.tag_engine import extract_tags
        from phraseloom.tag_rules import TagRules

        rules = TagRules(1, frozenset({"color"}), frozenset({"color"}), False)
        result = extract_tags("<color>HP {a}</>", rules=rules)
        self.assertEqual(result.text, "{1>HP {a}<2}")

    def test_default_rules_leave_unlisted_anchor_raw(self):
        from phraseloom.tag_engine import extract_tags

        result = extract_tags('<a href="shop">VIP10</a>')

        self.assertEqual(result.text, '<a href="shop">VIP10</a>')
        self.assertEqual(result.tags, ())
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

    def test_literal_line_markers_are_protected_and_restore_exactly(self):
        from phraseloom.tag_engine import (
            RAW_MARKER,
            extract_tags,
            restore_tags,
            serialize_known_tags,
        )

        source = r"First\nSecond\rThird"
        extraction = extract_tags(source)

        self.assertEqual(extraction.text, "First{1}Second{2}Third")
        self.assertEqual(
            tuple((tag.kind, tag.placeholder, tag.raw) for tag in extraction.tags),
            (
                (RAW_MARKER, "{1}", r"\n"),
                (RAW_MARKER, "{2}", r"\r"),
            ),
        )
        self.assertEqual(restore_tags(extraction.text, extraction.tags), source)

        serialized = serialize_known_tags(
            r"Premier\nDeuxième\rTroisième",
            extraction.tags,
        )
        self.assertEqual(serialized.text, "Premier{1}Deuxième{2}Troisième")
        self.assertEqual(
            restore_tags(serialized.text, extraction.tags),
            r"Premier\nDeuxième\rTroisième",
        )

    def test_actual_line_breaks_are_not_literal_line_markers(self):
        from phraseloom.tag_engine import extract_tags

        source = "First\nSecond\rThird"
        extraction = extract_tags(source)

        self.assertEqual(extraction.text, source)
        self.assertEqual(extraction.tags, ())

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

        result = extract_tags("</color>Text")

        self.assertEqual(result.text, "</color>Text")
        self.assertEqual(result.tags, ())
        self.assertTrue(any("unpaired close" in warning for warning in result.warnings))

    def test_unclosed_open_serializes_and_warns(self):
        from phraseloom.tag_engine import TAG_OPEN, extract_tags

        result = extract_tags("<color=#fff>Text")

        self.assertEqual(result.text, "{1>Text")
        self.assertEqual(len(result.tags), 1)
        self.assertEqual(result.tags[0].kind, TAG_OPEN)
        self.assertEqual(result.tags[0].raw, "<color=#fff>")
        self.assertTrue(
            any("open tag has no close partner" in warning for warning in result.warnings)
        )

    def test_shorthand_angle_close_pairs_with_nearest_open(self):
        from phraseloom.tag_engine import extract_tags

        result = extract_tags("<color>Text</>")

        self.assertEqual(result.text, "{1>Text<2}")
        self.assertEqual(result.tags[1].raw, "</>")
        self.assertEqual(result.warnings, ())

    def test_angle_alias_close_pairs_with_canonical_open(self):
        from phraseloom.tag_engine import TAG_CLOSE, TAG_OPEN, extract_tags

        result = extract_tags("<color=#{a}C{b}F>Text</c>")

        self.assertEqual(result.text, "{1>Text<2}")
        self.assertEqual(
            tuple((tag.kind, tag.placeholder, tag.raw, tag.partner_index) for tag in result.tags),
            (
                (TAG_OPEN, "{1>", "<color=#{a}C{b}F>", None),
                (TAG_CLOSE, "<2}", "</c>", 1),
            ),
        )
        self.assertEqual(result.warnings, ())

    def test_configured_single_angle_open_serializes_without_close_warning(self):
        from phraseloom.tag_engine import TAG_SELF, extract_tags

        result = extract_tags('<br>Line<img src="coin.png">')

        self.assertEqual(result.text, "{1}Line{2}")
        self.assertEqual(
            tuple((tag.kind, tag.placeholder, tag.raw) for tag in result.tags),
            (
                (TAG_SELF, "{1}", "<br>"),
                (TAG_SELF, "{2}", '<img src="coin.png">'),
            ),
        )
        self.assertEqual(result.warnings, ())

    def test_optional_pair_angle_open_without_named_close_serializes_as_single(self):
        from phraseloom.tag_engine import TAG_CLOSE, TAG_OPEN, TAG_SELF, extract_tags

        result = extract_tags("<i><size={a}>Voir</size><size={b}> division</size>")

        self.assertEqual(result.text, "{1}{2>Voir<3}{4> division<5}")
        self.assertEqual(
            tuple((tag.kind, tag.placeholder, tag.raw, tag.partner_index) for tag in result.tags),
            (
                (TAG_SELF, "{1}", "<i>", None),
                (TAG_OPEN, "{2>", "<size={a}>", None),
                (TAG_CLOSE, "<3}", "</size>", 2),
                (TAG_OPEN, "{4>", "<size={b}>", None),
                (TAG_CLOSE, "<5}", "</size>", 4),
            ),
        )
        self.assertEqual(result.warnings, ())

    def test_default_outline_without_named_close_serializes_as_single(self):
        from phraseloom.tag_engine import TAG_CLOSE, TAG_OPEN, TAG_SELF, extract_tags

        result = extract_tags(
            "<outline color=#{a}C{b} width={c}><color=#FFFDF{d}>Time</c>"
        )

        self.assertEqual(result.text, "{1}{2>Time<3}")
        self.assertEqual(
            tuple((tag.kind, tag.placeholder, tag.raw, tag.partner_index) for tag in result.tags),
            (
                (TAG_SELF, "{1}", "<outline color=#{a}C{b} width={c}>", None),
                (TAG_OPEN, "{2>", "<color=#FFFDF{d}>", None),
                (TAG_CLOSE, "<3}", "</c>", 2),
            ),
        )
        self.assertEqual(result.warnings, ())

    def test_optional_pair_angle_open_with_named_close_stays_paired(self):
        from phraseloom.tag_engine import TAG_CLOSE, TAG_OPEN, extract_tags

        result = extract_tags("<i>Voir</i>")

        self.assertEqual(result.text, "{1>Voir<2}")
        self.assertEqual(
            tuple((tag.kind, tag.placeholder, tag.raw, tag.partner_index) for tag in result.tags),
            (
                (TAG_OPEN, "{1>", "<i>", None),
                (TAG_CLOSE, "<2}", "</i>", 1),
            ),
        )
        self.assertEqual(result.warnings, ())

    def test_extracts_existing_placeholders_as_raw_braces(self):
        from phraseloom.tag_engine import (
            RAW_PLACEHOLDER,
            TAG_CLOSE,
            TAG_OPEN,
            extract_tags,
            restore_tags,
        )

        result = extract_tags("{t1_op}<color>x</color>")

        self.assertEqual(result.text, "{1}{2>x<3}")
        self.assertEqual(
            tuple(
                (tag.index, tag.kind, tag.placeholder, tag.raw, tag.partner_index)
                for tag in result.tags
            ),
            (
                (1, RAW_PLACEHOLDER, "{1}", "{t1_op}", None),
                (2, TAG_OPEN, "{2>", "<color>", None),
                (3, TAG_CLOSE, "<3}", "</color>", 2),
            ),
        )
        self.assertEqual(result.warnings, ())
        self.assertEqual(restore_tags(result.text, result.tags), "{t1_op}<color>x</color>")

    def test_misnested_named_tags_do_not_cross_pair(self):
        from phraseloom.tag_engine import extract_tags

        result = extract_tags("<color><size>x</color>y</size>")

        self.assertEqual(result.text, "{1>{2>x</color>y<3}")
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

    def test_restore_tags_replaces_known_protected_tokens_only(self):
        from phraseloom.tag_engine import extract_tags, restore_tags

        extraction = extract_tags("<color=#fff>VIP{0}</color>")

        restored = restore_tags("{1>VIP{2}<3} {9}", extraction.tags)

        self.assertEqual(restored, "<color=#fff>VIP{0}</color> {9}")

    def test_restore_tags_does_not_cascade_raw_brace_replacements(self):
        from phraseloom.tag_engine import extract_tags, restore_tags

        extraction = extract_tags("The No. {2} grand prize winner: {1}")

        restored = restore_tags("Gagnant du gros lot n° {1} : {2}", extraction.tags)

        self.assertEqual(restored, "Gagnant du gros lot n° {2} : {1}")

    def test_mq_rxt_tags_are_atomic_and_restore_exactly(self):
        from phraseloom.tag_engine import (
            RAW_PLACEHOLDER,
            TAG_SELF,
            extract_tags,
            restore_tags,
        )

        source = (
            r'在[mq:rxt displaytext="<span color=\&quot;{color1}\&quot;>" '
            r'val="<span color=\&quot;{color2}\&quot;>"]精英难度下的破碎中枢'
            r'[mq:rxt displaytext="</>" val="</>"]中，使用{num1}次'
            r'[mq:rxt displaytext="<hyperlink color=\&quot;{color3}\&quot; '
            r'action=\&quot;{num2}\&quot;>" val="<hyperlink color=\&quot;'
            r'{color4}\&quot; action=\&quot;{num3}\&quot;>"]'
            r'[mq:rxt displaytext="\{1}" val="\{2}"]'
            r'[mq:rxt displaytext="</>" val="</>"]'
            r'([mq:rxt displaytext="\\{3}" val="\\{4}"]/'
            r'[mq:rxt displaytext="\\{5}" val="\\{6}"])'
        )

        extraction = extract_tags(source)

        self.assertEqual(
            extraction.text,
            "在{1}精英难度下的破碎中枢{2}中，使用{3}次{4}{5}{6}({7}/{8})",
        )
        self.assertEqual(
            [tag.kind for tag in extraction.tags],
            [
                TAG_SELF,
                TAG_SELF,
                RAW_PLACEHOLDER,
                TAG_SELF,
                TAG_SELF,
                TAG_SELF,
                TAG_SELF,
                TAG_SELF,
            ],
        )
        self.assertEqual(extraction.tags[2].raw, "{num1}")
        self.assertEqual(extraction.warnings, ())
        self.assertEqual(restore_tags(extraction.text, extraction.tags), source)

    def test_mq_rxt_scanner_ignores_closing_bracket_inside_quotes(self):
        from phraseloom.tag_engine import extract_tags

        source = '[mq:rxt displaytext="left]right" val="value"]tail'

        extraction = extract_tags(source)

        self.assertEqual(extraction.text, "{1}tail")
        self.assertEqual(extraction.tags[0].raw, source[:-4])

    def test_validate_tag_placeholders_reports_extra_counts(self):
        from phraseloom.tag_engine import extract_tags, validate_tag_placeholders

        extraction = extract_tags("<color=#fff>VIP10</color>")

        validation = validate_tag_placeholders("{1>VIP10<2} {3}", extraction.tags)

        self.assertEqual(validation.warnings, ("protected_token_mismatch: extra {3}",))

    def test_validate_tag_placeholders_reports_missing_counts(self):
        from phraseloom.tag_engine import extract_tags, validate_tag_placeholders

        extraction = extract_tags("<color>x</color>")

        validation = validate_tag_placeholders("{1>x", extraction.tags)

        self.assertEqual(validation.warnings, ("protected_token_mismatch: missing <2}",))

    def test_self_closing_tag_validates_against_protected_token(self):
        from phraseloom.tag_engine import extract_tags, validate_tag_placeholders

        extraction = extract_tags("<br/>")

        validation = validate_tag_placeholders("{1}", extraction.tags)

        self.assertEqual(validation.warnings, ())

    def test_serialize_known_tags_preserves_repeated_raw_spans_in_order(self):
        from phraseloom.tag_engine import TAG_SELF, TagToken, serialize_known_tags

        tags = (
            TagToken(1, TAG_SELF, "{1}", "<br/>"),
            TagToken(2, TAG_SELF, "{2}", "<br/>"),
        )

        result = serialize_known_tags("<br/> A <br/>", tags)

        self.assertEqual(result.text, "{1} A {2}")
        self.assertEqual(result.tags, tags)
        self.assertEqual(result.warnings, ())

    def test_serialize_known_tags_reserves_explicit_token_before_duplicate_raw_tag(self):
        from phraseloom.tag_engine import TAG_SELF, TagToken, serialize_known_tags

        tags = (
            TagToken(1, TAG_SELF, "{1}", "<br/>"),
            TagToken(2, TAG_SELF, "{2}", "<br/>"),
        )

        result = serialize_known_tags("{1} A <br/>", tags)

        self.assertEqual(result.text, "{1} A {2}")
        self.assertEqual(result.tags, (tags[1],))
        self.assertEqual(result.warnings, ())

    def test_serialize_known_tags_does_not_cascade_raw_brace_replacements(self):
        from phraseloom.tag_engine import extract_tags, serialize_known_tags

        extraction = extract_tags("饱食度：{0}/{1}")

        result = serialize_known_tags("Satiété : {0}/{1}", extraction.tags)

        self.assertEqual(result.text, "Satiété : {1}/{2}")
        self.assertEqual(result.tags, extraction.tags)
        self.assertEqual(result.warnings, ())

    def test_serialize_known_tags_warns_when_source_span_not_found(self):
        from phraseloom.tag_engine import extract_tags, serialize_known_tags

        extraction = extract_tags("<br/>")

        result = serialize_known_tags("no matching span", extraction.tags)

        self.assertEqual(result.text, "no matching span")
        self.assertEqual(
            result.warnings,
            ("source_protected_span_not_found: <br/>",),
        )

    def test_nested_raw_braces_are_not_extracted_as_one_placeholder(self):
        from phraseloom.tag_engine import RAW_PLACEHOLDER, extract_tags

        result = extract_tags("Use {a{b}c}")

        self.assertEqual(result.text, "Use {a{1}c}")
        self.assertEqual(
            tuple((tag.kind, tag.placeholder, tag.raw) for tag in result.tags),
            ((RAW_PLACEHOLDER, "{1}", "{b}"),),
        )

    def test_serialize_known_tags_returns_empty_extraction_for_empty_target(self):
        from phraseloom.tag_engine import TagExtraction, extract_tags, serialize_known_tags

        extraction = extract_tags("<color>x</color>")

        result = serialize_known_tags("", extraction.tags)

        self.assertEqual(result, TagExtraction("", (), ()))


if __name__ == "__main__":
    unittest.main()
