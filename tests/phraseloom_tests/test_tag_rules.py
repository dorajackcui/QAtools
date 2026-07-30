import tempfile
import unittest
from unittest import mock
from pathlib import Path


class TagRulesTests(unittest.TestCase):
    def test_default_rules_allow_known_formatting_tags(self):
        from phraseloom.tag_rules import default_tag_rules

        rules = default_tag_rules()

        self.assertEqual(rules.version, 1)
        self.assertTrue(rules.allows_angle("color"))
        self.assertTrue(rules.allows_angle("COLOR"))
        self.assertTrue(rules.allows_angle("img"))
        self.assertTrue(rules.allows_angle("c"))
        self.assertFalse(rules.allows_angle("activate"))
        self.assertTrue(rules.allows_bbcode("color"))
        self.assertTrue(rules.protect_raw_braces)
        self.assertEqual(rules.canonical_angle("c"), "color")
        self.assertTrue(rules.is_angle_single("br"))
        self.assertTrue(rules.is_angle_single("img"))
        self.assertTrue(rules.is_angle_optional_pair("i"))
        self.assertTrue(rules.is_angle_optional_pair("outline"))
        self.assertFalse(rules.is_angle_optional_pair("u"))
        self.assertTrue(rules.allows_atomic_square("mq:rxt"))
        self.assertTrue(rules.allows_atomic_square("MQ:RXT"))
        self.assertFalse(rules.allows_atomic_square("mq:other"))

    def test_custom_rules_load_from_toml(self):
        from phraseloom.tag_rules import load_tag_rules

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "tag_rules.toml"
            path.write_text(
                "\n".join(
                    [
                        "version = 1",
                        "",
                        "[angle_tags]",
                        'mode = "allowlist"',
                        'allowed = ["foo", "bar", "br"]',
                        "",
                        "[angle_tags.aliases]",
                        'bar = "foo"',
                        "",
                        "[angle_tags.single]",
                        'tags = ["br"]',
                        "",
                        "[angle_tags.optional_pair]",
                        'tags = ["foo"]',
                        "",
                        "[bbcode_tags]",
                        'mode = "allowlist"',
                        'allowed = ["bar"]',
                        "",
                        "[atomic_square_tags]",
                        'mode = "allowlist"',
                        'allowed = ["MQ:RXT"]',
                        "",
                        "[raw_braces]",
                        "protect_all = false",
                    ]
                ),
                encoding="utf-8",
            )

            rules = load_tag_rules(path)

        self.assertTrue(rules.allows_angle("foo"))
        self.assertTrue(rules.allows_angle("bar"))
        self.assertFalse(rules.allows_angle("color"))
        self.assertEqual(rules.canonical_angle("bar"), "foo")
        self.assertTrue(rules.is_angle_single("br"))
        self.assertTrue(rules.is_angle_optional_pair("foo"))
        self.assertTrue(rules.allows_bbcode("bar"))
        self.assertTrue(rules.allows_atomic_square("mq:rxt"))
        self.assertFalse(rules.protect_raw_braces)

    def test_normalized_hash_ignores_order_and_case(self):
        from phraseloom.tag_rules import TagRules, normalized_tag_rules_hash

        left = TagRules(
            version=1,
            angle_allowed=frozenset({"color", "img"}),
            bbcode_allowed=frozenset({"b", "color"}),
            protect_raw_braces=True,
            source="left",
            atomic_square_allowed=frozenset({"mq:rxt"}),
        )
        right = TagRules(
            version=1,
            angle_allowed=frozenset({"IMG", "COLOR"}),
            bbcode_allowed=frozenset({"COLOR", "B"}),
            protect_raw_braces=True,
            source="right",
            atomic_square_allowed=frozenset({"MQ:RXT"}),
        )

        self.assertEqual(
            normalized_tag_rules_hash(left),
            normalized_tag_rules_hash(right),
        )

    def test_invalid_mode_reports_config_error(self):
        from phraseloom.errors import ConfigError
        from phraseloom.tag_rules import load_tag_rules

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "bad.toml"
            path.write_text(
                "\n".join(
                    [
                        "version = 1",
                        "",
                        "[angle_tags]",
                        'mode = "denylist"',
                        'allowed = ["color"]',
                        "",
                        "[bbcode_tags]",
                        'mode = "allowlist"',
                        'allowed = ["color"]',
                        "",
                        "[raw_braces]",
                        "protect_all = true",
                    ]
                ),
                encoding="utf-8",
            )

            with self.assertRaises(ConfigError) as raised:
                load_tag_rules(path)

        self.assertIn("angle_tags.mode must be 'allowlist'", str(raised.exception))

    def test_missing_custom_rules_file_reports_config_error(self):
        from phraseloom.errors import ConfigError
        from phraseloom.tag_rules import load_tag_rules

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "missing.toml"

            with self.assertRaises(ConfigError) as raised:
                load_tag_rules(path)

        self.assertIn("could not read tag rules", str(raised.exception))

    def test_invalid_utf8_custom_rules_file_reports_config_error(self):
        from phraseloom.errors import ConfigError
        from phraseloom.tag_rules import load_tag_rules

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "tag_rules.toml"
            path.write_bytes(b"\xff\xfe\xfa")

            with self.assertRaises(ConfigError) as raised:
                load_tag_rules(path)

        self.assertIn("could not read tag rules", str(raised.exception))

    def test_default_rules_toml_parse_failure_reports_config_error(self):
        from phraseloom import tag_rules
        from phraseloom.errors import ConfigError

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "tag_rules.toml"
            path.write_text("version = [", encoding="utf-8")

            tag_rules._default_tag_rules.cache_clear()
            with mock.patch.object(tag_rules.resources, "files", return_value=Path(tmp)):
                with self.assertRaises(ConfigError) as raised:
                    tag_rules.default_tag_rules()
            tag_rules._default_tag_rules.cache_clear()

        self.assertIn("invalid TOML", str(raised.exception))

    def test_boolean_version_reports_config_error(self):
        from phraseloom.errors import ConfigError
        from phraseloom.tag_rules import load_tag_rules

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "bad.toml"
            path.write_text(
                "\n".join(
                    [
                        "version = true",
                        "",
                        "[angle_tags]",
                        'mode = "allowlist"',
                        'allowed = ["color"]',
                        "",
                        "[bbcode_tags]",
                        'mode = "allowlist"',
                        'allowed = ["color"]',
                        "",
                        "[raw_braces]",
                        "protect_all = true",
                    ]
                ),
                encoding="utf-8",
            )

            with self.assertRaises(ConfigError) as raised:
                load_tag_rules(path)

        self.assertIn("tag rules version must be exactly 1", str(raised.exception))


if __name__ == "__main__":
    unittest.main()
