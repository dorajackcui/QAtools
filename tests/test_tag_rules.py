import tempfile
import unittest
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
                        'allowed = ["foo"]',
                        "",
                        "[bbcode_tags]",
                        'mode = "allowlist"',
                        'allowed = ["bar"]',
                        "",
                        "[raw_braces]",
                        "protect_all = false",
                    ]
                ),
                encoding="utf-8",
            )

            rules = load_tag_rules(path)

        self.assertTrue(rules.allows_angle("foo"))
        self.assertFalse(rules.allows_angle("color"))
        self.assertTrue(rules.allows_bbcode("bar"))
        self.assertFalse(rules.protect_raw_braces)

    def test_normalized_hash_ignores_order_and_case(self):
        from phraseloom.tag_rules import TagRules, normalized_tag_rules_hash

        left = TagRules(
            version=1,
            angle_allowed=frozenset({"color", "img"}),
            bbcode_allowed=frozenset({"b", "color"}),
            protect_raw_braces=True,
            source="left",
        )
        right = TagRules(
            version=1,
            angle_allowed=frozenset({"IMG", "COLOR"}),
            bbcode_allowed=frozenset({"COLOR", "B"}),
            protect_raw_braces=True,
            source="right",
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


if __name__ == "__main__":
    unittest.main()
