from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from tools.header_aliases import HeaderAliases, HeaderAliasStore


class HeaderAliasTests(unittest.TestCase):
    def test_create_trims_and_deduplicates_aliases_case_insensitively(self) -> None:
        aliases = HeaderAliases.create(
            source=(" Original ", "original", "SOURCE", ""),
            target=(" Translation ", "TRANSLATION", "target"),
        )

        self.assertEqual(aliases.source, ("Original",))
        self.assertEqual(aliases.target, ("Translation",))
        self.assertEqual(aliases.source_headers, frozenset({"source", "original"}))
        self.assertEqual(aliases.target_headers, frozenset({"target", "translation"}))

    def test_source_and_target_aliases_cannot_overlap(self) -> None:
        with self.assertRaisesRegex(ValueError, "不能重复"):
            HeaderAliases.create(source=("Text",), target=(" text ",))

        with self.assertRaisesRegex(ValueError, "source"):
            HeaderAliases.create(target=("Source",))

    def test_store_round_trips_aliases(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            config_path = Path(tmp_dir) / "header_aliases.json"
            store = HeaderAliasStore(config_path)

            store.save(
                HeaderAliases.create(
                    source=("原文", "English"),
                    target=("译文", "Chinese"),
                )
            )

            self.assertEqual(
                store.load(),
                HeaderAliases(
                    source=("原文", "English"),
                    target=("译文", "Chinese"),
                ),
            )
            payload = json.loads(config_path.read_text(encoding="utf-8"))
            self.assertEqual(payload["version"], 1)

    def test_store_reports_invalid_configuration(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            config_path = Path(tmp_dir) / "header_aliases.json"
            config_path.write_text("{invalid", encoding="utf-8")

            with self.assertRaisesRegex(ValueError, "无法读取表头别名配置"):
                HeaderAliasStore(config_path).load()


if __name__ == "__main__":
    unittest.main()
