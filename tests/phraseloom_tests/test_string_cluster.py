import unittest

from phraseloom.string_cluster import (
    _structurally_similar,
    cluster_similar_strings,
)


class StringClusterTests(unittest.TestCase):
    def test_groups_english_strings_with_shared_structure(self) -> None:
        clusters = cluster_similar_strings(
            [
                "Pikachu launched an attack",
                "Squirtle launched an attack",
                "Bulbasaur launched an attack",
                "Open settings",
            ]
        )

        self.assertEqual(len(clusters), 1)
        self.assertEqual(clusters[0].member_indexes, (0, 1, 2))

    def test_groups_cjk_strings_with_shared_structure(self) -> None:
        clusters = cluster_similar_strings(
            [
                "皮卡丘发动了攻击",
                "杰尼龟发动了攻击",
                "妙蛙种子发动了攻击",
            ]
        )

        self.assertEqual(len(clusters), 1)
        self.assertEqual(clusters[0].member_indexes, (0, 1, 2))

    def test_leaves_unrelated_strings_ungrouped(self) -> None:
        clusters = cluster_similar_strings(
            ["Open settings", "Delete account", "Connection failed"]
        )

        self.assertEqual(clusters, [])

    def test_ignores_generated_placeholders_as_similarity_signal(self) -> None:
        clusters = cluster_similar_strings(
            [
                "\u901a\u7528\u8865\u507f\u5668LV{num1}",
                "\u91ce\u5fc3\u5bb6{num1}",
                "\u5e38\u660e\u706f{num1}",
            ]
        )

        self.assertEqual(clusters, [])

    def test_keeps_meaningful_structure_around_placeholders(self) -> None:
        clusters = cluster_similar_strings(
            [
                "Craft pistol: duration reduced {num1}%",
                "Craft rifle: duration reduced {num1}%",
                "Craft shotgun: duration reduced {num1}%",
            ]
        )

        self.assertEqual(len(clusters), 1)
        self.assertEqual(clusters[0].member_indexes, (0, 1, 2))

    def test_does_not_match_short_label_to_long_description(self) -> None:
        self.assertFalse(
            _structurally_similar(
                "North Laboratory",
                (
                    "North Laboratory recovered ancient technical blueprints "
                    "and asked explorers to collect more field data."
                ),
                min_confidence=0.7,
            )
        )

    def test_does_not_merge_similarity_chains(self) -> None:
        sources = [
            "Open inventory panel",
            "Open inventory panel and close settings menu",
            "Close settings menu",
        ]

        self.assertTrue(
            _structurally_similar(
                sources[0],
                sources[1],
                min_confidence=0.7,
            )
        )
        self.assertTrue(
            _structurally_similar(
                sources[1],
                sources[2],
                min_confidence=0.7,
            )
        )
        self.assertFalse(
            _structurally_similar(
                sources[0],
                sources[2],
                min_confidence=0.7,
            )
        )
        self.assertEqual(cluster_similar_strings(sources), [])

    def test_groups_raw_numeric_variants_by_exact_template(self) -> None:
        clusters = cluster_similar_strings(
            [
                "Level 1 unlocked",
                "Level 2 unlocked",
                "Level 3 unlocked",
            ]
        )

        self.assertEqual(len(clusters), 1)
        self.assertEqual(clusters[0].source_pattern, "Level {num1} unlocked")
        self.assertEqual(clusters[0].member_indexes, (0, 1, 2))

    def test_requires_at_least_two_members(self) -> None:
        with self.assertRaisesRegex(ValueError, "at least 2"):
            cluster_similar_strings(["One"], min_group_size=1)


if __name__ == "__main__":
    unittest.main()
