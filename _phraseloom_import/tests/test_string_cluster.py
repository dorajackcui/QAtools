import unittest

from phraseloom.string_cluster import cluster_similar_strings


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

    def test_requires_at_least_two_members(self) -> None:
        with self.assertRaisesRegex(ValueError, "at least 2"):
            cluster_similar_strings(["One"], min_group_size=1)


if __name__ == "__main__":
    unittest.main()
