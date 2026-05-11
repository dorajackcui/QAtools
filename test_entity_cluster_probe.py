import unittest


class EntityClusterProbeTests(unittest.TestCase):
    def test_groups_one_entity_slot_when_surrounding_structure_matches(self):
        from entity_cluster_probe import find_entity_clusters

        rows = [
            ("Squirtle launched an attack and dealt damage.", "Carapuce a lancé une attaque et infligé des dégâts."),
            ("Pikachu launched an attack and dealt damage.", "Pikachu a lancé une attaque et infligé des dégâts."),
            ("Bulbasaur launched an attack and dealt damage.", "Bulbizarre a lancé une attaque et infligé des dégâts."),
            ("Login failed.", "Échec de la connexion."),
        ]

        clusters = find_entity_clusters(rows, min_group_size=2)
        by_pattern = {cluster.source_pattern: cluster for cluster in clusters}

        self.assertIn("{entity1} launched an attack and dealt damage.", by_pattern)
        cluster = by_pattern["{entity1} launched an attack and dealt damage."]
        self.assertEqual(cluster.coverage_count, 3)
        self.assertEqual(
            cluster.entity_values,
            ("Bulbasaur", "Pikachu", "Squirtle"),
        )
        self.assertGreaterEqual(cluster.confidence, 0.8)

    def test_ignores_numeric_only_variants(self):
        from entity_cluster_probe import find_entity_clusters

        rows = [
            ("Reach level 10", "Atteindre le niveau 10"),
            ("Reach level 20", "Atteindre le niveau 20"),
            ("Reach level 30", "Atteindre le niveau 30"),
        ]

        clusters = find_entity_clusters(rows, min_group_size=2)

        self.assertEqual(clusters, [])

    def test_groups_cjk_sentence_with_two_entity_slots(self):
        from entity_cluster_probe import find_entity_clusters

        rows = [
            (
                '装备<span color="{color1}">「蓝晶兽外观」</>后，在演示中切换蓝晶兽登场时，会播放专属入场效果。',
                '<span color="{color1}">「Bluebeast style」</> is equipped; switching to Bluebeast plays a unique entry effect.',
            ),
            (
                '装备<span color="{color1}">「红翼猫外观」</>后，在演示中切换红翼猫登场时，会播放专属入场效果。',
                '<span color="{color1}">「Redcat style」</> is equipped; switching to Redcat plays a unique entry effect.',
            ),
            (
                '装备<span color="{color1}">「银叶鸟外观」</>后，在演示中切换银叶鸟登场时，会播放专属入场效果。',
                '<span color="{color1}">「Silverbird style」</> is equipped; switching to Silverbird plays a unique entry effect.',
            ),
        ]

        clusters = find_entity_clusters(rows, min_group_size=3, max_entity_tokens=8)
        by_pattern = {cluster.source_pattern: cluster for cluster in clusters}
        expected = '装备<span color="{color1}">「{entity1}外观」</>后，在演示中切换{entity2}登场时，会播放专属入场效果。'

        self.assertIn(expected, by_pattern)
        cluster = by_pattern[expected]
        self.assertEqual(cluster.coverage_count, 3)
        self.assertEqual(cluster.unique_entity_count, 3)
        self.assertIn("蓝晶兽 / 蓝晶兽", cluster.entity_values)

    def test_prefers_broader_cjk_entity_boundary_over_suffix_split(self):
        from entity_cluster_probe import find_entity_clusters

        rows = [
            ("已解锁展示动作。进入演示时，可以和蓝晶兽进行互动。", ""),
            ("已解锁展示动作。进入演示时，可以和红蘑菇进行互动。", ""),
            ("已解锁展示动作。进入演示时，可以和银影蘑菇进行互动。", ""),
            ("已解锁展示动作。进入演示时，可以和黑帽蘑菇进行互动。", ""),
        ]

        clusters = find_entity_clusters(rows, min_group_size=3, max_entity_tokens=8)
        patterns = {cluster.source_pattern for cluster in clusters}

        self.assertIn(
            "已解锁展示动作。进入演示时，可以和{entity1}进行互动。",
            patterns,
        )
        self.assertNotIn(
            "已解锁展示动作。进入演示时，可以和{entity1}蘑菇进行互动。",
            patterns,
        )
        self.assertNotIn(
            "已解锁展示动作。进入演示时，{entity1}进行互动。",
            patterns,
        )


if __name__ == "__main__":
    unittest.main()
