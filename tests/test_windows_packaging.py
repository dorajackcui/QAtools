from __future__ import annotations

import struct
import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]


class WindowsPackagingTests(unittest.TestCase):
    def test_icon_source_uses_the_application_theme_colors(self) -> None:
        icon_source = (
            PROJECT_ROOT / "packaging" / "QAtools-icon.svg"
        ).read_text(encoding="utf-8").casefold()

        self.assertIn("#cc7d5e", icon_source)
        self.assertIn("#f9f9f7", icon_source)
        self.assertIn('id="q-mark"', icon_source)

    def test_release_build_uses_the_project_icon(self) -> None:
        build_script = (
            PROJECT_ROOT / "scripts" / "build_windows_release.ps1"
        ).read_text(encoding="utf-8")

        self.assertIn('"packaging\\QAtools.ico"', build_script)
        self.assertIn('"--icon", $iconPath', build_script)

    def test_project_icon_contains_multiple_windows_sizes(self) -> None:
        icon_data = (PROJECT_ROOT / "packaging" / "QAtools.ico").read_bytes()
        reserved, image_type, image_count = struct.unpack_from("<HHH", icon_data)

        self.assertEqual(reserved, 0)
        self.assertEqual(image_type, 1)
        self.assertGreaterEqual(image_count, 7)


if __name__ == "__main__":
    unittest.main()
