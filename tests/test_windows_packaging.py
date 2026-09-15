from __future__ import annotations

import struct
import os
import subprocess
import sys
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

    def test_release_build_wraps_fast_gui_bundle_in_one_installer(self) -> None:
        build_script = (
            PROJECT_ROOT / "scripts" / "build_windows_release.ps1"
        ).read_text(encoding="utf-8")

        self.assertIn('"--onedir"', build_script)
        self.assertIn('"--contents-directory", "_internal"', build_script)
        self.assertIn('$guiBundleDir = Join-Path $exeDir "QAtools"', build_script)
        self.assertNotIn('"--onefile"', build_script)
        self.assertNotIn('QAtools-CLI', build_script)
        self.assertIn('"packaging\\gui-excludes.txt"', build_script)
        self.assertIn('$installerName = "QAtools-v$Version-windows-$architecture-setup"', build_script)
        self.assertIn('& $innoSetupCompiler', build_script)
        self.assertNotIn('Compress-Archive', build_script)

    def test_installer_supports_per_user_overwrite_upgrades(self) -> None:
        installer_script = (
            PROJECT_ROOT / "packaging" / "QAtools.iss"
        ).read_text(encoding="utf-8")

        self.assertIn("AppId={{9A854BDD-9184-4B8D-9622-130C28E39182}", installer_script)
        self.assertIn("DefaultDirName={localappdata}\\Programs\\QAtools", installer_script)
        self.assertIn("PrivilegesRequired=lowest", installer_script)
        self.assertIn("UsePreviousAppDir=yes", installer_script)
        self.assertIn('Source: "{#SourceDir}\\*"', installer_script)

    def test_upgrade_cleanup_is_limited_to_retired_application_files(self) -> None:
        installer = (PROJECT_ROOT / "packaging" / "QAtools.iss").read_text(encoding="utf-8")
        section = installer.split("[InstallDelete]\n", 1)[1].split("\n[", 1)[0]
        paths = {line.split('Name: "', 1)[1].split('"', 1)[0]
                 for line in section.splitlines() if line.startswith("Type:")}
        self.assertEqual(paths, {
            r"{app}\QAtools-CLI.exe", r"{app}\QAtools-CLI.cmd",
            r"{app}\_internal\numpy", r"{app}\_internal\numpy.libs",
            r"{app}\_internal\numpy-*.dist-info", r"{app}\_internal\Pythonwin",
            r"{app}\_internal\yaml",
        })

    def test_real_gui_workflows_without_excluded_dependencies(self) -> None:
        code = (
            "import pathlib, sys, unittest\n"
            "for line in pathlib.Path('packaging/gui-excludes.txt').read_text().splitlines():\n"
            "    name = line.strip()\n"
            "    if name and not name.startswith('#'): sys.modules[name] = None\n"
            "suite = unittest.defaultTestLoader.discover('tests', pattern='test_gui_functional.py')\n"
            "result = unittest.TextTestRunner(verbosity=2).run(suite)\n"
            "raise SystemExit(not result.wasSuccessful())\n"
        )
        environment = dict(os.environ, QT_QPA_PLATFORM="offscreen")
        result = subprocess.run([sys.executable, "-c", code], cwd=PROJECT_ROOT,
                                env=environment, capture_output=True, text=True, timeout=240)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_project_icon_contains_multiple_windows_sizes(self) -> None:
        icon_data = (PROJECT_ROOT / "packaging" / "QAtools.ico").read_bytes()
        reserved, image_type, image_count = struct.unpack_from("<HHH", icon_data)

        self.assertEqual(reserved, 0)
        self.assertEqual(image_type, 1)
        self.assertGreaterEqual(image_count, 7)

    def test_every_icon_size_has_transparent_corners(self) -> None:
        from PySide6.QtGui import QImage

        data = (PROJECT_ROOT / "packaging" / "QAtools.ico").read_bytes()
        count = struct.unpack_from("<H", data, 4)[0]
        sizes = set()
        for index in range(count):
            length, offset = struct.unpack_from("<II", data, 6 + index * 16 + 8)
            image = QImage.fromData(data[offset:offset + length])
            self.assertFalse(image.isNull())
            size = image.width()
            sizes.add(size)
            with self.subTest(size=size):
                self.assertEqual(image.height(), size)
                for x, y in ((0, 0), (size - 1, 0), (0, size - 1), (size - 1, size - 1)):
                    # At 16 px the antialiased arc can partly cover a corner pixel.
                    self.assertLessEqual(image.pixelColor(x, y).alpha(), 20)
                self.assertEqual(image.pixelColor(size // 2, size // 2).alpha(), 255)
                self.assertGreater(image.pixelColor(size // 2, 0).alpha(), 250)
        self.assertEqual(sizes, {16, 20, 24, 32, 40, 48, 64, 128, 256})

    def test_installer_shortcuts_use_a_versioned_icon_asset(self) -> None:
        installer = (PROJECT_ROOT / "packaging" / "QAtools.iss").read_text(encoding="utf-8")
        self.assertIn('DestName: "QAtools-icon-{#AppVersion}.ico"', installer)
        shortcuts = installer.split("[Icons]\n", 1)[1].split("\n[", 1)[0]
        entries = [line for line in shortcuts.splitlines() if line.startswith("Name:")]
        self.assertEqual(len(entries), 2)
        for entry in entries:
            self.assertIn('IconFilename: "{app}\\QAtools-icon-{#AppVersion}.ico"', entry)


if __name__ == "__main__":
    unittest.main()
