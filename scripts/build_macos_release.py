#!/usr/bin/env python3
"""Build and smoke-test a GUI-only .app, then package it in a drag-install DMG."""

from __future__ import annotations

import argparse
import os
from pathlib import Path
import platform
import shutil
import subprocess
import sys
import tempfile
import tomllib


ROOT = Path(__file__).resolve().parents[1]


def run(*args: str, env=None, timeout=None) -> None:
    subprocess.run(args, cwd=ROOT, env=env, check=True, timeout=timeout)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--skip-tests", action="store_true", help="Skip source regression tests; always smoke-test the bundle.")
    args = parser.parse_args(argv)
    if sys.platform != "darwin":
        parser.error("This build must run on macOS.")
    architecture = platform.machine()
    if architecture not in {"arm64", "x86_64"}:
        parser.error(f"Unsupported Python architecture: {architecture}")
    version = tomllib.loads((ROOT / "pyproject.toml").read_text())["project"]["version"]
    release_name = f"QAtools-v{version}-macos-{architecture}"
    build_root = ROOT / "build/macos-release"
    dist_root = ROOT / "dist"
    build_root.mkdir(parents=True, exist_ok=True)
    dist_root.mkdir(parents=True, exist_ok=True)
    run(sys.executable, "-c", "import PyInstaller, PIL, ahocorasick, PySide6, openpyxl")
    environment = dict(os.environ, TMPDIR="/private/tmp")
    # Keep PyInstaller's cache inside the build tree as well.
    environment["PYINSTALLER_CONFIG_DIR"] = str(build_root / "cache")
    if not args.skip_tests:
        run(sys.executable, "-m", "unittest", "discover", "-s", "tests", "-v", env=environment)
    run(sys.executable, "-m", "compileall", "-q", "qatools", "phraseloom", "tools", "tests")
    run(sys.executable, "scripts/check_docs.py")
    run("git", "diff", "--check")
    run(
        sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean",
        "--distpath", str(build_root / "bundle"),
        "--workpath", str(build_root / "work"),
        str(ROOT / "packaging/QAtools-macos.spec"), env=environment,
    )
    app = build_root / "bundle/QAtools.app"
    executable = app / "Contents/MacOS/QAtools"
    with tempfile.TemporaryDirectory(prefix="qatools-mac-smoke-", dir="/private/tmp") as temporary:
        smoke_environment = dict(environment, QT_QPA_PLATFORM="offscreen", TMPDIR=temporary)
        run(str(executable), "--smoke-test", env=smoke_environment, timeout=90)
    run("/usr/bin/codesign", "--verify", "--deep", "--strict", str(app))
    # ditto preserves the bundle's framework symlinks and executable permissions.
    with tempfile.TemporaryDirectory(prefix="dmg-", dir=build_root) as temporary:
        staging = Path(temporary)
        run("/usr/bin/ditto", str(app), str(staging / "QAtools.app"))
        (staging / "Applications").symlink_to("/Applications", target_is_directory=True)
        shutil.copy2(ROOT / "packaging/README-macOS.txt", staging / "README.txt")
        # The source directory must not contain the image being written.
        with tempfile.TemporaryDirectory(prefix="image-", dir=build_root) as image_directory:
            image = Path(image_directory) / "release.dmg"
            run("/usr/bin/hdiutil", "create", "-volname", "QAtools", "-srcfolder", str(staging),
                "-format", "UDZO", str(image))
            run("/usr/bin/hdiutil", "verify", str(image))
            destination = dist_root / f"{release_name}.dmg"
            image.replace(destination)
    print(f"App: {app}\nDMG: {destination}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
