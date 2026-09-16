# -*- mode: python ; coding: utf-8 -*-
"""GUI-only macOS bundle; build with scripts/build_macos_release.py."""

from pathlib import Path
import platform
import tomllib

root = Path(SPECPATH).parent
version = tomllib.loads((root / "pyproject.toml").read_text())["project"]["version"]
excludes = [
    line.strip()
    for line in (root / "packaging/gui-excludes.txt").read_text().splitlines()
    if line.strip() and not line.lstrip().startswith("#")
]
a = Analysis(
    [str(root / "toolshub_gui.py")],
    pathex=[str(root)],
    binaries=[],
    datas=[(str(root / "phraseloom/tag_rules.toml"), "phraseloom")],
    hiddenimports=["PIL.Image", "ahocorasick"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, [],
    exclude_binaries=True,
    name="QAtools",
    debug=False,
    strip=False,
    upx=False,
    console=False,
    argv_emulation=False,
    target_arch=platform.machine(),
)
coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False, name="QAtools")
app = BUNDLE(
    coll,
    name="QAtools.app",
    icon=str(root / "packaging/QAtools-icon-rounded.png"),
    bundle_identifier="com.dorajackcui.qatools",
    version=version,
    info_plist={
        "CFBundleShortVersionString": version,
        "NSHighResolutionCapable": True,
        "LSMinimumSystemVersion": "13.0",
    },
)
