#!/usr/bin/env python3
"""Compatibility entry point for the shared PySide6 workflow page."""

from __future__ import annotations

import sys
from pathlib import Path

if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parents[2]))


def main(argv: list[str] | None = None) -> int:
    from toolshub_gui import main as launch_gui

    return launch_gui(["--tool", "workflow", "--check", "line-break", *(argv or [])])


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
