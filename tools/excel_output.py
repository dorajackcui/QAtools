#!/usr/bin/env python3
"""Shared helpers for naming generated Excel output files."""

from __future__ import annotations

from pathlib import Path


def build_prefixed_output_path(input_file: str | Path, prefix: str) -> Path:
    input_path = Path(input_file).expanduser()
    return input_path.with_name(f"{prefix}{input_path.name}")
