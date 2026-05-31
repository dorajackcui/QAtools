#!/usr/bin/env python3
"""Small shared helpers for Tkinter GUI modules."""

from __future__ import annotations


def parse_positive_int(raw_value: str, *, default: int, field_name: str) -> int:
    value = raw_value.strip() or str(default)
    try:
        parsed = int(value)
    except ValueError as exc:
        raise ValueError(f"{field_name}必须是整数。") from exc
    if parsed < 1:
        raise ValueError(f"{field_name}必须大于 0。")
    return parsed


def set_combobox_values(combobox, values: tuple[str, ...], variable, default_value: str | None) -> str:
    combobox["values"] = values
    selected = variable.get().strip()
    if selected not in values:
        selected = default_value or (values[0] if values else "")
        variable.set(selected)
    return selected
