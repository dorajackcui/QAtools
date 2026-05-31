#!/usr/bin/env python3
"""Shared local Codex subprocess runner."""

from __future__ import annotations

import subprocess
from pathlib import Path


def build_codex_exec_command(
    *,
    output_path: str | Path,
    model: str | None = None,
    reasoning_effort: str = "high",
    sandbox: str = "read-only",
    codex_command: str = "codex",
) -> list[str]:
    command = [
        codex_command,
        "--ask-for-approval",
        "never",
        "--config",
        f'model_reasoning_effort="{reasoning_effort}"',
        "exec",
        "--ephemeral",
        "--ignore-rules",
        "--skip-git-repo-check",
        "--sandbox",
        sandbox,
        "--output-last-message",
        str(output_path),
    ]
    if model:
        command.extend(["--model", model])
    command.append("-")
    return command


def run_codex_exec_prompt(
    prompt: str,
    *,
    output_path: str | Path,
    model: str | None = None,
    reasoning_effort: str = "high",
    timeout_seconds: int = 600,
    sandbox: str = "read-only",
    codex_command: str = "codex",
    error_prefix: str = "Codex execution failed",
) -> str:
    command = build_codex_exec_command(
        output_path=output_path,
        model=model,
        reasoning_effort=reasoning_effort,
        sandbox=sandbox,
        codex_command=codex_command,
    )
    completed = subprocess.run(
        command,
        input=prompt,
        capture_output=True,
        text=True,
        timeout=timeout_seconds,
        check=False,
    )
    if completed.returncode != 0:
        error_text = completed.stderr.strip() or completed.stdout.strip()
        raise RuntimeError(f"{error_prefix}: {error_text}")

    output_file = Path(output_path)
    return output_file.read_text(encoding="utf-8") if output_file.exists() else completed.stdout
