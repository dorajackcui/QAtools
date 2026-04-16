#!/usr/bin/env python3
"""Backward-compatible wrapper for the term pair checker CLI."""

from tools.term_pair_checker.extract_terms_from_excel import *  # noqa: F401,F403
from tools.term_pair_checker.extract_terms_from_excel import main


if __name__ == "__main__":
    main()
