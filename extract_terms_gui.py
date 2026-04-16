#!/usr/bin/env python3
"""Backward-compatible wrapper for the term pair checker GUI."""

from tools.term_pair_checker.extract_terms_gui import *  # noqa: F401,F403
from tools.term_pair_checker.extract_terms_gui import main


if __name__ == "__main__":
    main()
