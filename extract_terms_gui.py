#!/usr/bin/env python3
"""Backward-compatible wrapper for the term pair checker GUI."""

import sys

from tools.term_pair_checker.extract_terms_gui import main


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
