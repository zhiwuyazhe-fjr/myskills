#!/usr/bin/env python3
"""Convenience CLI entry point.

Example:
  py -3 .\\scripts\\md2docx.py -i input.md -o output.docx
"""

from docxskill.cli import main


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
