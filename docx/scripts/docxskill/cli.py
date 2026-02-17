from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Optional

from .convert import convert_markdown_to_docx
from .renderer import RenderOptions


def main(argv: Optional[list[str]] = None) -> int:
    p = argparse.ArgumentParser(
        prog="docxskill",
        description="Convert Markdown text into a .docx document with fixed Chinese academic formatting.",
    )
    p.add_argument("-i", "--input", help="Input .md file path. If omitted, read from stdin.")
    p.add_argument("-t", "--text", help="Markdown text input. Higher priority than --input/stdin.")
    p.add_argument("-o", "--output", required=True, help="Output .docx path.")
    p.add_argument("--template", help="Optional template .docx path.")
    p.add_argument(
        "--keep-template-body",
        action="store_true",
        help="When using template, keep original body content (default clears body but keeps section settings).",
    )

    args = p.parse_args(argv)

    md_text = _load_markdown_text(args.text, args.input)
    options = RenderOptions(clear_template_body=not args.keep_template_body)
    convert_markdown_to_docx(
        md_text,
        output_path=Path(args.output),
        template_path=Path(args.template) if args.template else None,
        options=options,
    )
    return 0


def _load_markdown_text(text_arg: Optional[str], input_path: Optional[str]) -> str:
    if text_arg:
        return text_arg
    if input_path:
        return Path(input_path).read_text(encoding="utf-8")
    return sys.stdin.read()


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
