from __future__ import annotations

from pathlib import Path
from typing import Optional

from .md_parser import parse_markdown
from .renderer import RenderOptions, open_document, render_blocks_to_docx


def convert_markdown_to_docx(
    markdown_text: str,
    *,
    output_path: str | Path,
    template_path: Optional[str | Path] = None,
    options: Optional[RenderOptions] = None,
) -> Path:
    options = options or RenderOptions()
    blocks = parse_markdown(markdown_text)
    doc = open_document(template_path=template_path, options=options)
    render_blocks_to_docx(doc, blocks, options=options)

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    return out
