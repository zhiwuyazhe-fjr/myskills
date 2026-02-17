from __future__ import annotations

from typing import Sequence

from docx.shared import Pt

from .md_parser import InlineSpan, ListItem

_TWO_CHARS_PT = Pt(24)


def list_max_depth(items: Sequence[ListItem]) -> int:
    def depth(item: ListItem) -> int:
        if not item.children:
            return 1
        return 1 + max(depth(child) for child in item.children)

    if not items:
        return 1
    return max(depth(item) for item in items)


def iter_inline_text(inlines: Sequence[InlineSpan]) -> str:
    return "".join(span.text for span in inlines)


def indent_for_level(level: int) -> Pt:
    # Level 1 aligns to the left margin; each deeper level adds two Chinese characters.
    return Pt(24 * max(0, level - 1))


def two_chars_indent() -> Pt:
    return _TWO_CHARS_PT
