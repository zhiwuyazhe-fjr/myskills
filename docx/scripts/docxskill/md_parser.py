from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Sequence


@dataclass(frozen=True)
class InlineSpan:
    text: str
    bold: bool = False


@dataclass(frozen=True)
class HeadingBlock:
    level: int
    inlines: list[InlineSpan]


@dataclass(frozen=True)
class ParagraphBlock:
    inlines: list[InlineSpan]


@dataclass
class ListItem:
    inlines: list[InlineSpan]
    children: list["ListItem"] = field(default_factory=list)


@dataclass(frozen=True)
class ListBlock:
    items: list[ListItem]


@dataclass(frozen=True)
class TableBlock:
    header: list[list[InlineSpan]]
    rows: list[list[list[InlineSpan]]]


Block = HeadingBlock | ParagraphBlock | ListBlock | TableBlock


_HEADING_RE = re.compile(r"^(#{1,6})\s+(.+?)\s*$")
_LIST_ITEM_RE = re.compile(r"^(?P<indent>[ \t]*)(?P<marker>(?:[-*+])|(?:\d+[.)]))\s+(?P<text>.+?)\s*$")
_BOLD_RE = re.compile(r"(\*\*.+?\*\*|__.+?__)")


def parse_markdown(markdown_text: str) -> list[Block]:
    lines = _normalize_lines(markdown_text)
    blocks: list[Block] = []

    para_buf: list[str] = []
    i = 0
    while i < len(lines):
        line = lines[i]

        if not line.strip():
            _flush_paragraph(blocks, para_buf)
            i += 1
            continue

        m_h = _HEADING_RE.match(line)
        if m_h:
            _flush_paragraph(blocks, para_buf)
            blocks.append(
                HeadingBlock(level=len(m_h.group(1)), inlines=parse_inlines(m_h.group(2).strip()))
            )
            i += 1
            continue

        if _LIST_ITEM_RE.match(line):
            _flush_paragraph(blocks, para_buf)
            list_block, i = _parse_list_block(lines, start=i)
            blocks.append(list_block)
            continue

        if _is_table_start(lines, i):
            _flush_paragraph(blocks, para_buf)
            table_block, i = _parse_table_block(lines, start=i)
            blocks.append(table_block)
            continue

        para_buf.append(line.strip())
        i += 1

    _flush_paragraph(blocks, para_buf)
    return blocks


def parse_inlines(text: str) -> list[InlineSpan]:
    s = (text or "").strip()
    if not s:
        return []

    spans: list[InlineSpan] = []
    pos = 0
    for m in _BOLD_RE.finditer(s):
        if m.start() > pos:
            spans.append(InlineSpan(text=s[pos : m.start()], bold=False))
        token = m.group(0)
        content = token[2:-2]
        if content:
            spans.append(InlineSpan(text=content, bold=True))
        pos = m.end()

    if pos < len(s):
        spans.append(InlineSpan(text=s[pos:], bold=False))

    return [span for span in spans if span.text]


def _normalize_lines(text: str) -> list[str]:
    return (text or "").replace("\r\n", "\n").replace("\r", "\n").split("\n")


def _flush_paragraph(blocks: list[Block], buf: list[str]) -> None:
    text = " ".join(s for s in (x.strip() for x in buf) if s)
    if text:
        blocks.append(ParagraphBlock(inlines=parse_inlines(text)))
    buf.clear()


def _indent_width(indent: str) -> int:
    return sum(4 if ch == "\t" else 1 for ch in indent)


def _parse_list_block(lines: Sequence[str], *, start: int) -> tuple[ListBlock, int]:
    root_items: list[ListItem] = []
    stack: list[tuple[int, list[ListItem]]] = [(-1, root_items)]
    last_item: ListItem | None = None
    i = start

    while i < len(lines):
        raw = lines[i]

        if not raw.strip():
            i += 1
            if i >= len(lines):
                break
            if _LIST_ITEM_RE.match(lines[i]) or _looks_like_continuation(lines[i]):
                continue
            break

        m = _LIST_ITEM_RE.match(raw)
        if m:
            indent = _indent_width(m.group("indent"))
            item = ListItem(inlines=parse_inlines(m.group("text").strip()))

            while len(stack) > 1 and indent <= stack[-1][0]:
                stack.pop()

            stack[-1][1].append(item)
            stack.append((indent, item.children))
            last_item = item
            i += 1
            continue

        if _looks_like_continuation(raw) and last_item is not None:
            cont = raw.strip()
            if cont:
                if last_item.inlines:
                    last_item.inlines.append(InlineSpan(text=" " + cont, bold=False))
                else:
                    last_item.inlines = [InlineSpan(text=cont, bold=False)]
            i += 1
            continue

        break

    return ListBlock(items=root_items), i


def _is_table_start(lines: Sequence[str], i: int) -> bool:
    if i + 1 >= len(lines):
        return False
    return _looks_like_table_row(lines[i]) and _is_table_separator(lines[i + 1])


def _looks_like_table_row(line: str) -> bool:
    s = line.strip()
    return "|" in s and len([c for c in _split_table_row(s) if c.strip()]) >= 1


def _is_table_separator(line: str) -> bool:
    cells = _split_table_row(line.strip())
    if not cells:
        return False
    for cell in cells:
        token = cell.strip()
        if not token:
            continue
        if not re.fullmatch(r":?-{3,}:?", token):
            return False
    return True


def _split_table_row(line: str) -> list[str]:
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    return [part.strip() for part in s.split("|")]


def _parse_table_block(lines: Sequence[str], *, start: int) -> tuple[TableBlock, int]:
    header_cells = _split_table_row(lines[start])
    header = [parse_inlines(cell) for cell in header_cells]

    rows: list[list[list[InlineSpan]]] = []
    i = start + 2
    col_count = len(header)

    while i < len(lines):
        raw = lines[i]
        if not raw.strip():
            break
        if not _looks_like_table_row(raw) or _LIST_ITEM_RE.match(raw) or _HEADING_RE.match(raw):
            break

        row_cells = _split_table_row(raw)
        if len(row_cells) < col_count:
            row_cells += [""] * (col_count - len(row_cells))
        elif len(row_cells) > col_count:
            row_cells = row_cells[:col_count]

        rows.append([parse_inlines(cell) for cell in row_cells])
        i += 1

    return TableBlock(header=header, rows=rows), i


def _looks_like_continuation(line: str) -> bool:
    if not line:
        return False
    if _LIST_ITEM_RE.match(line):
        return False
    return bool(line[:1].isspace())
