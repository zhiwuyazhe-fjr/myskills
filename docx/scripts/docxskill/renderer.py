from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Sequence

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

from .list_formatting import indent_for_level, list_max_depth, two_chars_indent
from .md_parser import HeadingBlock, InlineSpan, ListBlock, ListItem, ParagraphBlock, TableBlock


@dataclass(frozen=True)
class RenderOptions:
    clear_template_body: bool = True


def open_document(*, template_path: Optional[str | Path] = None, options: RenderOptions) -> Document:
    if template_path:
        p = Path(template_path)
        if p.exists() and p.stat().st_size > 0:
            doc = Document(str(p))
            if options.clear_template_body:
                _clear_body_keep_sectpr(doc)
            return doc
    return Document()


def _is_block_empty(block) -> bool:
    """检查block是否为空"""
    if isinstance(block, HeadingBlock):
        return not block.inlines or all(not span.text for span in block.inlines)
    if isinstance(block, ParagraphBlock):
        return not block.inlines or all(not span.text for span in block.inlines)
    return False


def render_blocks_to_docx(doc: Document, blocks, *, options: RenderOptions) -> None:
    for b in blocks:
        if isinstance(b, HeadingBlock):
            if _is_block_empty(b):
                continue
            _render_heading(doc, b)
            continue

        if isinstance(b, ParagraphBlock):
            # 跳过空段落
            if _is_block_empty(b):
                continue
            p = doc.add_paragraph()
            _apply_body_paragraph_format(p, first_line_indent=two_chars_indent(), left_indent=Pt(0))
            _append_inlines(p, b.inlines, size_pt=12)
            continue

        if isinstance(b, ListBlock):
            _render_list_block(doc, b)
            continue

        if isinstance(b, TableBlock):
            _render_table_block(doc, b)
            continue


def _render_heading(doc: Document, block: HeadingBlock) -> None:
    level = max(1, min(3, int(block.level)))

    # 直接使用英文标题样式，避免使用不存在的样式导致创建额外空段落
    # Heading 1/2/3 是Word内置样式
    p = doc.add_paragraph(style=f"Heading {level}")

    # 设置对齐方式
    if level == 1:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # 清除默认样式中的加粗，使用黑体但不加粗
    for run in p.runs:
        run.font.bold = False
        run.font.name = "黑体"

    # 设置标题段前段后间距为0，去掉多余换行
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(20)

    # 添加标题文字（如果默认样式没有内容）
    if not p.text and block.inlines:
        _append_inlines(p, block.inlines, size_pt=16 if level == 1 else 14, force_bold=False, cn_font="黑体", en_font="黑体")


def _render_list_block(doc: Document, block: ListBlock) -> None:
    depth = list_max_depth(block.items)
    if depth <= 1:
        for item in block.items:
            p = doc.add_paragraph()
            _apply_body_paragraph_format(p, first_line_indent=two_chars_indent(), left_indent=Pt(0))
            _append_inlines(p, item.inlines, size_pt=12)
        return

    _render_nested_list_items(doc, block.items, level=1, path=[], max_depth=depth)


def _render_nested_list_items(
    doc: Document,
    items: Sequence[ListItem],
    *,
    level: int,
    path: list[int],
    max_depth: int,
) -> None:
    for idx, item in enumerate(items, start=1):
        current_path = [*path, idx]
        p = doc.add_paragraph()

        if level < max_depth:
            _apply_body_paragraph_format(p, first_line_indent=Pt(0), left_indent=indent_for_level(level))
            prefix = ".".join(str(n) for n in current_path)
            _append_inlines(p, [InlineSpan(text=f"{prefix} ", bold=False), *item.inlines], size_pt=12)
        else:
            _apply_body_paragraph_format(p, first_line_indent=Pt(0), left_indent=indent_for_level(level))
            _append_inlines(p, item.inlines, size_pt=12)

        if item.children:
            _render_nested_list_items(doc, item.children, level=level + 1, path=current_path, max_depth=max_depth)


def _render_table_block(doc: Document, block: TableBlock) -> None:
    col_count = len(block.header)
    table = doc.add_table(rows=1 + len(block.rows), cols=col_count)

    for c_idx, cell_spans in enumerate(block.header):
        para = table.cell(0, c_idx).paragraphs[0]
        _apply_body_paragraph_format(para, first_line_indent=Pt(0), left_indent=Pt(0))
        _append_inlines(para, cell_spans, size_pt=12, force_bold=True)

    for r_idx, row in enumerate(block.rows, start=1):
        for c_idx, cell_spans in enumerate(row):
            para = table.cell(r_idx, c_idx).paragraphs[0]
            _apply_body_paragraph_format(para, first_line_indent=Pt(0), left_indent=Pt(0))
            _append_inlines(para, cell_spans, size_pt=12)


def _apply_body_paragraph_format(paragraph, *, first_line_indent: Pt, left_indent: Pt) -> None:
    pf = paragraph.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(20)
    pf.first_line_indent = first_line_indent
    pf.left_indent = left_indent
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT


def _append_inlines(
    paragraph,
    inlines: Sequence[InlineSpan],
    *,
    size_pt: int,
    force_bold: bool = False,
    cn_font: str = "宋体",
    en_font: str = "Times New Roman",
) -> None:
    if not inlines:
        run = paragraph.add_run("")
        _style_run(run, size_pt=size_pt, bold=force_bold, cn_font=cn_font, en_font=en_font)
        return

    for span in inlines:
        run = paragraph.add_run(span.text)
        _style_run(run, size_pt=size_pt, bold=(force_bold or span.bold), cn_font=cn_font, en_font=en_font)


def _style_run(run, *, size_pt: int, bold: bool, cn_font: str, en_font: str) -> None:
    run.bold = bool(bold)
    run.font.size = Pt(size_pt)
    run.font.name = en_font
    run.font.color.rgb = RGBColor(0, 0, 0)

    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn("w:ascii"), en_font)
    rFonts.set(qn("w:hAnsi"), en_font)
    rFonts.set(qn("w:eastAsia"), cn_font)


def _clear_body_keep_sectpr(doc: Document) -> None:
    body = doc._body._element  # type: ignore[attr-defined]
    for child in list(body):
        if child.tag.endswith("}sectPr"):
            continue
        body.remove(child)
