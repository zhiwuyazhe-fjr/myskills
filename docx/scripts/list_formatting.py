#!/usr/bin/env python3
"""
列表段落格式化工具（面向论文/报告场景）。

目标：在生成 docx 时，针对“只有一级列表”和“存在二级列表”的两种情况，
按约定控制编号、对齐和缩进。

规则（来自需求）：
1) 只有第一级列表：
   - 每一项作为独立段落
   - 首行缩进两个中文字符（约 Twips(40)）
2) 存在第二级列表：
   - 第一级列表：根据当前标题编号自动生成序号前缀（如 4.2.1 / 4.2.1.1），左对齐，取消首行缩进
   - 第二级列表：每一项作为独立段落，首行缩进两个中文字符（约 Twips(40)）
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Iterable, Optional

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Twips


_HEADING_NUM_RE = re.compile(r"^\s*(\d+(?:\.\d+)*)\b")


def extract_heading_number(text: str) -> Optional[str]:
    """
    从标题文本中提取形如 '4.2' / '4.2.1.1' 的编号前缀。
    找不到返回 None。
    """
    m = _HEADING_NUM_RE.match(text or "")
    return m.group(1) if m else None


@dataclass
class ListItem:
    """一级列表项，可包含二级列表项（children 仅支持一层）。"""

    text: str
    children: list[str] = field(default_factory=list)


def _two_chinese_chars_first_line_indent():
    # 经验值：2 个中文字符 ≈ Twips(40)（与仓库现有示例保持一致）
    return Twips(40)


def add_list_block(
    doc,
    items: Iterable[ListItem],
    *,
    current_heading_number: Optional[str] = None,
):
    """
    将列表块写入 doc（python-docx Document）。

    - items: ListItem 迭代器
    - current_heading_number: 当前标题编号（如 "4.2" 或 "4.2.1.1"）。
      若传 None，则在需要自动编号时会退化为使用 "1/2/3..." 作为前缀。
    """
    items = list(items)
    has_level2 = any(bool(it.children) for it in items)

    if not has_level2:
        # 场景①：仅一级列表（保持：每项独立段落 + 首行缩进 2 字符）
        for it in items:
            p = doc.add_paragraph(it.text)
            pf = p.paragraph_format
            pf.first_line_indent = _two_chinese_chars_first_line_indent()
        return

    # 场景②：存在二级列表
    for idx, it in enumerate(items, start=1):
        prefix_base = current_heading_number.strip() if current_heading_number else ""
        prefix = f"{prefix_base}.{idx}" if prefix_base else f"{idx}"
        p1 = doc.add_paragraph(f"{prefix} {it.text}")
        p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf1 = p1.paragraph_format
        # 取消首行缩进：显式设为 0，避免受样式影响
        pf1.first_line_indent = Twips(0)

        # 二级列表：每项独立段落 + 首行缩进 2 字符
        for child in it.children:
            p2 = doc.add_paragraph(child)
            pf2 = p2.paragraph_format
            pf2.first_line_indent = _two_chinese_chars_first_line_indent()

