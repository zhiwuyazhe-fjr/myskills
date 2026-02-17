#!/usr/bin/env python3
"""
基于模板生成格式精美的Word文档

使用说明：
    python generate_from_template.py

此脚本演示如何基于 模板.docx 生成格式规范的论文文档。
"""

import shutil
import os
from pathlib import Path
from docx import Document
from docx.shared import Pt, Cm, Twips, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn


def create_document_from_template():
    """基于模板创建文档"""

    # 模板文件路径
    template_path = Path(__file__).parent.parent / "模板.docx"
    output_path = Path(__file__).parent.parent / "output.docx"

    # 复制模板文件
    if template_path.exists():
        shutil.copy(template_path, output_path)
        doc = Document(output_path)
    else:
        print("警告：模板文件不存在，创建空白文档")
        doc = Document()

    # 清除现有内容（保留样式）
    # 注意：根据需要决定是否清除
    # for paragraph in doc.paragraphs[:]:
    #     p = paragraph._element
    #     p.getparent().remove(p)

    return doc, output_path


def add_heading_with_style(doc, text, level=1):
    """添加标题（保留模板样式）"""
    heading = doc.add_heading(text, level=level)
    return heading


def add_body_text(doc, text):
    """添加正文（首行缩进）"""
    para = doc.add_paragraph(text)

    # 设置首行缩进（2字符 = 约 Twips(20) * 2 = 40）
    para_format = para.paragraph_format
    para_format.first_line_indent = Twips(40)  # 约0.5寸

    return para


def add_figure_caption(doc, text):
    """添加图注"""
    para = doc.add_paragraph(text)
    para_format = para.paragraph_format
    para_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 设置样式（如果模板中有 caption 样式）
    for run in para.runs:
        run.font.name = 'Cambria'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
        run.font.size = Pt(10.5)  # 21/2 = 10.5pt

    return para


def generate_sample_document():
    """生成示例文档"""

    doc, output_path = create_document_from_template()

    # 添加一级标题
    doc.add_heading('第一章 绪论', level=1)

    # 添加正文
    add_body_text(doc, '本章首先介绍研究背景及意义，然后阐述国内外研究现状，最后说明本文的主要研究内容和方法。')

    # 添加二级标题
    doc.add_heading('1.1 研究背景', level=2)

    # 添加正文
    add_body_text(doc, '随着信息技术的快速发展，大数据、人工智能等技术在各行各业的应用越来越广泛。')

    # 添加三级标题
    doc.add_heading('1.1.1 国内研究现状', level=3)

    # 添加正文
    add_body_text(doc, '近年来，国内学者在大数据处理领域取得了丰硕的研究成果。')

    # 添加另一个三级标题
    doc.add_heading('1.1.2 国外研究现状', level=3)

    # 添加正文
    add_body_text(doc, '国外发达国家在相关领域的研究起步较早，形成了一套完整的技术体系。')

    # 添加二级标题
    doc.add_heading('1.2 研究意义', level=2)

    # 添加正文
    add_body_text(doc, '本研究对于推动技术发展和实际应用具有重要的理论价值和实践意义。')

    # 添加一级标题
    doc.add_heading('第二章 相关技术与理论', level=1)

    # 添加正文
    add_body_text(doc, '本章主要介绍本文研究所涉及的关键技术和理论基础。')

    # 添加二级标题
    doc.add_heading('2.1 关键技术', level=2)

    # 添加三级标题
    doc.add_heading('2.1.1 技术概述', level=3)

    add_body_text(doc, '关键技术包括以下几个方面：')

    # 添加三级标题
    doc.add_heading('2.1.2 技术实现', level=3)

    add_body_text(doc, '技术实现过程中需要注意以下问题：')

    # 添加图注示例
    doc.add_paragraph()  # 空行
    add_figure_caption(doc, '图 2.1 技术架构图')

    # 保存文档
    doc.save(output_path)
    print(f"文档已生成: {output_path}")


if __name__ == "__main__":
    generate_sample_document()
