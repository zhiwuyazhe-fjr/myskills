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

try:
    # 作为包执行（推荐）：python -m docx.scripts.generate_from_template
    from .list_formatting import ListItem, add_list_block, extract_heading_number
except ImportError:  # pragma: no cover
    # 作为脚本直接执行：python generate_from_template.py
    from list_formatting import ListItem, add_list_block, extract_heading_number


def create_document_from_template():
    """基于模板创建文档"""

    # 模板文件路径
    # 兼容两种常见命名：模板.docx / 模版.docx
    template_dir = Path(__file__).parent.parent
    template_path = template_dir / "模板.docx"
    if not template_path.exists():
        template_path = template_dir / "模版.docx"
    output_path = Path(__file__).parent.parent / "output.docx"

    # 复制模板文件
    if template_path.exists() and template_path.stat().st_size > 0:
        shutil.copy(template_path, output_path)
        doc = Document(output_path)
    else:
        print("警告：模板文件不存在或为空（0 字节），创建空白文档")
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

    # ==================== 列表示例（按需求优化） ====================
    # 场景①：只有第一级列表 —— 每项独立段落，首行缩进 2 字符（保持现状）
    add_body_text(doc, '仅一级列表示例：')
    add_list_block(
        doc,
        [
            ListItem("面向任务拆解，先列出关键步骤"),
            ListItem("给每一步补充输入/输出与验收标准"),
            ListItem("最后整体回顾并压缩表达"),
        ],
        current_heading_number=None,
    )

    doc.add_paragraph()  # 空行

    # 场景②：存在第二级列表 —— 一级自动编号、左对齐、取消首行缩进；二级每项独立段落首行缩进 2 字符
    heading_text = "4.2 研究方法"
    doc.add_heading(heading_text, level=2)
    current_num = extract_heading_number(heading_text)
    add_body_text(doc, '含二级列表示例：')
    add_list_block(
        doc,
        [
            ListItem(
                "数据准备",
                children=[
                    "收集并清洗原始数据",
                    "统一字段口径并补齐缺失值",
                ],
            ),
            ListItem(
                "模型训练",
                children=[
                    "划分训练/验证集并设定评价指标",
                    "记录参数与实验结果，便于复现",
                ],
            ),
        ],
        current_heading_number=current_num,
    )

    # 添加图注示例
    doc.add_paragraph()  # 空行
    add_figure_caption(doc, '图 2.1 技术架构图')

    # 保存文档
    doc.save(output_path)
    print(f"文档已生成: {output_path}")


if __name__ == "__main__":
    generate_sample_document()
