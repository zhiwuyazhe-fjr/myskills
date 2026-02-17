---
name: docx
description: 将用户提供的 Markdown 文档或 md 风格纯文本转换为预设中文排版规范的 .docx 文件。用于需要严格控制标题层级样式、正文字体与行距、列表层级编号缩进、加粗与表格渲染的场景。
---

# DOCX 技能

## 工作流

1. 读取输入（文件路径或直接给出的 Markdown 文本）。
2. 按顺序解析结构：标题、列表、表格、正文段落。
3. 按固定规则渲染为 `.docx`（见 `scripts/docxskill/renderer.py`）。
4. 返回输出路径，并简要说明已应用的排版规则。

## 强制排版规则

- 一级标题：黑体、小三（16pt）、加粗、居中、黑色。
- 二级标题：黑体、四号（14pt）、加粗、左对齐、黑色。
- 三级标题：黑体、四号（14pt）、加粗、左对齐、黑色。
- 正文：中文宋体、英文 Times New Roman、小四（12pt）、黑色。
- 正文段落：固定 20 磅行距，首行缩进 2 个中文字符，段前段后均为 0。
- 列表规则：
  - 仅一级列表：每一项作为独立正文段落，首行缩进 2 个中文字符。
  - 两级及以上列表：除最后一级外，其余各级按层级连续编号；一级列表不首行缩进；二级及以后各级在前一级基础上再缩进 2 个中文字符。
- 加粗（`**text**` / `__text__`）必须正确渲染。
- Markdown 表格必须渲染为 Word 表格。
- 所有文字均为黑色。

## 调用入口

- 命令行：`python scripts/md2docx.py -i input.md -o output.docx`
- 模块方式：`python -m docxskill -i input.md -o output.docx`
- 核心函数：`scripts/docxskill/convert.py` 中的 `convert_markdown_to_docx()`

## 优先读取文件

- `scripts/docxskill/md_parser.py`
- `scripts/docxskill/renderer.py`
- `scripts/docxskill/list_formatting.py`
