---
name: docx
description: "创建、编辑 Word 文档。支持基于模板生成精美文档、读取文档、修改文档"
---

# DOCX 技能

## 概述

本技能用于创建、编辑 .docx 文件。

## 快速开始

### 1. 基于模板生成文档（推荐）

**最简单的方式**：复制 `模板.docx`，然后使用 Python 添加内容：

```python
from docx import Document

# 打开模板
doc = Document('模板.docx')

# 添加标题
doc.add_heading('第一章 绪论', level=1)
doc.add_heading('1.1 研究背景', level=2)
doc.add_heading('1.1.1 国内研究现状', level=3)

# 添加正文（自动首行缩进）
para = doc.add_paragraph('正文内容')

# 保存
doc.save('输出.docx')
```

### 2. 使用 pandoc 转换

```bash
# 从 Markdown/HTML/txt 生成
pandoc 输入文件.md -o 输出.docx
pandoc 输入文件.txt -o 输出.docx

# 读取文档内容
pandoc 文档.docx -o 输出.md
```

## 模板格式说明

### 样式结构

| 样式 | 用途 | 编号格式 |
|------|------|----------|
| heading 1 (styleId=1) | 一级标题（章） | 第X章 |
| heading 2 (styleId=2) | 二级标题（节） | X.Y |
| heading 3 (styleId=3) | 三级标题（小节） | X.Y.Z |
| 正文1 (styleId=10) | 正文段落 | 首行缩进2字符 |

### 字体格式

- **一级标题**: 黑体，32pt，居中对齐
- **二级标题**: 黑体，28pt，左对齐
- **三级标题**: 黑体，28pt，左对齐
- **正文**: Times New Roman/宋体，24pt，首行缩进

## 详细用法

### Python（推荐）

安装依赖：
```bash
pip install python-docx
```

完整示例：
```python
from docx import Document
from docx.shared import Pt, Twips

doc = Document('模板.docx')

# 一级标题
doc.add_heading('第一章 绪论', level=1)

# 二级标题
doc.add_heading('1.1 研究背景', level=2)

# 三级标题
doc.add_heading('1.1.1 国内研究现状', level=3)

# 正文（首行缩进）
para = doc.add_paragraph('正文内容...')
para.paragraph_format.first_line_indent = Twips(40)

doc.save('输出.docx')
```

### JavaScript (docx-js)

安装依赖：
```bash
npm install docx
```

完整示例：
```javascript
const { Document, Packer, Paragraph, HeadingLevel, AlignmentType, TextRun } = require("docx");

const doc = new Document({
    sections: [{
        children: [
            new Paragraph({
                text: "第一章 绪论",
                heading: HeadingLevel.HEADING_1,
                alignment: AlignmentType.CENTER,
            }),
            new Paragraph({
                children: [new TextRun("正文内容...")],
                indentation: { firstLineChars: 200 },
            }),
        ],
    }],
});

Packer.toBuffer(doc).then(buffer => {
    require('fs').writeFileSync('输出.docx', buffer);
});
```

### 修改现有文档

解压文档：
```bash
unzip 文档.docx -d unpacked
```

编辑 `word/document.xml`，然后打包：
```bash
zip -r 输出.docx unpacked/*
```

## 关键文件

- `模板.docx` - 模板文件
- `scripts/generate_from_template.py` - Python 示例脚本
- `scripts/generate_docx.js` - JavaScript 示例脚本
- `docx-js.md` - docx-js 库详细用法
- `ooxml.md` - 底层 XML 操作参考

## 依赖

- `pandoc` - 格式转换
- `docx` (npm) - JavaScript 创建文档
- `python-docx` (pip) - Python 操作文档
