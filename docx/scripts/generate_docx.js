/**
 * 基于模板生成格式精美的Word文档
 *
 * 使用方法：
 * 1. 安装依赖：npm install docx file-saver
 * 2. 运行：node generate_docx.js
 */

const { Document, Packer, Paragraph, TextRun, HeadingLevel, WidthType, AlignmentType } = require("docx");
const fs = require("fs");

// 创建文档
const doc = new Document({
    sections: [{
        properties: {},
        children: [
            // 一级标题（第一章）
            new Paragraph({
                text: "第一章 绪论",
                heading: HeadingLevel.HEADING_1,
                alignment: AlignmentType.CENTER,
                spacing: {
                    before: 480,
                    after: 360,
                },
            }),

            // 正文
            new Paragraph({
                children: [
                    new TextRun({
                        text: "本章首先介绍研究背景及意义，然后阐述国内外研究现状，最后说明本文的主要研究内容和方法。",
                        font: "Times New Roman",
                        size: 24,
                    }),
                ],
                indentation: {
                    firstLineChars: 200,
                },
                spacing: {
                    line: 300,
                },
            }),

            // 二级标题（1.1）
            new Paragraph({
                text: "1.1 研究背景",
                heading: HeadingLevel.HEADING_2,
                spacing: {
                    before: 240,
                    after: 240,
                },
            }),

            // 正文
            new Paragraph({
                children: [
                    new TextRun({
                        text: "随着信息技术的快速发展，大数据、人工智能等技术在各行各业的应用越来越广泛。",
                        font: "Times New Roman",
                        size: 24,
                    }),
                ],
                indentation: {
                    firstLineChars: 200,
                },
                spacing: {
                    line: 300,
                },
            }),

            // 三级标题（1.1.1）
            new Paragraph({
                text: "1.1.1 国内研究现状",
                heading: HeadingLevel.HEADING_3,
                spacing: {
                    before: 240,
                    after: 240,
                },
            }),

            // 正文
            new Paragraph({
                children: [
                    new TextRun({
                        text: "近年来，国内学者在大数据处理领域取得了丰硕的研究成果。",
                        font: "Times New Roman",
                        size: 24,
                    }),
                ],
                indentation: {
                    firstLineChars: 200,
                },
                spacing: {
                    line: 300,
                },
            }),

            // 另一个三级标题
            new Paragraph({
                text: "1.1.2 国外研究现状",
                heading: HeadingLevel.HEADING_3,
                spacing: {
                    before: 240,
                    after: 240,
                },
            }),

            // 正文
            new Paragraph({
                children: [
                    new TextRun({
                        text: "国外发达国家在相关领域的研究起步较早，形成了一套完整的技术体系。",
                        font: "Times New Roman",
                        size: 24,
                    }),
                ],
                indentation: {
                    firstLineChars: 200,
                },
                spacing: {
                    line: 300,
                },
            }),

            // 二级标题
            new Paragraph({
                text: "1.2 研究意义",
                heading: HeadingLevel.HEADING_2,
                spacing: {
                    before: 240,
                    after: 240,
                },
            }),

            // 正文
            new Paragraph({
                children: [
                    new TextRun({
                        text: "本研究对于推动技术发展和实际应用具有重要的理论价值和实践意义。",
                        font: "Times New Roman",
                        size: 24,
                    }),
                ],
                indentation: {
                    firstLineChars: 200,
                },
                spacing: {
                    line: 300,
                },
            }),

            // 第一章结束

            // 第二章
            new Paragraph({
                text: "第二章 相关技术与理论",
                heading: HeadingLevel.HEADING_1,
                alignment: AlignmentType.CENTER,
                spacing: {
                    before: 480,
                    after: 360,
                },
            }),

            new Paragraph({
                children: [
                    new TextRun({
                        text: "本章主要介绍本文研究所涉及的关键技术和理论基础。",
                        font: "Times New Roman",
                        size: 24,
                    }),
                ],
                indentation: {
                    firstLineChars: 200,
                },
                spacing: {
                    line: 300,
                },
            }),

            // 二级标题
            new Paragraph({
                text: "2.1 关键技术",
                heading: HeadingLevel.HEADING_2,
                spacing: {
                    before: 240,
                    after: 240,
                },
            }),

            // 三级标题
            new Paragraph({
                text: "2.1.1 技术概述",
                heading: HeadingLevel.HEADING_3,
                spacing: {
                    before: 240,
                    after: 240,
                },
            }),

            new Paragraph({
                children: [
                    new TextRun({
                        text: "关键技术包括以下几个方面：",
                        font: "Times New Roman",
                        size: 24,
                    }),
                ],
                indentation: {
                    firstLineChars: 200,
                },
                spacing: {
                    line: 300,
                },
            }),

            // 三级标题
            new Paragraph({
                text: "2.1.2 技术实现",
                heading: HeadingLevel.HEADING_3,
                spacing: {
                    before: 240,
                    after: 240,
                },
            }),

            new Paragraph({
                children: [
                    new TextRun({
                        text: "技术实现过程中需要注意以下问题：",
                        font: "Times New Roman",
                        size: 24,
                    }),
                ],
                indentation: {
                    firstLineChars: 200,
                },
                spacing: {
                    line: 300,
                },
            }),

            // 图注示例
            new Paragraph({
                children: [
                    new TextRun({
                        text: "图 2.1 技术架构图",
                        font: "Cambria",
                        size: 20,
                    }),
                ],
                alignment: AlignmentType.CENTER,
            }),
        ],
    }],
});

// 保存文档
Packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync("output.docx", buffer);
    console.log("文档已生成: output.docx");
});
