#!/usr/bin/env python3
"""
处理 Word 文档评论和编辑的库。

用法：
    from skills.docx.scripts.document import Document

    # 初始化
    doc = Document('workspace/unpacked')
    doc = Document('workspace/unpacked', author="John Doe", initials="JD")

    # 查找节点
    node = doc["word/document.xml"].get_node(tag="w:p", line_number=10)

    # 添加评论
    doc.add_comment(start=node, end=node, text="Comment text")
    doc.reply_to_comment(parent_comment_id=0, text="Reply text")

    # 保存
    doc.save()
"""

import html
import random
import shutil
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from defusedxml import minidom
from ooxml.scripts.pack import pack_document
from ooxml.scripts.validation.docx import DOCXSchemaValidator

from .utilities import XMLEditor

# 模板文件路径
TEMPLATE_DIR = Path(__file__).parent / "templates"


class DocxXMLEditor(XMLEditor):
    """自动为新元素应用 RSID、作者和日期的 XMLEditor。

    插入新内容时自动向支持的元素添加属性：
    - w:rsidR、w:rsidRDefault、w:rsidP（用于 w:p 和 w:r 元素）
    - w:author 和 w:date（用于 w:comment 元素）
    - w:id（用于 w:comment 元素）

    属性：
        dom (defusedxml.minidom.Document): 用于直接操作的 DOM 文档
    """

    def __init__(
        self, xml_path, rsid: str, author: str = "Claude", initials: str = "C"
    ):
        """使用必需的 RSID 和可选的作者进行初始化。

        参数：
            xml_path: 要编辑的 XML 文件路径
            rsid: 自动应用于新元素的 RSID
            author: 评论的作者名称（默认："Claude"）
            initials: 作者姓名首字母（默认："C"）
        """
        super().__init__(xml_path)
        self.rsid = rsid
        self.author = author
        self.initials = initials

    def _ensure_w14_namespace(self):
        """确保根元素上声明了 w14 命名空间。"""
        root = self.dom.documentElement
        if not root.hasAttribute("xmlns:w14"):  # type: ignore
            root.setAttribute(  # type: ignore
                "xmlns:w14",
                "http://schemas.microsoft.com/office/word/2010/wordml",
            )

    def _inject_attributes_to_nodes(self, nodes):
        """向适用的 DOM 节点注入 RSID、作者和日期属性。

        向支持的元素添加属性：
        - w:r：获取 w:rsidR
        - w:p：获取 w:rsidR、w:rsidRDefault、w:rsidP、w14:paraId、w14:textId
        - w:t：如果文本有前导/尾随空格则获取 xml:space="preserve"
        - w:comment：获取 w:author、w:date、w:initials

        参数：
            nodes: 要处理的 DOM 节点列表
        """
        timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

        def add_rsid_to_p(elem):
            if not elem.hasAttribute("w:rsidR"):
                elem.setAttribute("w:rsidR", self.rsid)
            if not elem.hasAttribute("w:rsidRDefault"):
                elem.setAttribute("w:rsidRDefault", self.rsid)
            if not elem.hasAttribute("w:rsidP"):
                elem.setAttribute("w:rsidP", self.rsid)
            # 如果不存在则添加 w14:paraId 和 w14:textId
            if not elem.hasAttribute("w14:paraId"):
                self._ensure_w14_namespace()
                elem.setAttribute("w14:paraId", _generate_hex_id())
            if not elem.hasAttribute("w14:textId"):
                self._ensure_w14_namespace()
                elem.setAttribute("w14:textId", _generate_hex_id())

        def add_rsid_to_r(elem):
            if not elem.hasAttribute("w:rsidR"):
                elem.setAttribute("w:rsidR", self.rsid)

        def add_comment_attrs(elem):
            if not elem.hasAttribute("w:author"):
                elem.setAttribute("w:author", self.author)
            if not elem.hasAttribute("w:date"):
                elem.setAttribute("w:date", timestamp)
            if not elem.hasAttribute("w:initials"):
                elem.setAttribute("w:initials", self.initials)

        def add_xml_space_to_t(elem):
            # 如果 w:t 的文本有前导/尾随空格则添加 xml:space="preserve"
            if (
                elem.firstChild
                and elem.firstChild.nodeType == elem.firstChild.TEXT_NODE
            ):
                text = elem.firstChild.data
                if text and (text[0].isspace() or text[-1].isspace()):
                    if not elem.hasAttribute("xml:space"):
                        elem.setAttribute("xml:space", "preserve")

        for node in nodes:
            if node.nodeType != node.ELEMENT_NODE:
                continue

            # 处理节点本身
            if node.tagName == "w:p":
                add_rsid_to_p(node)
            elif node.tagName == "w:r":
                add_rsid_to_r(node)
            elif node.tagName == "w:t":
                add_xml_space_to_t(node)
            elif node.tagName == "w:comment":
                add_comment_attrs(node)

            # 处理子节点
            for elem in node.getElementsByTagName("w:p"):
                add_rsid_to_p(elem)
            for elem in node.getElementsByTagName("w:r"):
                add_rsid_to_r(elem)
            for elem in node.getElementsByTagName("w:t"):
                add_xml_space_to_t(elem)
            for elem in node.getElementsByTagName("w:comment"):
                add_comment_attrs(elem)

    def replace_node(self, elem, new_content):
        """替换节点并自动注入属性。"""
        nodes = super().replace_node(elem, new_content)
        self._inject_attributes_to_nodes(nodes)
        return nodes

    def insert_after(self, elem, xml_content):
        """在之后插入并自动注入属性。"""
        nodes = super().insert_after(elem, xml_content)
        self._inject_attributes_to_nodes(nodes)
        return nodes

    def insert_before(self, elem, xml_content):
        """在之前插入并自动注入属性。"""
        nodes = super().insert_before(elem, xml_content)
        self._inject_attributes_to_nodes(nodes)
        return nodes

    def append_to(self, elem, xml_content):
        """追加并自动注入属性。"""
        nodes = super().append_to(elem, xml_content)
        self._inject_attributes_to_nodes(nodes)
        return nodes


def _generate_hex_id() -> str:
    """生成随机的 8 位十六进制 ID，用于 para/durable ID。"""
    return f"{random.randint(1, 0x7FFFFFFE):08X}"


def _generate_rsid() -> str:
    """生成随机的 8 位十六进制 RSID。"""
    return "".join(random.choices("0123456789ABCDEF", k=8))


class Document:
    """管理解压缩 Word 文档中的评论。"""

    def __init__(
        self,
        unpacked_dir,
        rsid=None,
        author="Claude",
        initials="C",
    ):
        """
        使用解压缩的 Word 文档目录路径进行初始化。
        自动设置评论基础设施（people.xml、RSID）。

        参数：
            unpacked_dir: 解压缩的 DOCX 目录路径（必须包含 word/ 子目录）
            rsid: 用于所有元素的可选 RSID。如果未提供，将生成一个。
            author: 评论的默认作者名称（默认："Claude"）
            initials: 评论的默认作者姓名首字母（默认："C"）
        """
        self.original_path = Path(unpacked_dir)

        if not self.original_path.exists() or not self.original_path.is_dir():
            raise ValueError(f"Directory not found: {unpacked_dir}")

        # 创建包含解压缩内容和基准子目录的临时目录
        self.temp_dir = tempfile.mkdtemp(prefix="docx_")
        self.unpacked_path = Path(self.temp_dir) / "unpacked"
        shutil.copytree(self.original_path, self.unpacked_path)

        # 将原始目录打包为临时 .docx 用于验证基准
        self.original_docx = Path(self.temp_dir) / "original.docx"
        pack_document(self.original_path, self.original_docx, validate=False)

        self.word_path = self.unpacked_path / "word"

        # 如果未提供则生成 RSID
        self.rsid = rsid if rsid else _generate_rsid()
        print(f"Using RSID: {self.rsid}")

        # 设置默认作者和姓名首字母
        self.author = author
        self.initials = initials

        # 延迟加载编辑器的缓存
        self._editors = {}

        # 评论文件路径
        self.comments_path = self.word_path / "comments.xml"
        self.comments_extended_path = self.word_path / "commentsExtended.xml"
        self.comments_ids_path = self.word_path / "commentsIds.xml"
        self.comments_extensible_path = self.word_path / "commentsExtensible.xml"

        # 加载现有评论并确定下一个 ID
        self.existing_comments = self._load_existing_comments()
        self.next_comment_id = self._get_next_comment_id()

        # 方便访问 document.xml 编辑器
        self._document = self["word/document.xml"]

        # 设置评论基础设施
        self._setup_comments()

        # 将作者添加到 people.xml
        self._add_author_to_people(author)

    def __getitem__(self, xml_path: str) -> DocxXMLEditor:
        """获取或创建指定 XML 文件的 DocxXMLEditor。"""
        if xml_path not in self._editors:
            file_path = self.unpacked_path / xml_path
            if not file_path.exists():
                raise ValueError(f"XML file not found: {xml_path}")
            self._editors[xml_path] = DocxXMLEditor(
                file_path, rsid=self.rsid, author=self.author, initials=self.initials
            )
        return self._editors[xml_path]

    def add_comment(self, start, end, text: str) -> int:
        """添加从一元素跨越到另一元素的评论。"""
        comment_id = self.next_comment_id
        para_id = _generate_hex_id()
        durable_id = _generate_hex_id()
        timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

        # 立即将评论范围添加到 document.xml
        self._document.insert_before(start, self._comment_range_start_xml(comment_id))

        # 如果结束节点是段落，则在其中追加评论标记
        if end.tagName == "w:p":
            self._document.append_to(end, self._comment_range_end_xml(comment_id))
        else:
            self._document.insert_after(end, self._comment_range_end_xml(comment_id))

        # 立即添加到 comments.xml
        self._add_to_comments_xml(
            comment_id, para_id, text, self.author, self.initials, timestamp
        )

        # 立即添加到 commentsExtended.xml
        self._add_to_comments_extended_xml(para_id, parent_para_id=None)

        # 立即添加到 commentsIds.xml
        self._add_to_comments_ids_xml(para_id, durable_id)

        # 立即添加到 commentsExtensible.xml
        self._add_to_comments_extensible_xml(durable_id)

        # 更新 existing_comments 以便回复功能正常
        self.existing_comments[comment_id] = {"para_id": para_id}

        self.next_comment_id += 1
        return comment_id

    def reply_to_comment(
        self,
        parent_comment_id: int,
        text: str,
    ) -> int:
        """添加对现有评论的回复。"""
        if parent_comment_id not in self.existing_comments:
            raise ValueError(f"Parent comment with id={parent_comment_id} not found")

        parent_info = self.existing_comments[parent_comment_id]
        comment_id = self.next_comment_id
        para_id = _generate_hex_id()
        durable_id = _generate_hex_id()
        timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

        # 立即将评论范围添加到 document.xml
        parent_start_elem = self._document.get_node(
            tag="w:commentRangeStart", attrs={"w:id": str(parent_comment_id)}
        )
        parent_ref_elem = self._document.get_node(
            tag="w:commentReference", attrs={"w:id": str(parent_comment_id)}
        )

        self._document.insert_after(
            parent_start_elem, self._comment_range_start_xml(comment_id)
        )
        parent_ref_run = parent_ref_elem.parentNode
        self._document.insert_after(
            parent_ref_run, f'<w:commentRangeEnd w:id="{comment_id}"/>'
        )
        self._document.insert_after(
            parent_ref_run, self._comment_ref_run_xml(comment_id)
        )

        # 立即添加到 comments.xml
        self._add_to_comments_xml(
            comment_id, para_id, text, self.author, self.initials, timestamp
        )

        # 立即添加到 commentsExtended.xml（带父级）
        self._add_to_comments_extended_xml(
            para_id, parent_para_id=parent_info["para_id"]
        )

        # 立即添加到 commentsIds.xml
        self._add_to_comments_ids_xml(para_id, durable_id)

        # 立即添加到 commentsExtensible.xml
        self._add_to_comments_extensible_xml(durable_id)

        # 更新 existing_comments 以便回复功能正常
        self.existing_comments[comment_id] = {"para_id": para_id}

        self.next_comment_id += 1
        return comment_id

    def __del__(self):
        """删除时清理临时目录。"""
        if hasattr(self, "temp_dir") and Path(self.temp_dir).exists():
            shutil.rmtree(self.temp_dir)

    def validate(self) -> None:
        """根据 XSD 模式验证文档。"""
        schema_validator = DOCXSchemaValidator(
            self.unpacked_path, self.original_docx, verbose=False
        )

        if not schema_validator.validate():
            raise ValueError("Schema validation failed")

    def save(self, destination=None, validate=True) -> None:
        """将所有修改的 XML 文件保存到磁盘。"""
        # 仅在评论文件存在时确保评论关系和内容类型
        if self.comments_path.exists():
            self._ensure_comment_relationships()
            self._ensure_comment_content_types()

        # 保存临时目录中的所有修改的 XML 文件
        for editor in self._editors.values():
            editor.save()

        # 默认进行验证
        if validate:
            self.validate()

        # 将内容从临时目录复制到目标位置
        target_path = Path(destination) if destination else self.original_path
        shutil.copytree(self.unpacked_path, target_path, dirs_exist_ok=True)

    # ==================== 私有方法 ====================

    def _get_next_comment_id(self):
        """获取下一个可用的评论 ID。"""
        if not self.comments_path.exists():
            return 0

        editor = self["word/comments.xml"]
        max_id = -1
        for comment_elem in editor.dom.getElementsByTagName("w:comment"):
            comment_id = comment_elem.getAttribute("w:id")
            if comment_id:
                try:
                    max_id = max(max_id, int(comment_id))
                except ValueError:
                    pass
        return max_id + 1

    def _load_existing_comments(self):
        """从文件加载现有评论以启用回复功能。"""
        if not self.comments_path.exists():
            return {}

        editor = self["word/comments.xml"]
        existing = {}

        for comment_elem in editor.dom.getElementsByTagName("w:comment"):
            comment_id = comment_elem.getAttribute("w:id")
            if not comment_id:
                continue

            # 从评论内的 w:p 元素中查找 para_id
            para_id = None
            for p_elem in comment_elem.getElementsByTagName("w:p"):
                para_id = p_elem.getAttribute("w14:paraId")
                if para_id:
                    break

            if not para_id:
                continue

            existing[int(comment_id)] = {"para_id": para_id}

        return existing

    def _setup_comments(self):
        """在解压缩目录中设置评论基础设施。"""
        # 创建或更新 word/people.xml
        people_file = self.word_path / "people.xml"
        self._update_people_xml(people_file)

        # 更新 XML 文件
        self._add_content_type_for_people(self.unpacked_path / "[Content_Types].xml")
        self._add_relationship_for_people(
            self.word_path / "_rels" / "document.xml.rels"
        )

        # 使用 RSID 更新 settings.xml
        self._update_settings(self.word_path / "settings.xml")

    def _update_people_xml(self, path):
        """如果 people.xml 不存在则创建它。"""
        if not path.exists():
            shutil.copy(TEMPLATE_DIR / "people.xml", path)

    def _add_content_type_for_people(self, path):
        """如果 [Content_Types].xml 中还没有，则添加 people.xml 内容类型。"""
        editor = self["[Content_Types].xml"]

        if self._has_override(editor, "/word/people.xml"):
            return

        root = editor.dom.documentElement
        override_xml = '<Override PartName="/word/people.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml"/>'
        editor.append_to(root, override_xml)

    def _add_relationship_for_people(self, path):
        """如果 document.xml.rels 中还没有，则添加 people.xml 关系。"""
        editor = self["word/_rels/document.xml.rels"]

        if self._has_relationship(editor, "people.xml"):
            return

        root = editor.dom.documentElement
        root_tag = root.tagName  # type: ignore
        prefix = root_tag.split(":")[0] + ":" if ":" in root_tag else ""
        next_rid = editor.get_next_rid()

        rel_xml = f'<{prefix}Relationship Id="{next_rid}" Type="http://schemas.microsoft.com/office/2011/relationships/people" Target="people.xml"/>'
        editor.append_to(root, rel_xml)

    def _update_settings(self, path):
        """在 settings.xml 中添加 RSID。"""
        editor = self["word/settings.xml"]
        root = editor.get_node(tag="w:settings")
        prefix = root.tagName.split(":")[0] if ":" in root.tagName else "w"

        # 检查 rsids 部分是否存在
        rsids_elements = editor.dom.getElementsByTagName(f"{prefix}:rsids")

        if not rsids_elements:
            rsids_xml = f'''<{prefix}:rsids>
  <{prefix}:rsidRoot {prefix}:val="{self.rsid}"/>
  <{prefix}:rsid {prefix}:val="{self.rsid}"/>
</{prefix}:rsids}>'''

            compat_elements = editor.dom.getElementsByTagName(f"{prefix}:compat")
            if compat_elements:
                editor.insert_after(compat_elements[0], rsids_xml)
            else:
                editor.append_to(root, rsids_xml)
        else:
            rsids_elem = rsids_elements[0]
            rsid_exists = any(
                elem.getAttribute(f"{prefix}:val") == self.rsid
                for elem in rsids_elem.getElementsByTagName(f"{prefix}:rsid")
            )

            if not rsid_exists:
                rsid_xml = f'<{prefix}:rsid {prefix}:val="{self.rsid}"/>'
                editor.append_to(rsids_elem, rsid_xml)

    def _add_to_comments_xml(
        self, comment_id, para_id, text, author, initials, timestamp
    ):
        """向 comments.xml 添加单个评论。"""
        if not self.comments_path.exists():
            shutil.copy(TEMPLATE_DIR / "comments.xml", self.comments_path)

        editor = self["word/comments.xml"]
        root = editor.get_node(tag="w:comments")

        escaped_text = (
            text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        )
        comment_xml = f'''<w:comment w:id="{comment_id}">
  <w:p w14:paraId="{para_id}" w14:textId="77777777">
    <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r>
    <w:r><w:rPr><w:color w:val="000000"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t>{escaped_text}</w:t></w:r>
  </w:p>
</w:comment>'''
        editor.append_to(root, comment_xml)

    def _add_to_comments_extended_xml(self, para_id, parent_para_id):
        """向 commentsExtended.xml 添加单个评论。"""
        if not self.comments_extended_path.exists():
            shutil.copy(
                TEMPLATE_DIR / "commentsExtended.xml", self.comments_extended_path
            )

        editor = self["word/commentsExtended.xml"]
        root = editor.get_node(tag="w15:commentsEx")

        if parent_para_id:
            xml = f'<w15:commentEx w15:paraId="{para_id}" w15:paraIdParent="{parent_para_id}" w15:done="0"/>'
        else:
            xml = f'<w15:commentEx w15:paraId="{para_id}" w15:done="0"/>'
        editor.append_to(root, xml)

    def _add_to_comments_ids_xml(self, para_id, durable_id):
        """向 commentsIds.xml 添加单个评论。"""
        if not self.comments_ids_path.exists():
            shutil.copy(TEMPLATE_DIR / "commentsIds.xml", self.comments_ids_path)

        editor = self["word/commentsIds.xml"]
        root = editor.get_node(tag="w16cid:commentsIds")

        xml = f'<w16cid:commentId w16cid:paraId="{para_id}" w16cid:durableId="{durable_id}"/>'
        editor.append_to(root, xml)

    def _add_to_comments_extensible_xml(self, durable_id):
        """向 commentsExtensible.xml 添加单个评论。"""
        if not self.comments_extensible_path.exists():
            shutil.copy(
                TEMPLATE_DIR / "commentsExtensible.xml", self.comments_extensible_path
            )

        editor = self["word/commentsExtensible.xml"]
        root = editor.get_node(tag="w16cex:commentsExtensible")

        xml = f'<w16cex:commentExtensible w16cex:durableId="{durable_id}"/>'
        editor.append_to(root, xml)

    def _comment_range_start_xml(self, comment_id):
        """生成评论范围开始的 XML。"""
        return f'<w:commentRangeStart w:id="{comment_id}"/>'

    def _comment_range_end_xml(self, comment_id):
        """生成带有引用 Run 的评论范围结束的 XML。"""
        return f'''<w:commentRangeEnd w:id="{comment_id}"/>
<w:r>
  <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
  <w:commentReference w:id="{comment_id}"/>
</w:r>'''

    def _comment_ref_run_xml(self, comment_id):
        """生成评论引用 Run 的 XML。"""
        return f'''<w:r>
  <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
  <w:commentReference w:id="{comment_id}"/>
</w:r>'''

    def _has_relationship(self, editor, target):
        """检查是否存在具有给定目标的 relationship。"""
        for rel_elem in editor.dom.getElementsByTagName("Relationship"):
            if rel_elem.getAttribute("Target") == target:
                return True
        return False

    def _has_override(self, editor, part_name):
        """检查是否存在具有给定部件名称的 override。"""
        for override_elem in editor.dom.getElementsByTagName("Override"):
            if override_elem.getAttribute("PartName") == part_name:
                return True
        return False

    def _has_author(self, editor, author):
        """检查作者是否已存在于 people.xml 中。"""
        for person_elem in editor.dom.getElementsByTagName("w15:person"):
            if person_elem.getAttribute("w15:author") == author:
                return True
        return False

    def _add_author_to_people(self, author):
        """将作者添加到 people.xml。"""
        people_path = self.word_path / "people.xml"

        if not people_path.exists():
            raise ValueError("people.xml should exist after _setup_comments")

        editor = self["word/people.xml"]
        root = editor.get_node(tag="w15:people")

        if self._has_author(editor, author):
            return

        escaped_author = html.escape(author, quote=True)
        person_xml = f'''<w15:person w15:author="{escaped_author}">
  <w15:presenceInfo w15:providerId="None" w15:userId="{escaped_author}"/>
</w15:person>'''
        editor.append_to(root, person_xml)

    def _ensure_comment_relationships(self):
        """确保 word/_rels/document.xml.rels 包含评论关系。"""
        editor = self["word/_rels/document.xml.rels"]

        if self._has_relationship(editor, "comments.xml"):
            return

        root = editor.dom.documentElement
        root_tag = root.tagName  # type: ignore
        prefix = root_tag.split(":")[0] + ":" if ":" in root_tag else ""
        next_rid_num = int(editor.get_next_rid()[3:])

        rels = [
            (
                next_rid_num,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
                "comments.xml",
            ),
            (
                next_rid_num + 1,
                "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
                "commentsExtended.xml",
            ),
            (
                next_rid_num + 2,
                "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds",
                "commentsIds.xml",
            ),
            (
                next_rid_num + 3,
                "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible",
                "commentsExtensible.xml",
            ),
        ]

        for rel_id, rel_type, target in rels:
            rel_xml = f'<{prefix}Relationship Id="rId{rel_id}" Type="{rel_type}" Target="{target}"/>'
            editor.append_to(root, rel_xml)

    def _ensure_comment_content_types(self):
        """确保 [Content_Types].xml 包含评论内容类型。"""
        editor = self["[Content_Types].xml"]

        if self._has_override(editor, "/word/comments.xml"):
            return

        root = editor.dom.documentElement

        overrides = [
            (
                "/word/comments.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
            ),
            (
                "/word/commentsExtended.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml",
            ),
            (
                "/word/commentsIds.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml",
            ),
            (
                "/word/commentsExtensible.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml",
            ),
        ]

        for part_name, content_type in overrides:
            override_xml = (
                f'<Override PartName="{part_name}" ContentType="{content_type}"/>'
            )
            editor.append_to(root, override_xml)
