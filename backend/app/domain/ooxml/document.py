"""文档拆分与合并。

对应原 utils/docx_split.py、utils/merge_docx_without_com.py、
utils/merge_docs_preserve_orientation.py、service/converter.py 的 step1/step4。

原实现的两处关键缺陷在这里被修掉：
1. merge_docx_without_com 直接 deepcopy 正文元素，**没有重映射关系 ID**，
   被追加文档中的图片/页眉引用会指向错误目标（跨文档合并时图片丢失）。
   这里新增 RelationshipRemapper，按部件内容去重后重建关系。
2. merge_docs_preserve_orientation 依赖 COM 的 InsertFile/InsertBreak，
   只能在 Windows 运行。这里全部用 OOXML 结构操作完成。
"""

from __future__ import annotations

import posixpath
import re
from copy import deepcopy
from dataclasses import dataclass, field

from lxml import etree

from app.core.constants import layout_rules
from app.core.errors import LayoutError
from app.core.logging import get_logger
from app.core.namespaces import (
    PART_CONTENT_TYPES,
    PART_DOCUMENT,
    PART_DOCUMENT_RELS,
    PART_SETTINGS,
    PART_STYLES,
    R,
    qn,
    w_attr,
)
from app.domain.ooxml.package import DocxPackage, paragraph_text
from app.domain.ooxml.styles import StyleResolver
from app.domain.ooxml.toc import is_generated_toc_line

logger = get_logger(__name__)

_XML_SPACE = qn("xml:space")

# 关系类型中需要随文档一起迁移的内部部件
_RELOCATABLE_TYPES = (
    "image",
    "header",
    "footer",
    "footnotes",
    "endnotes",
    "hyperlink",
)


@dataclass
class SplitResult:
    """拆分结果：封面包与正文包。"""

    cover: DocxPackage
    content: DocxPackage
    project_name: str | None = None
    split_reason: str = ""


class DocumentSplitter:
    """把源文档拆成"封面/目录前"与"正文"两个包。"""

    def __init__(self, pkg: DocxPackage) -> None:
        self.pkg = pkg

    # ------------------------------------------------------------------
    def split(self) -> SplitResult:
        body = self.pkg.body
        children = list(body)

        toc_keywords = {kw.strip() for kw in layout_rules()["toc"]["keywords"]}
        toc_index = self._find_toc_index(children, toc_keywords)
        content_start = self._find_content_start(children, toc_index + 1 if toc_index >= 0 else 0)

        if content_start is None:
            raise LayoutError("无法定位正文起始位置，文档结构不符合排版要求")

        cover = self.pkg.clone()
        content = self.pkg.clone()

        self._truncate_after(cover, content_start)
        self._truncate_before(content, content_start)

        reason = "toc" if toc_index >= 0 else "first_heading"
        project_name = self._guess_project_name(cover)

        logger.info(
            "文档拆分完成: 目录位置=%s, 正文起始=%d, 依据=%s",
            toc_index,
            content_start,
            reason,
        )
        return SplitResult(
            cover=cover,
            content=content,
            project_name=project_name,
            split_reason=reason,
        )

    # ------------------------------------------------------------------
    @staticmethod
    def _find_toc_index(children: list[etree._Element], keywords: set[str]) -> int:
        normalized = {re.sub(r"\s+", "", item) for item in keywords}
        for index, child in enumerate(children):
            if child.tag != qn("w:p"):
                continue
            if re.sub(r"\s+", "", paragraph_text(child)) in normalized:
                return index
        return -1

    def _find_content_start(self, children: list[etree._Element], start: int) -> int | None:
        """从 start 起找第一个真正的正文标题段落。"""
        resolver = StyleResolver(self.pkg)

        for index in range(start, len(children)):
            child = children[index]
            if child.tag != qn("w:p"):
                continue
            style_el = child.find(f"{qn('w:pPr')}/{qn('w:pStyle')}")
            if style_el is None:
                continue
            style_id = style_el.get(w_attr("val")) or ""
            if not resolver.is_heading(style_id):
                continue
            if is_generated_toc_line(paragraph_text(child)):
                continue
            return index

        # 退路：第一个编号标题。旧目录行（末尾粘着页码）不算正文起点。
        for index in range(start, len(children)):
            child = children[index]
            if child.tag != qn("w:p"):
                continue
            text = paragraph_text(child)
            if text and text[0].isdigit() and not is_generated_toc_line(text):
                return index
        return None

    @staticmethod
    def _truncate_after(pkg: DocxPackage, keep_count: int) -> None:
        """只保留前 keep_count 个 body 子元素，末尾 sectPr 保留。"""
        body = pkg.body
        children = list(body)
        sect_pr = children[-1] if children and children[-1].tag == qn("w:sectPr") else None

        for child in children[keep_count:]:
            if child is sect_pr:
                continue
            body.remove(child)

        if sect_pr is not None and sect_pr.getparent() is None:
            body.append(sect_pr)

    @staticmethod
    def _truncate_before(pkg: DocxPackage, start: int) -> None:
        """删除前 start 个 body 子元素，末尾 sectPr 保留。"""
        body = pkg.body
        children = list(body)
        sect_pr = children[-1] if children and children[-1].tag == qn("w:sectPr") else None

        for child in children[:start]:
            body.remove(child)

        if sect_pr is not None and sect_pr.getparent() is None:
            body.append(sect_pr)

    @staticmethod
    def _guess_project_name(cover: DocxPackage) -> str | None:
        """从封面文本中猜测项目名称（尽力而为，不作为失败条件）。"""
        candidates: list[str] = []
        for paragraph in cover.paragraphs():
            text = paragraph_text(paragraph)
            if not text or len(text) < 4:
                continue
            candidates.append(text)

        for text in candidates:
            if "项目" in text:
                return text
        return candidates[0] if candidates else None


# ======================================================================
# 关系重映射
# ======================================================================
@dataclass
class _RelEntry:
    rel_id: str
    rel_type: str
    target: str
    external: bool


class RelationshipRemapper:
    """把来源文档的关系迁移到目标文档，返回 old_id -> new_id 映射。

    这是合并正确性的关键：rId 只在单个文档内唯一，
    跨文档直接拼接会让图片、页眉、超链接指向错误目标。
    """

    def __init__(
        self,
        base: DocxPackage,
        source: DocxPackage,
        only_types: tuple[str, ...] | None = None,
    ) -> None:
        """only_types 用于只迁移关心的关系类型。

        封面设计人员表移植时只需要 image（表内可能嵌图），
        全量迁移会把页眉页脚等无干部件一起搬过来。
        """
        self.base = base
        self.source = source
        self._only_types = only_types
        self._counter = self._max_existing_id()
        self.mapping: dict[str, str] = {}

    def _max_existing_id(self) -> int:
        highest = 0
        for rel in self._iter_base_rels():
            rel_id = rel.get("Id") or ""
            digits = "".join(ch for ch in rel_id if ch.isdigit())
            if digits:
                highest = max(highest, int(digits))
        return highest

    def _iter_base_rels(self):
        if not self.base.has(PART_DOCUMENT_RELS):
            return []
        root = self.base.xml(PART_DOCUMENT_RELS)
        return root.findall(qn("pr:Relationship"))

    def _next_id(self) -> str:
        self._counter += 1
        return f"rId{self._counter}"

    # ------------------------------------------------------------------
    def remap(self) -> dict[str, str]:
        entries = self._read_source_rels()
        if not entries:
            return {}

        base_rels_root = self._ensure_base_rels_root()

        for entry in entries:
            if self._only_types is not None and not any(
                allowed in entry.rel_type for allowed in self._only_types
            ):
                continue

            new_id = self._next_id()

            if entry.external:
                self._append_rel(base_rels_root, new_id, entry.rel_type, entry.target, external=True)
                self.mapping[entry.rel_id] = new_id
                continue

            if not self._should_relocate(entry.rel_type):
                continue

            source_part = self._resolve_part_name(entry.target)
            if not self.source.has(source_part):
                logger.debug("跳过悬空关系: %s -> %s", entry.rel_id, source_part)
                continue

            data = self.source.read(source_part)
            target_part = self._ensure_target_part(source_part, data)
            relative_target = posixpath.relpath(target_part, "word")
            self._append_rel(base_rels_root, new_id, entry.rel_type, relative_target)
            self.mapping[entry.rel_id] = new_id

        if self.mapping:
            logger.info("迁移 %d 条文档关系", len(self.mapping))
        return self.mapping

    @staticmethod
    def _should_relocate(rel_type: str) -> bool:
        return any(key in rel_type for key in _RELOCATABLE_TYPES)

    @staticmethod
    def _resolve_part_name(target: str) -> str:
        """document.xml.rels 中的 Target 相对于 word/。"""
        if target.startswith("/"):
            return target.lstrip("/")
        return posixpath.normpath(posixpath.join("word", target))

    def _read_source_rels(self) -> list[_RelEntry]:
        if not self.source.has(PART_DOCUMENT_RELS):
            return []
        root = self.source.xml(PART_DOCUMENT_RELS)
        entries: list[_RelEntry] = []
        for rel in root.findall(qn("pr:Relationship")):
            rel_id = rel.get("Id")
            if not rel_id:
                continue
            entries.append(
                _RelEntry(
                    rel_id=rel_id,
                    rel_type=rel.get("Type") or "",
                    target=rel.get("Target") or "",
                    external=(rel.get("TargetMode") or "") == "External",
                )
            )
        return entries

    def _ensure_base_rels_root(self) -> etree._Element:
        if self.base.has(PART_DOCUMENT_RELS):
            return self.base.xml(PART_DOCUMENT_RELS)

        root = etree.Element(qn("pr:Relationships"), nsmap={None: "http://schemas.openxmlformats.org/package/2006/relationships"})
        self.base.add_part(PART_DOCUMENT_RELS, etree.tostring(root, xml_declaration=True, encoding="UTF-8"))
        return self.base.xml(PART_DOCUMENT_RELS)

    @staticmethod
    def _append_rel(
        root: etree._Element,
        rel_id: str,
        rel_type: str,
        target: str,
        *,
        external: bool = False,
    ) -> None:
        rel = etree.SubElement(root, qn("pr:Relationship"))
        rel.set("Id", rel_id)
        rel.set("Type", rel_type)
        rel.set("Target", target)
        if external:
            rel.set("TargetMode", "External")

    def _ensure_target_part(self, source_part: str, data: bytes) -> str:
        """若内容已存在则复用，否则复制到不冲突的新路径。"""
        if self.base.has(source_part) and self.base.read(source_part) == data:
            return source_part

        candidate = source_part
        if self.base.has(candidate):
            stem, dot, ext = source_part.rpartition(".")
            directory = posixpath.dirname(source_part)
            index = 1
            while True:
                candidate = f"{stem}_{index}{dot}{ext}" if dot else f"{source_part}_{index}"
                if directory:
                    candidate = posixpath.join(directory, posixpath.basename(candidate))
                if not self.base.has(candidate):
                    break
                index += 1

        self.base.add_part(candidate, data)
        self._ensure_content_type(candidate)
        return candidate

    def _ensure_content_type(self, part_name: str) -> None:
        """确保 [Content_Types].xml 声明了该扩展名。"""
        if not self.base.has(PART_CONTENT_TYPES):
            return
        _, dot, ext = part_name.rpartition(".")
        if not dot:
            return
        ext = ext.lower()

        root = self.base.xml(PART_CONTENT_TYPES)
        for default in root.findall(qn("ct:Default")):
            if (default.get("Extension") or "").lower() == ext:
                return

        default = etree.SubElement(root, qn("ct:Default"))
        default.set("Extension", ext)
        default.set("ContentType", _CONTENT_TYPE_BY_EXT.get(ext, "application/octet-stream"))


_CONTENT_TYPE_BY_EXT = {
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "gif": "image/gif",
    "bmp": "image/bmp",
    "tif": "image/tiff",
    "tiff": "image/tiff",
    "emf": "image/x-emf",
    "wmf": "image/x-wmf",
    "svg": "image/svg+xml",
    "xml": "application/xml",
}


def rewrite_relationship_ids(elements: list[etree._Element], mapping: dict[str, str]) -> int:
    """把元素树中 r: 命名空间下的 ID 属性替换为新 ID。"""
    if not mapping:
        return 0

    rewritten = 0
    rel_attrs = (
        f"{{{R}}}embed",
        f"{{{R}}}id",
        f"{{{R}}}link",
        f"{{{R}}}pict",
        f"{{{R}}}href",
    )

    for root in elements:
        for node in root.iter():
            for attr in rel_attrs:
                value = node.get(attr)
                if value and value in mapping:
                    node.set(attr, mapping[value])
                    rewritten += 1
    return rewritten


# ======================================================================
# 合并
# ======================================================================
@dataclass
class MergeReport:
    appended_body_elements: int = 0
    remapped_relationships: int = 0
    rewritten_references: int = 0
    merged_styles: int = 0
    base_sections: int = 0
    final_sections: int = 0
    warnings: list[str] = field(default_factory=list)


class DocumentMerger:
    """把若干文档追加到基础文档末尾，每份追加内容独立成节。"""

    def __init__(self, base: DocxPackage) -> None:
        self.base = base

    def append(self, source: DocxPackage) -> MergeReport:
        report = MergeReport()

        remapper = RelationshipRemapper(self.base, source)
        mapping = remapper.remap()
        report.remapped_relationships = len(mapping)

        # 0. 先合并样式：content 引用的 pStyle 必须能在合并后的 styles.xml 中解析。
        #    封面包与正文包样式表不同，若不合并会出现"引用了未定义的段落样式"。
        report.merged_styles = self._merge_styles(source)

        base_body = self.base.body
        source_body = source.body

        base_sections = len(self.base.section_properties())
        report.base_sections = base_sections

        # 1. 把 base 的末尾 sectPr 转成"段落级节属性"，形成分节
        self._demote_trailing_section_properties(base_body)

        # 2. 追加 source 正文（不含其末尾 sectPr）
        #    关系 ID 只改追加进来的内容。封面里的图标如果被一起改写，
        #    r:embed 会指到正文的页眉或尾注，左上角公司图标就显示不出来。
        appended = 0
        appended_nodes: list[etree._Element] = []
        for child in list(source_body):
            if child.tag == qn("w:sectPr"):
                continue
            cloned = deepcopy(child)
            base_body.append(cloned)
            appended_nodes.append(cloned)
            appended += 1
        report.appended_body_elements = appended

        # 3. 把 source 的末尾 sectPr 作为整篇的新结束节属性
        source_sect = source_body.find(qn("w:sectPr"))
        if source_sect is not None:
            cloned_sect = deepcopy(source_sect)
            base_body.append(cloned_sect)
            appended_nodes.append(cloned_sect)
        else:
            report.warnings.append("被追加文档缺少末尾 sectPr，页面设置可能沿用前一节")

        # 4. 只重写追加内容里的 r:id
        report.rewritten_references = rewrite_relationship_ids(appended_nodes, mapping)

        # 只有原稿里有 VML 流程图文本框时，才沿用原稿字距。普通文档仍用封面模板的设置。
        if _has_vml_textbox(source):
            self._keep_source_text_layout(source)

        report.final_sections = len(self.base.section_properties())
        logger.info(
            "合并完成: 追加 %d 个元素, 迁移关系 %d 条, 节数 %d -> %d",
            appended,
            len(mapping),
            report.base_sections,
            report.final_sections,
        )
        return report

    # ------------------------------------------------------------------
    def _merge_styles(self, source: DocxPackage) -> int:
        """把 source 的样式并入 base 的 styles.xml（缺 styleId 才补）。

        content 与封面模板的样式表不同，不合并会导致正文引用的
        pStyle 悬空（例如源文档用 a/a0/a1/Normal 这些 WPS 自定义样式）。
        """
        if not self.base.has(PART_STYLES) or not source.has(PART_STYLES):
            return 0

        base_root = self.base.xml(PART_STYLES)
        base_by_id = {
            style.get(w_attr("styleId")): style
            for style in base_root.findall(qn("w:style"))
            if style.get(w_attr("styleId"))
        }
        used_ids = set(base_by_id)

        source_root = source.xml(PART_STYLES)
        merged = 0
        rename: dict[str, str] = {}

        for style in source_root.findall(qn("w:style")):
            style_id = style.get(w_attr("styleId"))
            if not style_id:
                continue

            existing = base_by_id.get(style_id)
            if existing is None:
                base_root.append(deepcopy(style))
                base_by_id[style_id] = style
                used_ids.add(style_id)
                merged += 1
                continue

            # 同一个 styleId 在封面和正文里可能不是同一个样式。
            # 例如正文的 "1" 是标题一，封面的 "1" 是正文。这时必须换一个 id，
            # 否则正文标题会被封面的正文样式盖掉。
            if _style_display_name(existing) == _style_display_name(style):
                imported = {name.lower() for name in layout_rules().get("imported_styles", [])}
                if _style_display_name(existing) in imported:
                    clone = deepcopy(style)
                    clone.set(w_attr("styleId"), style_id)
                    parent = existing.getparent()
                    if parent is not None:
                        parent.replace(existing, clone)
                        base_by_id[style_id] = clone
                        merged += 1
                continue

            new_id = style_id
            suffix = 1
            while new_id in used_ids:
                new_id = f"{style_id}_{suffix}"
                suffix += 1
            clone = deepcopy(style)
            clone.set(w_attr("styleId"), new_id)
            base_root.append(clone)
            used_ids.add(new_id)
            rename[style_id] = new_id
            merged += 1

        if rename:
            _rewrite_style_ids(source.body, rename)
            logger.info("样式 ID 冲突已重映射 %d 处", len(rename))

        if merged:
            logger.info("合并样式 %d 条", merged)
        return merged

    def _keep_source_text_layout(self, source: DocxPackage) -> None:
        """沿用正文的字距和兼容模式，避免文本框里的字被裁切。"""
        if not self.base.has(PART_SETTINGS) or not source.has(PART_SETTINGS):
            return

        base_root = self.base.xml(PART_SETTINGS)
        source_root = source.xml(PART_SETTINGS)
        for tag in ("w:characterSpacingControl", "w:noPunctuationKerning"):
            _sync_child(base_root, source_root, qn(tag))

        source_compat = source_root.find(qn("w:compat"))
        if source_compat is None:
            return
        base_compat = base_root.find(qn("w:compat"))
        if base_compat is None:
            base_compat = etree.SubElement(base_root, qn("w:compat"))
        for tag in (
            "w:doNotWrapTextWithPunct",
            "w:doNotUseEastAsianBreakRules",
            "w:balanceSingleByteDoubleByteWidth",
        ):
            _sync_child(base_compat, source_compat, qn(tag))
        _sync_compat_mode(base_compat, source_compat)

    @staticmethod
    def _demote_trailing_section_properties(body: etree._Element) -> None:
        """把 body 末尾的 w:sectPr 包进 w:p/w:pPr，使其成为分节符。"""
        last = body[-1] if len(body) else None
        if last is None or last.tag != qn("w:sectPr"):
            return

        body.remove(last)
        paragraph = etree.SubElement(body, qn("w:p"))
        ppr = etree.SubElement(paragraph, qn("w:pPr"))
        ppr.append(last)


def _has_vml_textbox(pkg: DocxPackage) -> bool:
    return any(node.tag == qn("v:textbox") for node in pkg.body.iter())


def _sync_child(base_parent: etree._Element, source_parent: etree._Element, tag: str) -> None:
    """按源文档有没有这个节点，增删或替换基础文档里的同名节点。"""
    source_el = source_parent.find(tag)
    base_el = base_parent.find(tag)
    if source_el is None:
        if base_el is not None:
            base_parent.remove(base_el)
        return
    clone = deepcopy(source_el)
    if base_el is None:
        base_parent.append(clone)
    else:
        base_parent.replace(base_el, clone)


def _sync_compat_mode(base_compat: etree._Element, source_compat: etree._Element) -> None:
    source_mode = _compat_mode(source_compat)
    if source_mode is None:
        return
    base_mode = _compat_mode(base_compat)
    clone = deepcopy(source_mode)
    if base_mode is None:
        base_compat.append(clone)
    else:
        base_compat.replace(base_mode, clone)


def _compat_mode(compat: etree._Element) -> etree._Element | None:
    for element in compat.findall(qn("w:compatSetting")):
        if element.get(w_attr("name")) == "compatibilityMode":
            return element
    return None


def _style_display_name(style: etree._Element) -> str:
    name_el = style.find(qn("w:name"))
    if name_el is None:
        return ""
    return (name_el.get(w_attr("val")) or "").lower()


def _rewrite_style_ids(root: etree._Element, rename: dict[str, str]) -> None:
    """把段落、字符、表格样式引用改到新的 styleId。"""
    tags = {qn("w:pStyle"), qn("w:rStyle"), qn("w:tblStyle")}
    for node in root.iter():
        if node.tag not in tags:
            continue
        value = node.get(w_attr("val")) or ""
        if value in rename:
            node.set(w_attr("val"), rename[value])
