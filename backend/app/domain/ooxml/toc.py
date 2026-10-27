"""目录（TOC）处理。

对应原 utils/remove_toc_sdt_test.py、utils/docx_update_toc_title.py、
service/converter.py 的 step4_1_insert_toc_after_cover。

关键差异：
- 原实现插入目录通过 Word COM 的 TablesOfContents.Add，只能在 Windows 运行；
- 这里直接写入 OOXML 域代码（fldChar + instrText），由排版引擎刷新。
  Word / LibreOffice 打开后更新域即可得到真实页码，无需 COM。
"""

from __future__ import annotations

import re

from lxml import etree

from app.core.constants import layout_rules
from app.core.logging import get_logger
from app.core.namespaces import qn, w_attr
from app.domain.ooxml.package import DocxPackage, paragraph_text
from app.domain.ooxml.styles import StyleResolver

logger = get_logger(__name__)

_TOC_FIELD_INSTR = ' TOC \\o "{levels}" \\h \\z \\u '
_PLACEHOLDER_TEXT = "按 F9 更新目录"
_TOC_TITLE_SEARCH = re.compile(r"目\s*录")


class TocProcessor:
    """目录内容控件清理、标题替换与域插入。"""

    def __init__(self, pkg: DocxPackage) -> None:
        self.pkg = pkg

    # ------------------------------------------------------------------
    # 清理
    # ------------------------------------------------------------------
    def remove_toc_blocks(self) -> int:
        """移除文档中已有的自动目录（TOC 内容控件），保留分节符。

        与原实现的差异：原 utils/remove_toc_sdt_test.py 采用"拆壳留内容"，
        会把目录条目变成普通段落 —— 这些条目形如 "1 总论 …… 1"，
        随后会被标题识别误判为正文标题，导致目录重复、标题错乱。
        正确做法是整块删除（内容一并丢弃），只保留其中的 w:sectPr。

        同时保证：若目录块前没有"目录"标题段落，补一个。
        """
        body = self.pkg.body
        removed = 0

        for sdt in list(body.iter(qn("w:sdt"))):
            if not self._is_toc_sdt(sdt):
                continue

            parent = sdt.getparent()
            if parent is None:
                continue

            index = parent.index(sdt)
            preserved = self._extract_section_paragraphs(sdt)

            if not self._has_title_before(parent, index):
                preserved.insert(0, self._build_plain_paragraph(
                    layout_rules()["toc"]["default_title"]
                ))

            for offset, element in enumerate(preserved):
                parent.insert(index + offset, element)

            parent.remove(sdt)
            removed += 1

        if removed:
            logger.info("移除已有自动目录 %d 处（保留分节符）", removed)
        return removed

    def remove_content_controls(self) -> int:
        """兼容旧接口：拆除目录内容控件。

        默认走 remove_toc_blocks 的整块删除语义，避免目录条目污染标题识别。
        """
        return self.remove_toc_blocks()

    @staticmethod
    def _is_toc_sdt(sdt: etree._Element) -> bool:
        """依据 docPartGallery=Table of Contents 精确判定目录控件。

        比"文本以目录开头"可靠：前者是 Word 的结构标记，后者会误伤正文。
        """
        for node in sdt.iter(qn("w:docPartGallery")):
            if (node.get(w_attr("val")) or "").strip().lower() == "table of contents":
                return True
        return False

    @staticmethod
    def _extract_section_paragraphs(sdt: etree._Element) -> list[etree._Element]:
        """从目录块中提取携带 sectPr 的段落（分节信息不能丢）。"""
        preserved: list[etree._Element] = []
        for paragraph in sdt.iter(qn("w:p")):
            sect = paragraph.find(f"{qn('w:pPr')}/{qn('w:sectPr')}")
            if sect is not None:
                preserved.append(etree.fromstring(etree.tostring(paragraph)))
        return preserved

    @staticmethod
    def _has_title_before(parent: etree._Element, index: int) -> bool:
        """检查目录块之前是否已有"目录"标题段落。"""
        for sibling in reversed(parent[:index]):
            if sibling.tag != qn("w:p"):
                continue
            text = re.sub(r"\s+", "", paragraph_text(sibling))
            if not text:
                continue
            return bool(_TOC_TITLE_SEARCH.search(text))
        return False

    @staticmethod
    def _build_plain_paragraph(text: str) -> etree._Element:
        p = etree.Element(qn("w:p"))
        run = etree.SubElement(p, qn("w:r"))
        t = etree.SubElement(run, qn("w:t"))
        t.set(qn("xml:space"), "preserve")
        t.text = text
        return p

    # ------------------------------------------------------------------
    # 标题替换
    # ------------------------------------------------------------------
    def replace_title(self, new_title: str | None = None) -> bool:
        """把首个目录标题文本替换为指定标题，保留其原有格式。"""
        rules = layout_rules()["toc"]
        target_title = new_title or rules["default_title"]
        keywords = {kw.strip() for kw in rules["keywords"]}

        for node in self.pkg.body.iter(qn("w:t")):
            text = (node.text or "").strip()
            if text in keywords:
                node.text = target_title
                node.set(qn("xml:space"), "preserve")
                logger.info("目录标题已替换为: %s", target_title)
                return True

        logger.info("未找到目录标题，跳过替换")
        return False

    # ------------------------------------------------------------------
    # 插入
    # ------------------------------------------------------------------
    def insert_field(self, title: str | None = None, levels: str | None = None) -> bool:
        """在正文首个一级标题前插入目录标题与 TOC 域。

        原实现的插入位置依赖 Word 的分页信息（GoTo 第 4 页），无法在纯 OOXML 下复现。
        这里改用语义定位：第一个 Heading 1 段落之前，语义上等价于"正文开始处"，
        且不依赖排版引擎计算页码，结果稳定可测。
        """
        rules = layout_rules()["toc"]
        toc_title = title or rules["default_title"]
        toc_levels = levels or rules["levels"]

        body = self.pkg.body
        anchor = self._find_first_heading_paragraph()
        if anchor is None:
            logger.warning("未找到一级标题，无法插入目录")
            return False

        title_para = self._build_title_paragraph(toc_title)
        field_para = self._build_toc_field_paragraph(toc_levels)

        index = body.index(anchor)
        body.insert(index, title_para)
        body.insert(index + 1, field_para)

        logger.info("目录已插入（标题: %s, 级别: %s）", toc_title, toc_levels)
        return True

    def _find_first_heading_paragraph(self) -> etree._Element | None:
        resolver = StyleResolver(self.pkg)
        for paragraph in self.pkg.paragraphs():
            style_el = paragraph.find(f"{qn('w:pPr')}/{qn('w:pStyle')}")
            if style_el is None:
                continue
            style_id = style_el.get(w_attr("val")) or ""
            if not resolver.is_heading(style_id):
                continue
            # 跳过目录自身
            if paragraph_text(paragraph) in ("", "目录"):
                continue
            return paragraph
        return None

    @staticmethod
    def _build_title_paragraph(title: str) -> etree._Element:
        p = etree.Element(qn("w:p"))
        ppr = etree.SubElement(p, qn("w:pPr"))
        jc = etree.SubElement(ppr, qn("w:jc"))
        jc.set(w_attr("val"), "center")

        run = etree.SubElement(p, qn("w:r"))
        rpr = etree.SubElement(run, qn("w:rPr"))
        etree.SubElement(rpr, qn("w:b"))
        text = etree.SubElement(run, qn("w:t"))
        text.set(qn("xml:space"), "preserve")
        text.text = title
        return p

    @staticmethod
    def _build_toc_field_paragraph(levels: str) -> etree._Element:
        p = etree.Element(qn("w:p"))

        def fld_char(kind: str) -> etree._Element:
            run = etree.SubElement(p, qn("w:r"))
            el = etree.SubElement(run, qn("w:fldChar"))
            el.set(w_attr("fldCharType"), kind)
            return run

        fld_char("begin")

        instr_run = etree.SubElement(p, qn("w:r"))
        instr = etree.SubElement(instr_run, qn("w:instrText"))
        instr.set(qn("xml:space"), "preserve")
        instr.text = _TOC_FIELD_INSTR.format(levels=levels)

        fld_char("separate")

        placeholder_run = etree.SubElement(p, qn("w:r"))
        placeholder = etree.SubElement(placeholder_run, qn("w:t"))
        placeholder.text = _PLACEHOLDER_TEXT

        fld_char("end")
        return p

    # ------------------------------------------------------------------
    # 检测
    # ------------------------------------------------------------------
    def has_toc_field(self) -> bool:
        """判断文档是否已包含 TOC 域。"""
        for instr in self.pkg.body.iter(qn("w:instrText")):
            if instr.text and "TOC" in instr.text.upper():
                return True
        return False
