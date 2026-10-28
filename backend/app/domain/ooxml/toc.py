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
import zipfile

from lxml import etree

from pathlib import Path

from app.core.constants import layout_rules, pt_to_half_points
from app.core.logging import get_logger
from app.core.namespaces import PART_STYLES, qn, w_attr
from app.domain.ooxml.package import DocxPackage, paragraph_text
from app.domain.ooxml.styles import StyleProcessor, StyleResolver

logger = get_logger(__name__)

_TOC_FIELD_INSTR = ' TOC \\o "{levels}" \\h \\z \\u '
_PLACEHOLDER_TEXT = "按 F9 更新目录"
_TOC_TITLE_SEARCH = re.compile(r"目\s*录")
# 旧目录行：编号后粘着页码，如 "1 总论1"、"1.1项目背景3"。正文标题 "1 总论" 不以数字结尾。
_TOC_ENTRY_LINE = re.compile(r"^\d+(?:\.\d+)*\D+\d+$")


def is_generated_toc_line(text: str) -> bool:
    """判断一段是不是旧目录条目（编号标题后面直接粘着页码）。"""
    compact = re.sub(r"\s+", "", text or "")
    if not compact or len(compact) > 40:
        return False
    return _TOC_ENTRY_LINE.fullmatch(compact) is not None


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

        removed += self._remove_plain_toc_entries()

        if removed:
            logger.info("移除已有目录 %d 处（保留分节符与正文标题）", removed)
        return removed

    def _remove_plain_toc_entries(self) -> int:
        """删掉「目录」标题后面、正文标题前面的纯文本目录行。

        这些行不是内容控件，形如 "1 总论1"。不能按标题文字去删，
        正文里的 "1 总论" 不以页码结尾，会留在原地。
        """
        body = self.pkg.body
        removing = False
        removed = 0

        for child in list(body):
            if child.tag != qn("w:p"):
                if removing:
                    removing = False
                continue

            text = paragraph_text(child)
            compact = re.sub(r"\s+", "", text)
            if compact and _TOC_TITLE_SEARCH.fullmatch(compact):
                removing = True
                continue

            if not removing:
                continue

            if not text or is_generated_toc_line(text):
                body.remove(child)
                removed += 1
                continue

            removing = False

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

        self.apply_entry_fonts()
        self.flatten_entry_indents()
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
            text = paragraph_text(paragraph)
            if text in ("", "目录") or is_generated_toc_line(text):
                continue
            return paragraph
        return None

    def flatten_entry_indents(self) -> int:
        """目录条目左侧顶格，不按级别缩进。"""
        if not self.pkg.has(PART_STYLES):
            return 0

        root = self.pkg.xml(PART_STYLES)
        touched = 0
        for style in root.findall(qn("w:style")):
            name_el = style.find(qn("w:name"))
            name = (name_el.get(w_attr("val")) if name_el is not None else "") or ""
            if not name.lower().startswith("toc "):
                continue
            _force_left_flush(_ensure_ppr(style))
            touched += 1
        return touched

    def apply_entry_fonts(self) -> int:
        """目录条目样式：中文宋体四号，西文（引导点、页码）Times New Roman 四号。"""
        fonts = layout_rules()["toc"]["fonts"]
        return StyleProcessor(self.pkg).apply_style_font(
            ["toc 1", "toc 2", "toc 3"],
            east_asia=fonts["entry_east_asia"],
            ascii_font=fonts["entry_ascii"],
            size_pt=float(fonts["size_pt"]),
            bold=False,
        )

    def _build_title_paragraph(self, title: str) -> etree._Element:
        fonts = layout_rules()["toc"]["fonts"]
        size = str(pt_to_half_points(float(fonts["size_pt"])))

        p = etree.Element(qn("w:p"))
        ppr = etree.SubElement(p, qn("w:pPr"))
        jc = etree.SubElement(ppr, qn("w:jc"))
        jc.set(w_attr("val"), "center")

        run = etree.SubElement(p, qn("w:r"))
        rpr = etree.SubElement(run, qn("w:rPr"))
        rfonts = etree.SubElement(rpr, qn("w:rFonts"))
        rfonts.set(w_attr("ascii"), fonts["title_ascii"])
        rfonts.set(w_attr("hAnsi"), fonts["title_ascii"])
        rfonts.set(w_attr("eastAsia"), fonts["title_east_asia"])
        for tag in ("w:sz", "w:szCs"):
            node = etree.SubElement(rpr, qn(tag))
            node.set(w_attr("val"), size)
        if fonts.get("title_bold", True):
            etree.SubElement(rpr, qn("w:b"))
            etree.SubElement(rpr, qn("w:bCs"))

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


def graft_refreshed_toc(base_path: Path, refreshed_path: Path) -> str:
    """域刷新若丢掉图片，就把刷新后的目录块嫁接回刷新前的文档。

    LibreOffice 的整包 store 曾经把 word/media 清空。刷新前的包里图片、
    页眉和封面都在，只缺展开后的目录。目录块从刷新结果里取回来写回原包。
    """
    base_media = _count_media(base_path)
    refreshed_media = _count_media(refreshed_path)
    if base_media == 0 or refreshed_media >= base_media:
        return "refresh_kept"

    with DocxPackage.open(refreshed_path) as refreshed:
        payloads = [etree.tostring(element) for element in _toc_result_elements(refreshed)]

    if not payloads:
        logger.warning("域刷新后图片减少，但没有可嫁接的目录块")
        return "no_toc_result"

    with DocxPackage.open(base_path) as base:
        if not _replace_toc_anchor(base, payloads):
            logger.warning("域刷新后图片减少，原文档里没有 TOC 域可替换")
            return "no_anchor"
        base.save(refreshed_path)

    logger.info(
        "目录已从刷新结果嫁接回原文档（图片 %d → 刷新后 %d）",
        base_media,
        refreshed_media,
    )
    return "grafted"


def strip_toc_leading_tabs(path: Path) -> int:
    """成品目录里，标题文字前面的制表符删掉。页码前的引导点保留。"""
    if not path.is_file():
        return 0
    removed = 0
    with DocxPackage.open(path) as pkg:
        styled = TocProcessor(pkg).flatten_entry_indents()
        targets = _toc_result_elements(pkg)
        if not targets and not styled:
            return 0
        for element in targets:
            before = len(list(element.iter(qn("w:tab"))))
            _flatten_toc_paragraphs(element)
            after = len(list(element.iter(qn("w:tab"))))
            removed += before - after
        pkg.save(path)
    return removed


def _count_media(path: Path) -> int:
    with zipfile.ZipFile(path) as archive:
        return sum(1 for name in archive.namelist() if name.startswith("word/media/"))


def _toc_result_elements(pkg: DocxPackage) -> list[etree._Element]:
    """取出刷新后的目录块：优先内容控件，否则取带 PAGEREF 的段落。"""
    body = pkg.body
    for child in body:
        if child.tag == qn("w:sdt") and TocProcessor._is_toc_sdt(child):
            return [child]

    found: list[etree._Element] = []
    for child in body:
        if child.tag != qn("w:p"):
            continue
        instructions = "".join(node.text or "" for node in child.iter(qn("w:instrText")))
        if "PAGEREF" in instructions.upper() or "HYPERLINK" in instructions.upper():
            found.append(child)
    return found


def _replace_toc_anchor(pkg: DocxPackage, payloads: list[bytes]) -> bool:
    anchor = None
    for paragraph in pkg.body.iter(qn("w:p")):
        for instr in paragraph.iter(qn("w:instrText")):
            if instr.text and "TOC" in instr.text.upper() and "PAGEREF" not in instr.text.upper():
                anchor = paragraph
                break
        if anchor is not None:
            break

    if anchor is None:
        return False

    parent = anchor.getparent()
    if parent is None:
        return False

    index = parent.index(anchor)
    for offset, payload in enumerate(payloads):
        element = etree.fromstring(payload)
        # 刷新结果里的目录块可能夹着分节符，那些页眉页脚关系在原文档里不存在。
        for sect in list(element.iter(qn("w:sectPr"))):
            owner = sect.getparent()
            if owner is not None:
                owner.remove(sect)
        _flatten_toc_paragraphs(element)
        parent.insert(index + offset, element)
    parent.remove(anchor)
    return True


def _ensure_ppr(style: etree._Element) -> etree._Element:
    ppr = style.find(qn("w:pPr"))
    if ppr is None:
        ppr = etree.SubElement(style, qn("w:pPr"))
    return ppr


def _drop_indent_tab_stops(ppr: etree._Element) -> None:
    """去掉把标题往右推的左侧制表位，保留页码前的右对齐引导点。"""
    tabs = ppr.find(qn("w:tabs"))
    if tabs is None:
        return
    for tab in list(tabs.findall(qn("w:tab"))):
        kind = (tab.get(w_attr("val")) or "").lower()
        if kind != "right":
            tabs.remove(tab)
    if len(tabs) == 0:
        ppr.remove(tabs)


def _drop_leading_tabs(paragraph: etree._Element) -> None:
    """删掉标题文字前面的制表符。页码前面的引导点制表符留着。"""
    seen_text = False
    for node in list(paragraph.iter()):
        if node.tag == qn("w:t") and (node.text or "").strip():
            seen_text = True
            continue
        if node.tag == qn("w:tab") and not seen_text:
            parent = node.getparent()
            # 只删运行里的制表符。w:tabs 里的 w:tab 是制表位，不是字符。
            if parent is not None and parent.tag == qn("w:r"):
                parent.remove(node)


def _force_left_flush(ppr: etree._Element) -> None:
    """段落顶格、左对齐。"""
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        ind = etree.SubElement(ppr, qn("w:ind"))
    for attr in ("start", "startChars"):
        if w_attr(attr) in ind.attrib:
            del ind.attrib[w_attr(attr)]
    for attr in ("left", "leftChars", "firstLine", "firstLineChars", "hanging", "hangingChars"):
        ind.set(w_attr(attr), "0")
    jc = ppr.find(qn("w:jc"))
    if jc is None:
        jc = etree.SubElement(ppr, qn("w:jc"))
    jc.set(w_attr("val"), "left")
    _drop_indent_tab_stops(ppr)


def _flatten_toc_paragraphs(root: etree._Element) -> None:
    """刷新后的目录条目去掉直接缩进，左侧顶格。目录标题保持居中。"""
    for paragraph in root.iter(qn("w:p")):
        compact = re.sub(r"\s+", "", paragraph_text(paragraph))
        if compact and _TOC_TITLE_SEARCH.fullmatch(compact):
            continue
        ppr = paragraph.find(qn("w:pPr"))
        if ppr is None:
            ppr = etree.Element(qn("w:pPr"))
            paragraph.insert(0, ppr)
        _force_left_flush(ppr)
        _drop_leading_tabs(paragraph)
