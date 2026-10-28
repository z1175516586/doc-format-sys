"""样式解析与规范化。

对应原 utils/fix_docx.py 中通过 python-docx 的 `para.style = 'Heading 1'` 间接完成的工作。
这里显式做 styleId 解析，因为源文档常来自 WPS/其它工具，
样式名与 styleId 的对应关系完全不可假设（实测同一份文档里
标题样式的 styleId 可能是 '1'、'213'、'afffffff0' 之类）。
"""

from __future__ import annotations

from dataclasses import dataclass

from lxml import etree

from app.core.constants import pt_to_half_points
from app.core.namespaces import PART_STYLES, qn, w_attr
from app.domain.ooxml.package import DocxPackage


@dataclass(frozen=True)
class StyleEntry:
    """一条样式定义。"""

    style_id: str
    name: str
    style_type: str


class StyleResolver:
    """在 styles.xml 中按名称或 ID 查找样式。"""

    def __init__(self, pkg: DocxPackage) -> None:
        self.pkg = pkg
        self._by_id: dict[str, StyleEntry] = {}
        self._by_name: dict[str, StyleEntry] = {}
        self._load()

    def _load(self) -> None:
        if not self.pkg.has(PART_STYLES):
            return

        root = self.pkg.xml(PART_STYLES)
        for style in root.findall(qn("w:style")):
            style_id = style.get(w_attr("styleId")) or ""
            if not style_id:
                continue
            name_el = style.find(qn("w:name"))
            name = (name_el.get(w_attr("val")) if name_el is not None else "") or ""
            entry = StyleEntry(
                style_id=style_id,
                name=name,
                style_type=style.get(w_attr("type")) or "",
            )
            self._by_id[style_id] = entry
            if name:
                self._by_name[name.lower()] = entry

    # ------------------------------------------------------------------
    @property
    def entries(self) -> list[StyleEntry]:
        return list(self._by_id.values())

    def id_of(self, name: str) -> str | None:
        entry = self._by_name.get(name.lower())
        return entry.style_id if entry else None

    def name_of(self, style_id: str) -> str:
        entry = self._by_id.get(style_id)
        return entry.name if entry else ""

    def is_heading(self, style_id: str) -> bool:
        """判断某个 styleId 是否为标题样式。

        判据用"样式名以 heading 开头"而非固定 ID，
        因为中文 Word 的用户界面名是"标题 1"，但底层 name 仍是 "heading 1"。
        """
        name = self.name_of(style_id).lower().replace("_", " ").strip()
        return name.startswith("heading ") and name.split(" ", 1)[1].strip().isdigit()

    def heading_level(self, style_id: str) -> int | None:
        name = self.name_of(style_id).lower().replace("_", " ").strip()
        if not name.startswith("heading "):
            return None
        suffix = name.split(" ", 1)[1].strip()
        return int(suffix) if suffix.isdigit() else None

    def heading_style_id(self, level: int) -> str | None:
        """取指定级别的标题 styleId，缺失时返回 None（不臆造）。"""
        return self.id_of(f"heading {level}")


class StyleProcessor:
    """样式应用与清理。"""

    def __init__(self, pkg: DocxPackage) -> None:
        self.pkg = pkg
        self.resolver = StyleResolver(pkg)

    # ------------------------------------------------------------------
    def set_paragraph_style(
        self,
        paragraph: etree._Element,
        style_id: str,
        *,
        space_before_pt: float | None = None,
        space_after_pt: float | None = None,
    ) -> bool:
        """设置段落样式，可选同时设置段前段后间距。"""
        ppr = paragraph.find(qn("w:pPr"))
        if ppr is None:
            ppr = etree.Element(qn("w:pPr"))
            paragraph.insert(0, ppr)

        pstyle = ppr.find(qn("w:pStyle"))
        if pstyle is None:
            pstyle = etree.Element(qn("w:pStyle"))
            # pStyle 必须是 pPr 的第一个子元素，否则部分版本 Word 会忽略
            ppr.insert(0, pstyle)
        pstyle.set(w_attr("val"), style_id)

        if space_before_pt is not None or space_after_pt is not None:
            spacing = ppr.find(qn("w:spacing"))
            if spacing is None:
                spacing = etree.SubElement(ppr, qn("w:spacing"))
            if space_before_pt is not None:
                spacing.set(w_attr("before"), str(int(space_before_pt * 20)))
            if space_after_pt is not None:
                spacing.set(w_attr("after"), str(int(space_after_pt * 20)))

        return True

    def paragraph_style_id(self, paragraph: etree._Element) -> str:
        pstyle = paragraph.find(f"{qn('w:pPr')}/{qn('w:pStyle')}")
        return (pstyle.get(w_attr("val")) or "") if pstyle is not None else ""

    def adopt_named_styles(self, source: DocxPackage, names: list[str]) -> int:
        """用源文档里同名样式的定义覆盖目标样式，保留目标原有 styleId。

        段落上的 pStyle 继续指向原来的 styleId，显示效果改成模板的字体、
        字号和段距。同名样式有多份时全部覆盖，避免封面和正文各持一份旧定义。
        """
        if not source.has(PART_STYLES) or not self.pkg.has(PART_STYLES):
            return 0

        wanted = {name.lower() for name in names}
        src_root = source.xml(PART_STYLES)
        dst_root = self.pkg.xml(PART_STYLES)
        adopted = 0

        for style in src_root.findall(qn("w:style")):
            name = _style_name(style)
            if name.lower() not in wanted:
                continue

            matches = [item for item in dst_root.findall(qn("w:style")) if _style_name(item).lower() == name.lower()]
            if not matches:
                clone = etree.fromstring(etree.tostring(style))
                dst_root.append(clone)
                adopted += 1
                continue

            for existing in matches:
                clone = etree.fromstring(etree.tostring(style))
                style_id = existing.get(w_attr("styleId"))
                if style_id:
                    clone.set(w_attr("styleId"), style_id)
                parent = existing.getparent()
                if parent is not None:
                    parent.replace(existing, clone)
                    adopted += 1

        if adopted:
            self.resolver = StyleResolver(self.pkg)
        return adopted

    def apply_style_font(
        self,
        style_names: list[str],
        *,
        east_asia: str,
        ascii_font: str,
        size_pt: float,
        bold: bool = False,
    ) -> int:
        """按样式名写入字体。用于目录样式：中文宋体，西文 Times New Roman。"""
        if not self.pkg.has(PART_STYLES):
            return 0

        wanted = {name.lower() for name in style_names}
        root = self.pkg.xml(PART_STYLES)
        touched = 0
        for style in root.findall(qn("w:style")):
            if _style_name(style).lower() not in wanted:
                continue
            _write_style_font(style, east_asia, ascii_font, size_pt, bold)
            touched += 1
        return touched

    @staticmethod
    def clear_direct_run_font(paragraph: etree._Element) -> None:
        """去掉段落 run 上的直接字体，让段落样式里的字体生效。"""
        for run in paragraph.findall(qn("w:r")):
            rpr = run.find(qn("w:rPr"))
            if rpr is None:
                continue
            for tag in ("w:rFonts", "w:sz", "w:szCs"):
                node = rpr.find(qn(tag))
                if node is not None:
                    rpr.remove(node)

    def clear_paragraph_style(self, paragraph: etree._Element) -> None:
        ppr = paragraph.find(qn("w:pPr"))
        if ppr is None:
            return
        pstyle = ppr.find(qn("w:pStyle"))
        if pstyle is not None:
            ppr.remove(pstyle)

    # ------------------------------------------------------------------
    def usage(self) -> dict[str, int]:
        """统计正文样式使用情况，用于回归报告。"""
        counts: dict[str, int] = {}
        for paragraph in self.pkg.paragraphs():
            style_id = self.paragraph_style_id(paragraph)
            if style_id:
                counts[style_id] = counts.get(style_id, 0) + 1
        return counts


def _style_name(style: etree._Element) -> str:
    name_el = style.find(qn("w:name"))
    if name_el is None:
        return ""
    return name_el.get(w_attr("val")) or ""


def _write_style_font(
    style: etree._Element,
    east_asia: str,
    ascii_font: str,
    size_pt: float,
    bold: bool,
) -> None:
    rpr = style.find(qn("w:rPr"))
    if rpr is None:
        rpr = etree.SubElement(style, qn("w:rPr"))

    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = etree.Element(qn("w:rFonts"))
        rpr.insert(0, fonts)
    fonts.set(w_attr("ascii"), ascii_font)
    fonts.set(w_attr("hAnsi"), ascii_font)
    fonts.set(w_attr("eastAsia"), east_asia)

    size = str(pt_to_half_points(size_pt))
    for tag in ("w:sz", "w:szCs"):
        node = rpr.find(qn(tag))
        if node is None:
            node = etree.SubElement(rpr, qn(tag))
        node.set(w_attr("val"), size)

    for tag in ("w:b", "w:bCs"):
        node = rpr.find(qn(tag))
        if bold and node is None:
            etree.SubElement(rpr, qn(tag))
        if not bold and node is not None:
            rpr.remove(node)
