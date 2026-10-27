"""样式解析与规范化。

对应原 utils/fix_docx.py 中通过 python-docx 的 `para.style = 'Heading 1'` 间接完成的工作。
这里显式做 styleId 解析，因为源文档常来自 WPS/其它工具，
样式名与 styleId 的对应关系完全不可假设（实测同一份文档里
标题样式的 styleId 可能是 '1'、'213'、'afffffff0' 之类）。
"""

from __future__ import annotations

from dataclasses import dataclass

from lxml import etree

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
