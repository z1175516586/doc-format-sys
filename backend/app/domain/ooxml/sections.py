"""节属性与页面方向处理。

对应原 utils/fix_section_orientation.py 与 document_orientation_processor.py。
原实现基于正则与字符串替换，且横竖拆分严重依赖 Windows COM（InsertBreak/InsertFile）。
此版本只做纯 OOXML 可完成的部分，COM 相关流程由 converters 层承担。
"""

from __future__ import annotations

from dataclasses import dataclass

from lxml import etree

from app.core.logging import get_logger
from app.core.namespaces import qn, w_attr
from app.domain.ooxml.package import DocxPackage

logger = get_logger(__name__)


@dataclass(frozen=True)
class SectionInfo:
    """单个节的页面信息。"""

    index: int
    width: int
    height: int
    orientation: str
    is_landscape: bool


class SectionProcessor:
    """页面方向归一化与统计。"""

    def __init__(self, pkg: DocxPackage) -> None:
        self.pkg = pkg

    # ------------------------------------------------------------------
    # 统计
    # ------------------------------------------------------------------
    def describe(self) -> list[SectionInfo]:
        infos: list[SectionInfo] = []
        for idx, sect_pr in enumerate(self.pkg.section_properties(), start=1):
            pg_sz = sect_pr.find(qn("w:pgSz"))
            if pg_sz is None:
                infos.append(SectionInfo(idx, 0, 0, "unknown", False))
                continue

            width = self._to_int(pg_sz.get(w_attr("w")))
            height = self._to_int(pg_sz.get(w_attr("h")))
            orient = pg_sz.get(w_attr("orient")) or ("landscape" if width > height else "portrait")
            infos.append(
                SectionInfo(
                    index=idx,
                    width=width,
                    height=height,
                    orientation=orient,
                    is_landscape=width > height,
                )
            )
        return infos

    @staticmethod
    def _to_int(value: str | None) -> int:
        try:
            return int(value) if value else 0
        except (TypeError, ValueError):
            return 0

    # ------------------------------------------------------------------
    # 归一化
    # ------------------------------------------------------------------
    def normalize_orientation(self) -> int:
        """依据页面实际宽高修正 w:orient 与宽高值。

        规则：宽 > 高 → landscape；否则 portrait。
        当 orient 与宽高互相矛盾时（Word 中常见于手工改过页面设置的文档），
        以宽高为准并交换数值，避免"看起来横版实际竖版"。
        返回被修正的节数量。
        """
        fixed = 0
        for idx, sect_pr in enumerate(self.pkg.section_properties(), start=1):
            pg_sz = sect_pr.find(qn("w:pgSz"))
            if pg_sz is None:
                continue

            width = self._to_int(pg_sz.get(w_attr("w")))
            height = self._to_int(pg_sz.get(w_attr("h")))
            if width <= 0 or height <= 0:
                continue

            desired = "landscape" if width > height else "portrait"
            current = pg_sz.get(w_attr("orient"))

            if current == desired and (desired == "portrait" or width > height):
                # 已一致；竖版无 orient 属性也是合法的，保持原样
                continue

            pg_sz.set(w_attr("orient"), desired)
            if desired == "portrait" and width > height:
                pg_sz.set(w_attr("w"), str(height))
                pg_sz.set(w_attr("h"), str(width))
            fixed += 1
            logger.debug("第 %d 节方向修正为 %s", idx, desired)

        if fixed:
            logger.info("修正 %d 个节的页面方向", fixed)
        return fixed

    # ------------------------------------------------------------------
    # 边界查询
    # ------------------------------------------------------------------
    @property
    def has_landscape(self) -> bool:
        return any(info.is_landscape for info in self.describe())

    def orientation_sequence(self) -> list[str]:
        """返回按文档顺序的页面方向序列，如 ['portrait', 'landscape']。"""
        return [info.orientation for info in self.describe()]

    # ------------------------------------------------------------------
    # 节属性复制
    # ------------------------------------------------------------------
    def clone_page_setup(self, source_index: int, target_index: int) -> bool:
        """把某一节的页面设置（pgSz/pgMar）复制到另一节。

        用于文档合并后统一横／竖版页面参数。索引从 1 开始。
        """
        sects = self.pkg.section_properties()
        if not (0 < source_index <= len(sects) and 0 < target_index <= len(sects)):
            return False

        source = sects[source_index - 1]
        target = sects[target_index - 1]

        for tag in ("w:pgSz", "w:pgMar"):
            src_el = source.find(qn(tag))
            if src_el is None:
                continue
            old = target.find(qn(tag))
            if old is not None:
                target.remove(old)
            target.append(etree.fromstring(etree.tostring(src_el)))

        return True

    # ------------------------------------------------------------------
    # 插入分节符
    # ------------------------------------------------------------------
    def insert_break_after(self, anchor: etree._Element) -> bool:
        """在指定元素之后插入分节符，继承当前节页面设置。

        对应原 service/converter.py 的 step8_insert_section_break，但原实现依赖
        COM 的 InsertBreak。这里直接在 anchor 后插入一个携带 sectPr 的空段落，
        语义等价：Word 会在该段落处强制分节。

        典型用途：让目录独占一节，从而只给正文编页码。
        """
        governing = self._governing_section_properties(anchor)
        if governing is None:
            logger.warning("无法定位当前节属性，跳过分节符插入")
            return False

        paragraph = etree.Element(qn("w:p"))
        ppr = etree.SubElement(paragraph, qn("w:pPr"))
        ppr.append(etree.fromstring(etree.tostring(governing)))

        anchor.addnext(paragraph)
        logger.info("已在指定元素后插入分节符")
        return True

    def _governing_section_properties(self, element: etree._Element) -> etree._Element | None:
        """找到控制该元素的节属性。

        OOXML 中段落级 sectPr 表示"该段落是某一节的最后一节"，
        因此控制某元素页面设置的，是它之后的第一个 sectPr（含 body 末尾那个）。
        """
        body = self.pkg.body
        children = list(body)
        try:
            start = children.index(element)
        except ValueError:
            start = 0

        for child in children[start:]:
            if child.tag == qn("w:sectPr"):
                return child
            ppr = child.find(qn("w:pPr"))
            if ppr is not None:
                sect = ppr.find(qn("w:sectPr"))
                if sect is not None:
                    return sect
        return None
