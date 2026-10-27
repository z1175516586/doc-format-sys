"""图片与图形对象处理。

对应原 utils/docx_picture.py、utils/rm_pic_ind.py。
原实现依赖 `winerror import NOERROR`（Windows 专有），在 Linux 直接 ImportError，
这是旧代码无法跨平台的典型例子。
"""

from __future__ import annotations

from lxml import etree

from app.core.constants import picture_settings
from app.core.logging import get_logger
from app.core.namespaces import qn, w_attr
from app.domain.ooxml.package import DocxPackage

logger = get_logger(__name__)

_DRAWING_TAGS = (qn("w:drawing"), qn("w:pict"), qn("w:object"))


class PictureProcessor:
    """正文图片段落格式归一化（页眉页脚图片不处理）。"""

    def __init__(self, pkg: DocxPackage) -> None:
        self.pkg = pkg

    # ------------------------------------------------------------------
    # 查询
    # ------------------------------------------------------------------
    @staticmethod
    def has_picture(paragraph: etree._Element) -> bool:
        return any(paragraph.find(f".//{tag}") is not None for tag in _DRAWING_TAGS)

    def picture_paragraphs(self) -> list[etree._Element]:
        return [p for p in self.pkg.paragraphs() if self.has_picture(p)]

    def count(self) -> int:
        return len(self.picture_paragraphs())

    # ------------------------------------------------------------------
    # 处理
    # ------------------------------------------------------------------
    def reset_indent(self) -> int:
        """清除图片段落缩进，返回处理数量。

        图片段落若继承正文首行缩进，会导致图片偏右、跨页异常。
        """
        count = 0
        for paragraph in self.picture_paragraphs():
            ppr = self._ensure_ppr(paragraph)
            ind = ppr.find(qn("w:ind"))
            if ind is None:
                ind = etree.SubElement(ppr, qn("w:ind"))
            for attr in ("left", "leftChars", "firstLine", "firstLineChars", "right", "rightChars"):
                ind.set(w_attr(attr), "0")
            count += 1

        if count:
            logger.info("重置 %d 个图片段落缩进", count)
        return count

    def apply_alignment(self) -> int:
        """按配置设置图片段落对齐与行距。"""
        config = picture_settings()
        alignment_map = {
            "left": "left",
            "center": "center",
            "right": "right",
            "justify": "both",
        }
        jc_value = alignment_map.get(str(config.get("alignment", "center")), "center")

        count = 0
        for paragraph in self.picture_paragraphs():
            ppr = self._ensure_ppr(paragraph)

            ind = ppr.find(qn("w:ind"))
            if ind is None:
                ind = etree.SubElement(ppr, qn("w:ind"))
            for attr in ("left", "leftChars", "firstLine", "firstLineChars"):
                ind.set(w_attr(attr), "0")

            jc = ppr.find(qn("w:jc"))
            if jc is None:
                jc = etree.SubElement(ppr, qn("w:jc"))
            jc.set(w_attr("val"), jc_value)

            spacing = ppr.find(qn("w:spacing"))
            if spacing is None:
                spacing = etree.SubElement(ppr, qn("w:spacing"))
            line_value = int(round(float(config.get("line_spacing", 1.0)) * 240))
            spacing.set(w_attr("line"), str(line_value))
            spacing.set(w_attr("lineRule"), "auto")
            spacing.set(w_attr("before"), str(int(config.get("before_spacing", 0))))
            spacing.set(w_attr("after"), str(int(config.get("after_spacing", 0))))

            count += 1

        if count:
            logger.info("归一化 %d 个图片段落格式", count)
        return count

    # ------------------------------------------------------------------
    # 内部
    # ------------------------------------------------------------------
    @staticmethod
    def _ensure_ppr(paragraph: etree._Element) -> etree._Element:
        ppr = paragraph.find(qn("w:pPr"))
        if ppr is None:
            ppr = etree.Element(qn("w:pPr"))
            paragraph.insert(0, ppr)
        return ppr
