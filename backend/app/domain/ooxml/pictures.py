"""图片与图形对象处理。

对应原 utils/docx_picture.py、utils/rm_pic_ind.py。
原实现依赖 `winerror import NOERROR`（Windows 专有），在 Linux 直接 ImportError，
这是旧代码无法跨平台的典型例子。
"""

from __future__ import annotations

from copy import deepcopy

from lxml import etree

from app.core.constants import picture_settings
from app.core.logging import get_logger
from app.core.namespaces import PART_STYLES, qn, w_attr
from app.domain.ooxml.package import DocxPackage

logger = get_logger(__name__)

_DRAWING_TAGS = (qn("w:drawing"), qn("w:pict"), qn("w:object"))


def _merge_properties(dest: etree._Element, source: etree._Element, *, override: bool) -> None:
    """合并子节点。override 为真时源覆盖目标，否则只补目标还没有的。"""
    for child in source:
        if child.tag == qn("w:rPr"):
            continue
        existing = dest.find(child.tag)
        if existing is None:
            dest.append(deepcopy(child))
            continue
        for attr, value in child.attrib.items():
            if override or existing.get(attr) is None:
                existing.set(attr, value)


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

            # 普通图片仍改成单倍行距。只有已经用固定行距撑满的 VML 图组才保持原行高，
            # 否则整张图会被裁成一行。
            spacing = ppr.find(qn("w:spacing"))
            if not self._diagram_reserves_line(paragraph):
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

    def preserve_diagram_textbox_source_style(self) -> int:
        """把源文档样式写成文本框的直接格式，框高不改。

        文本框里已经写明的字号、字体保持不动。空行没有自己的字号，行高跟着
        段落标记走。必须在模板样式替换 Normal 之前调用，否则空行会变成模板的
        四号字，几行回车就把最后一行顶出原框。
        """
        if not self.pkg.has(PART_STYLES):
            return 0
        styles = self._paragraph_style_index()
        kept = 0
        for textbox in self.pkg.body.iter(qn("v:textbox")):
            content = textbox.find(qn("w:txbxContent"))
            if content is None:
                continue
            for paragraph in content.findall(qn("w:p")):
                style_ppr, style_rpr = self._effective_paragraph_formatting(paragraph, styles)
                ppr = self._ensure_ppr(paragraph)
                if style_ppr is not None:
                    _merge_properties(ppr, style_ppr, override=False)
                if style_rpr is not None:
                    mark = ppr.find(qn("w:rPr"))
                    if mark is None:
                        mark = etree.SubElement(ppr, qn("w:rPr"))
                    _merge_properties(mark, style_rpr, override=False)
                    for run in paragraph.findall(qn("w:r")):
                        run_rpr = run.find(qn("w:rPr"))
                        if run_rpr is None:
                            run_rpr = etree.Element(qn("w:rPr"))
                            run.insert(0, run_rpr)
                        _merge_properties(run_rpr, style_rpr, override=False)
                kept += 1
        if kept:
            logger.info("流程图文本框已保留源文档样式 %d 段", kept)
        return kept

    # ------------------------------------------------------------------
    # 内部
    # ------------------------------------------------------------------
    def _paragraph_style_index(self) -> dict[str, etree._Element]:
        styles = self.pkg.xml(PART_STYLES)
        return {
            style_id: style
            for style in styles.findall(qn("w:style"))
            if (style_id := style.get(w_attr("styleId")))
            and style.get(w_attr("type")) in (None, "paragraph")
        }

    def _effective_paragraph_formatting(
        self,
        paragraph: etree._Element,
        styles: dict[str, etree._Element],
    ) -> tuple[etree._Element | None, etree._Element | None]:
        """按文档默认值、样式继承链，得到这段在源文档里实际使用的段落和字符格式。"""
        ppr = etree.Element(qn("w:pPr"))
        rpr = etree.Element(qn("w:rPr"))
        self._apply_doc_defaults(ppr, rpr)

        style_id = self._paragraph_style_id(paragraph, styles)
        chain: list[etree._Element] = []
        seen: set[str] = set()
        while style_id and style_id not in seen and style_id in styles:
            seen.add(style_id)
            chain.append(styles[style_id])
            based = styles[style_id].find(qn("w:basedOn"))
            style_id = based.get(w_attr("val")) if based is not None else None
        for style in reversed(chain):
            style_ppr = style.find(qn("w:pPr"))
            if style_ppr is not None:
                _merge_properties(ppr, style_ppr, override=True)
                nested = style_ppr.find(qn("w:rPr"))
                if nested is not None:
                    _merge_properties(rpr, nested, override=True)
            direct = style.find(qn("w:rPr"))
            if direct is not None:
                _merge_properties(rpr, direct, override=True)
        return ppr, rpr

    def _apply_doc_defaults(self, ppr: etree._Element, rpr: etree._Element) -> None:
        styles = self.pkg.xml(PART_STYLES)
        defaults = styles.find(qn("w:docDefaults"))
        if defaults is None:
            return
        paragraph_defaults = defaults.find(qn("w:pPrDefault"))
        if paragraph_defaults is not None:
            default_ppr = paragraph_defaults.find(qn("w:pPr"))
            if default_ppr is not None:
                _merge_properties(ppr, default_ppr, override=True)
        run_defaults = defaults.find(qn("w:rPrDefault"))
        if run_defaults is not None:
            default_rpr = run_defaults.find(qn("w:rPr"))
            if default_rpr is not None:
                _merge_properties(rpr, default_rpr, override=True)

    def _paragraph_style_id(self, paragraph: etree._Element, styles: dict[str, etree._Element]) -> str | None:
        ppr = paragraph.find(qn("w:pPr"))
        if ppr is not None:
            pstyle = ppr.find(qn("w:pStyle"))
            style_id = pstyle.get(w_attr("val")) if pstyle is not None else None
            if style_id:
                return style_id
        for style_id, style in styles.items():
            if style.get(w_attr("default")) in ("1", "true", "on"):
                return style_id
        for style_id, style in styles.items():
            name = style.find(qn("w:name"))
            if name is not None and (name.get(w_attr("val")) or "").lower() == "normal":
                return style_id
        return None

    @staticmethod
    def _diagram_reserves_line(paragraph: etree._Element) -> bool:
        spacing = paragraph.find(f"{qn('w:pPr')}/{qn('w:spacing')}")
        if spacing is None or (spacing.get(w_attr("lineRule")) or "").lower() != "exact":
            return False
        return any("height:" in (group.get("style") or "") for group in paragraph.iter(qn("v:group")))

    @staticmethod
    def _ensure_ppr(paragraph: etree._Element) -> etree._Element:
        ppr = paragraph.find(qn("w:pPr"))
        if ppr is None:
            ppr = etree.Element(qn("w:pPr"))
            paragraph.insert(0, ppr)
        return ppr
