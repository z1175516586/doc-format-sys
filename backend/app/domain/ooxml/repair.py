"""文档结构修复：非法关系、缺失样式。

原实现位于 utils/fix_docx.py，使用正则处理 XML。这里改为 lxml 结构化处理，
避免正则误伤（例如属性顺序不同、自闭合标签差异）。
"""

from __future__ import annotations

from lxml import etree

from app.core.constants import layout_rules
from app.core.logging import get_logger
from app.core.namespaces import PART_STYLES, qn, w_attr
from app.domain.ooxml.package import DocxPackage

logger = get_logger(__name__)

_INVALID_TARGET = "../NULL"


class DocumentRepairer:
    """针对无法被 python-docx / LibreOffice 正常打开的文档做最小修复。"""

    def __init__(self, pkg: DocxPackage) -> None:
        self.pkg = pkg

    # ------------------------------------------------------------------
    # 非法关系
    # ------------------------------------------------------------------
    def fix_null_relationships(self) -> int:
        """删除 Target="../NULL" 的非法关系。

        这类关系多由跨文档复制粘贴产生。Word 可以容忍，但 python-docx 读取会失败。
        返回删除的关系数量。
        """
        removed = 0
        for name in self.pkg.iter_parts(suffix=".rels"):
            root = self.pkg.xml(name)
            for rel in root.findall(qn("pr:Relationship")):
                if (rel.get("Target") or "").strip() == _INVALID_TARGET:
                    root.remove(rel)
                    removed += 1

        if removed:
            logger.info("清理非法 NULL 关系 %d 条", removed)
        return removed

    # ------------------------------------------------------------------
    # 标题样式
    # ------------------------------------------------------------------
    def ensure_heading_styles(self) -> list[int]:
        """确保 styles.xml 中存在 Heading 1..N 段落样式，返回补齐的级别。"""
        if not self.pkg.has(PART_STYLES):
            logger.warning("文档缺少 styles.xml，跳过标题样式修复")
            return []

        rules = layout_rules()["headings"]
        max_level = int(rules["max_level"])

        root = self.pkg.xml(PART_STYLES)
        existing = self._existing_heading_levels(root)

        default_style_id = self._default_paragraph_style_id(root)
        added: list[int] = []

        for level in range(1, max_level + 1):
            if level in existing:
                continue
            self._append_heading_style(root, level, default_style_id)
            added.append(level)

        if added:
            logger.info("补齐缺失标题样式: %s", added)
        return added

    @staticmethod
    def _existing_heading_levels(styles_root: etree._Element) -> set[int]:
        levels: set[int] = set()
        for style in styles_root.findall(qn("w:style")):
            if style.get(w_attr("type")) != "paragraph":
                continue
            name_el = style.find(qn("w:name"))
            if name_el is None:
                continue
            value = (name_el.get(w_attr("val")) or "").lower().replace("_", " ").strip()
            if not value.startswith("heading "):
                continue
            suffix = value.split(" ", 1)[1].strip()
            if suffix.isdigit():
                levels.add(int(suffix))
        return levels

    @staticmethod
    def _default_paragraph_style_id(styles_root: etree._Element) -> str | None:
        for style in styles_root.findall(qn("w:style")):
            if style.get(w_attr("type")) != "paragraph":
                continue
            if style.get(w_attr("default")) in ("1", "true", "on"):
                return style.get(w_attr("styleId"))
        return None

    def _append_heading_style(
        self,
        styles_root: etree._Element,
        level: int,
        based_on: str | None,
    ) -> None:
        """追加一个最小可用的标题样式。

        styleId 采用 "Heading{level}"，与 OOXML 惯例一致，避免原实现
        "100/101/102/103" 这类与内置样式冲突的取值。
        """
        style = etree.SubElement(styles_root, qn("w:style"))
        style.set(w_attr("type"), "paragraph")
        style.set(w_attr("styleId"), f"Heading{level}")

        name = etree.SubElement(style, qn("w:name"))
        name.set(w_attr("val"), f"heading {level}")

        if based_on:
            based = etree.SubElement(style, qn("w:basedOn"))
            based.set(w_attr("val"), based_on)

        next_style = etree.SubElement(style, qn("w:next"))
        next_style.set(w_attr("val"), based_on or f"Heading{level}")

        etree.SubElement(style, qn("w:qFormat"))

        ppr = etree.SubElement(style, qn("w:pPr"))
        etree.SubElement(ppr, qn("w:keepNext"))
        etree.SubElement(ppr, qn("w:keepLines"))
        outline = etree.SubElement(ppr, qn("w:outlineLvl"))
        outline.set(w_attr("val"), str(level - 1))

    # ------------------------------------------------------------------
    # 标题样式的分页约束
    # ------------------------------------------------------------------
    def remove_heading_keep_options(self) -> int:
        """移除标题样式中的"与下段同页/段中不分页"约束。

        对应原 utils/remove_heading_keep_options.py。
        标题带 keepNext 时，若标题后紧跟跨页表格，会整块被推到下一页，
        造成大片空白，因此排版收尾阶段统一清除。
        """
        if not self.pkg.has(PART_STYLES):
            return 0

        max_level = int(layout_rules()["headings"]["max_level"])
        targets = {f"heading {level}".lower() for level in range(1, max_level + 1)}
        targets |= {f"heading{level}".lower() for level in range(1, max_level + 1)}

        root = self.pkg.xml(PART_STYLES)
        removed = 0

        for style in root.findall(qn("w:style")):
            name_el = style.find(qn("w:name"))
            if name_el is None:
                continue
            name = (name_el.get(w_attr("val")) or "").lower()
            if name not in targets:
                continue

            ppr = style.find(qn("w:pPr"))
            if ppr is None:
                continue
            for tag in ("w:keepNext", "w:keepLines"):
                for node in ppr.findall(qn(tag)):
                    ppr.remove(node)
                    removed += 1

        if removed:
            logger.info("移除标题分页约束 %d 处", removed)
        return removed

    # ------------------------------------------------------------------
    # 组合
    # ------------------------------------------------------------------
    def run_all(self) -> dict[str, object]:
        """执行全部结构修复，返回修复报告。"""
        return {
            "null_relationships_removed": self.fix_null_relationships(),
            "heading_styles_added": self.ensure_heading_styles(),
        }
