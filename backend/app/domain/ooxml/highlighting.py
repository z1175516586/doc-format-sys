"""突出显示与底纹清理。

对应原 utils/docx_remove_highlight.py。
原实现直接执行在 import 时（模块级调用），无法被复用，这里改为类方法。
"""

from __future__ import annotations

from lxml import etree

from app.core.logging import get_logger
from app.core.namespaces import qn, w_attr
from app.domain.ooxml.package import DocxPackage

logger = get_logger(__name__)

_COLOR_AUTO = "auto"


class HighlightProcessor:
    """移除高亮、底纹，并把非自动颜色归一为自动色。"""

    def __init__(self, pkg: DocxPackage) -> None:
        self.pkg = pkg

    def remove_all(self) -> dict[str, int]:
        """清理全部 word/*.xml 部件中的高亮标记。

        返回各类型的处理计数，便于写入任务报告。
        """
        stats = {"highlight": 0, "shading": 0, "color": 0}

        for name in self.pkg.iter_parts(prefix="word/", suffix=".xml"):
            root = self.pkg.xml(name)
            stats["highlight"] += self._remove_tags(root, qn("w:highlight"))
            stats["shading"] += self._remove_tags(root, qn("w:shd"))
            stats["color"] += self._normalize_colors(root)

        if any(stats.values()):
            logger.info("高亮清理: %s", stats)
        return stats

    @staticmethod
    def _remove_tags(root: etree._Element, tag: str) -> int:
        removed = 0
        for node in list(root.iter(tag)):
            parent = node.getparent()
            if parent is not None:
                parent.remove(node)
                removed += 1
        return removed

    @staticmethod
    def _normalize_colors(root: etree._Element) -> int:
        """把非 auto 的字体颜色置为 auto，保留节点以免破坏 run 属性顺序。"""
        changed = 0
        for node in root.iter(qn("w:color")):
            value = (node.get(w_attr("val")) or "").lower()
            if value and value != _COLOR_AUTO:
                node.set(w_attr("val"), _COLOR_AUTO)
                changed += 1
        return changed
