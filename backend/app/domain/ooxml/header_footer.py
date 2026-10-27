"""页眉页脚处理。

对应原 utils/del_footer.py、utils/remove_header_image.py、utils/docx_header_footer_replace.py。
原实现直接清空 footer 部件全部子元素，会连带删掉排版必需的结构；
这里只清空内容层（w:p / w:tbl），保留部件根节点。
"""

from __future__ import annotations

from lxml import etree

from app.core.logging import get_logger
from app.core.namespaces import qn
from app.domain.ooxml.package import DocxPackage

logger = get_logger(__name__)

_HEADER_PREFIX = "word/header"
_FOOTER_PREFIX = "word/footer"
_IMAGE_TAGS = {qn("w:drawing"), qn("w:pict"), qn("w:object")}


class HeaderFooterProcessor:
    """页眉页脚内容清理与替换。"""

    def __init__(self, pkg: DocxPackage) -> None:
        self.pkg = pkg

    # ------------------------------------------------------------------
    # 页脚
    # ------------------------------------------------------------------
    def clear_all_footers(self) -> int:
        """清空所有页脚内容，返回处理的部件数量。

        页脚常带页码、密级标记等，与目标模板冲突，因此排版前统一清空，
        后续由目录/正文步骤重新生成。
        """
        count = 0
        for name in self.pkg.iter_parts(prefix=_FOOTER_PREFIX, suffix=".xml"):
            root = self.pkg.xml(name)
            self._clear_content(root)
            count += 1

        if count:
            logger.info("清空 %d 个页脚部件", count)
        return count

    # ------------------------------------------------------------------
    # 页眉
    # ------------------------------------------------------------------
    def remove_header_images(self) -> int:
        """移除页眉中的图片（旧 logo），返回移除数量。"""
        removed = 0
        for name in self.pkg.iter_parts(prefix=_HEADER_PREFIX, suffix=".xml"):
            root = self.pkg.xml(name)
            for parent in root.iter():
                for child in list(parent):
                    if child.tag in _IMAGE_TAGS:
                        parent.remove(child)
                        removed += 1

        if removed:
            logger.info("移除页眉图片 %d 处", removed)
        return removed

    def header_texts(self) -> list[str]:
        """读取所有页眉文本，用于校验替换结果。"""
        texts: list[str] = []
        for name in self.pkg.iter_parts(prefix=_HEADER_PREFIX, suffix=".xml"):
            root = self.pkg.xml(name)
            for node in root.iter(qn("w:t")):
                if node.text and node.text.strip():
                    texts.append(node.text.strip())
        return texts

    # ------------------------------------------------------------------
    # 内部
    # ------------------------------------------------------------------
    @staticmethod
    def _clear_content(root: etree._Element) -> None:
        """删除内容层元素，保留根节点与 w:sectPr 之类的结构节点。"""
        for child in list(root):
            if child.tag in (qn("w:p"), qn("w:tbl")):
                root.remove(child)
