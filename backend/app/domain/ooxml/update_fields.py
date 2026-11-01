"""打开时更新域（``w:updateFields``）。

目录页码有两种来源，本模块负责第二种：

1. 服务器端排版引擎预刷新（LibreOffice）—— 结果立刻可见，但引擎是全局
   单点，并发时容易崩（soffice.bin 异常），必须串行排队。
2. 让 Word / WPS 打开文档时自己更新域 —— 页码由客户端排版引擎计算，最准，
   且服务器端完全不占用 LibreOffice。

实测（Word 2016+，COM 自动化）：把目录页码人为改成 999 的文档

- 不加 ``w:updateFields``：打开后仍是 999，Word 不会自己更新；
- 加上 ``w:updateFields``：打开后自动被纠正为真实页码，条目齐全。

所以只要写这个标志，就不必为了页码去抢 LibreOffice。
"""

from __future__ import annotations

import zipfile
from pathlib import Path

from lxml import etree

from app.core.logging import get_logger
from app.core.namespaces import PART_SETTINGS, qn, w_attr
from app.domain.ooxml.package import DocxPackage

logger = get_logger(__name__)

# CT_Settings 是 xsd:sequence。这些元素按 schema 必须排在 w:updateFields 之后，
# 插错位置 Word 会判定 settings.xml 需要修复。
_AFTER_UPDATE_FIELDS: tuple[str, ...] = (
    "hdrShapeDefaults",
    "footnotePr",
    "endnotePr",
    "compat",
    "docVars",
    "rsids",
    "mathPr",
    "attachedSchema",
    "themeFontLang",
    "clrSchemeMapping",
    "doNotIncludeSubdocsInStats",
    "doNotAutoCompressPictures",
    "forceUpgrade",
    "captions",
    "readModeInkLockDown",
    "smartTagType",
    "schemaLibrary",
    "shapeDefaults",
    "doNotEmbedSmartTags",
    "decimalSymbol",
    "listSeparator",
)


class UpdateFieldsFlag:
    """``w:updateFields`` 的读写。单一实体，只负责这一个开关。"""

    def __init__(self, part_name: str = PART_SETTINGS) -> None:
        self.part_name = part_name

    # ------------------------------------------------------------------
    def enable(self, pkg: DocxPackage) -> bool:
        """置为打开时更新域。返回是否发生了修改。"""
        if not pkg.has(self.part_name):
            logger.warning("缺少 %s，无法写入 updateFields", self.part_name)
            return False

        root = pkg.xml(self.part_name)
        node = root.find(qn("w:updateFields"))
        if node is None:
            node = etree.Element(qn("w:updateFields"))
            self._insert_ordered(root, node)
            changed = True
        else:
            changed = (node.get(w_attr("val")) or "").lower() != "true"

        node.set(w_attr("val"), "true")
        if changed:
            logger.info("已设置打开文档时更新域（updateFields）")
        return changed

    def disable(self, pkg: DocxPackage) -> bool:
        """移除该开关。返回是否发生了修改。"""
        if not pkg.has(self.part_name):
            return False
        root = pkg.xml(self.part_name)
        node = root.find(qn("w:updateFields"))
        if node is None:
            return False
        root.remove(node)
        return True

    def is_enabled(self, pkg: DocxPackage) -> bool:
        if not pkg.has(self.part_name):
            return False
        node = pkg.xml(self.part_name).find(qn("w:updateFields"))
        if node is None:
            return False
        return (node.get(w_attr("val")) or "").lower() in ("true", "1", "on")

    # ------------------------------------------------------------------
    @staticmethod
    def _insert_ordered(root: etree._Element, node: etree._Element) -> None:
        """按 CT_Settings 的 schema 顺序插入，避免 Word 报结构非法。"""
        for child in root:
            try:
                local = etree.QName(child).localname
            except ValueError:
                continue  # 注释 / 处理指令
            if local in _AFTER_UPDATE_FIELDS:
                child.addprevious(node)
                return
        root.append(node)


def enable_update_fields(path: Path) -> bool:
    """对落盘文件设置打开时更新域。"""
    if not path.is_file():
        return False
    with DocxPackage.open(path) as pkg:
        if not UpdateFieldsFlag().enable(pkg):
            return False
        pkg.save(path)
    return True


def settings_part_names(path: Path) -> list[str]:
    """调试用：列出 settings.xml 的顶层元素顺序。"""
    with zipfile.ZipFile(path) as archive:
        root = etree.fromstring(archive.read(PART_SETTINGS))
    return [etree.QName(child).localname for child in root]
