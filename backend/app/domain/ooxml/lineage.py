"""在 docx 中写入、读取转换树节点身份。

写入位置是 ``word/settings.xml`` 的 ``w:docVars``——Word 原生「文档变量」。
Word / WPS 打开、编辑、另存时会自己维护这块内容，所以 node_id 不会丢。

早期版本写在 ``customXml/dfsLineage.xml`` 部件里，Word 另存时会把整个部件
清掉，导致 node_id 消失；读取仍保留对旧文档的兜底识别。
"""

from __future__ import annotations

from pathlib import Path

from lxml import etree

from app.core.constants import lineage_config
from app.core.errors import DocumentCorruptedError
from app.core.namespaces import (
    PART_CONTENT_TYPES,
    PART_DOCUMENT_RELS,
    PART_SETTINGS,
    R,
    W,
    qn,
    w_attr,
)
from app.domain.ooxml.package import DocxPackage

# CT_Settings 中排在 w:docVars 之后的元素，按 schema 顺序排列。
# w:docVars 必须插在这些元素之前，否则 Word 判定 settings.xml 结构非法。
# 背景：CT_Settings 是 xsd:sequence，order 错了 Word 会提示需要修复。
_AFTER_DOCVARS: tuple[str, ...] = (
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

# 缺少 w:settings 时补一个最小部件
_MIN_SETTINGS_XML = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    b'<w:settings xmlns:w="'
    + W.encode("utf-8")
    + b'"></w:settings>'
)

_SETTINGS_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
)


class LineageStamp:
    """文档身份戳：一个类，只负责 node_id 的读写。"""

    def __init__(self) -> None:
        config = lineage_config()
        self.docvar_name = str(config.get("docvar_name") or "DocHistoryId")
        self.settings_part = str(config.get("settings_part") or PART_SETTINGS)

        # 以下仅用于读取旧文档（customXml 方案），不再写入
        self.part_name = str(config["part_name"])
        self.props_name = str(config["props_name"])
        self.rels_name = str(config["rels_name"])
        self.namespace = str(config["namespace"])
        self.datastore_ns = str(config["datastore_ns"])
        self.item_guid = str(config["item_guid"])
        self.content_type = str(config["content_type"])
        self.props_content_type = str(config["props_content_type"])

    # ------------------------------------------------------------------
    # 读
    # ------------------------------------------------------------------
    def read(self, path: str | Path) -> str | None:
        """从文档里取出 node_id。不是 docx 或没有身份时返回 None。"""
        try:
            with DocxPackage.open(path) as pkg:
                return self.read_package(pkg)
        except DocumentCorruptedError:
            return None

    def read_package(self, pkg: DocxPackage) -> str | None:
        node_id = self._read_docvar(pkg)
        if node_id:
            return node_id
        return self._read_custom_xml(pkg)

    def _read_docvar(self, pkg: DocxPackage) -> str | None:
        """从 settings.xml 的 w:docVars 读取。这是当前唯一的写入位置。"""
        if not pkg.has(self.settings_part):
            return None
        try:
            root = pkg.xml(self.settings_part)
        except DocumentCorruptedError:
            return None

        doc_vars = root.find(qn("w:docVars"))
        if doc_vars is None:
            return None
        for var in doc_vars.findall(qn("w:docVar")):
            if var.get(w_attr("name")) != self.docvar_name:
                continue
            value = (var.get(w_attr("val")) or "").strip()
            if value:
                return value
        return None

    def _read_custom_xml(self, pkg: DocxPackage) -> str | None:
        """兼容旧文档：识别写在 customXml 部件里的身份。"""
        if pkg.has(self.part_name):
            node_id = self._parse_part(pkg.read(self.part_name))
            if node_id:
                return node_id
        for name in pkg.iter_parts(prefix="customXml/", suffix=".xml"):
            node_id = self._parse_part(pkg.read(name))
            if node_id:
                return node_id
        return None

    # ------------------------------------------------------------------
    # 写
    # ------------------------------------------------------------------
    def write(self, path: str | Path, node_id: str) -> None:
        """把 node_id 写进文档。已有身份则覆盖。"""
        with DocxPackage.open(path) as pkg:
            self.write_package(pkg, node_id)
            pkg.save(path)

    def write_package(self, pkg: DocxPackage, node_id: str) -> None:
        self._ensure_settings_part(pkg)
        self._set_docvar(pkg, node_id)
        # 旧方案留下的部件一并清掉，避免同一份文档出现两个身份
        self._drop_custom_xml(pkg)

    def _set_docvar(self, pkg: DocxPackage, node_id: str) -> None:
        root = pkg.xml(self.settings_part)
        doc_vars = self._ensure_docvars(root)
        for var in doc_vars.findall(qn("w:docVar")):
            if var.get(w_attr("name")) == self.docvar_name:
                var.set(w_attr("val"), node_id)
                return
        var = etree.SubElement(doc_vars, qn("w:docVar"))
        var.set(w_attr("name"), self.docvar_name)
        var.set(w_attr("val"), node_id)

    @staticmethod
    def _ensure_docvars(root: etree._Element) -> etree._Element:
        existing = root.find(qn("w:docVars"))
        if existing is not None:
            return existing

        doc_vars = etree.Element(qn("w:docVars"))
        for child in root:
            try:
                local = etree.QName(child).localname
            except ValueError:
                continue  # 注释 / 处理指令
            if local in _AFTER_DOCVARS:
                child.addprevious(doc_vars)
                return doc_vars
        root.append(doc_vars)
        return doc_vars

    def _ensure_settings_part(self, pkg: DocxPackage) -> None:
        """settings.xml 缺失时补一个，并挂上 rel 与 content type。"""
        if pkg.has(self.settings_part):
            return

        pkg.add_part(self.settings_part, _MIN_SETTINGS_XML)
        self._ensure_override(pkg, self.settings_part, _SETTINGS_CONTENT_TYPE)

        if pkg.has(PART_DOCUMENT_RELS):
            root = pkg.xml(PART_DOCUMENT_RELS)
            rid = self._next_rid(root)
            rel = etree.SubElement(root, qn("pr:Relationship"))
            rel.set("Id", rid)
            rel.set("Type", f"{R}/settings")
            rel.set("Target", Path(self.settings_part).name)

    @staticmethod
    def _next_rid(root: etree._Element) -> str:
        used = {rel.get("Id") for rel in root.findall(qn("pr:Relationship"))}
        index = 1
        while f"rId{index}" in used:
            index += 1
        return f"rId{index}"

    def _drop_custom_xml(self, pkg: DocxPackage) -> None:
        removed = False
        for name in (self.part_name, self.props_name, self.rels_name):
            if pkg.remove_part(name):
                removed = True
        if not removed:
            return
        if pkg.has(PART_CONTENT_TYPES):
            root = pkg.xml(PART_CONTENT_TYPES)
            parts = {"/" + name for name in (self.part_name, self.props_name)}
            for override in root.findall(qn("ct:Override")):
                if override.get("PartName") in parts:
                    root.remove(override)

    # ------------------------------------------------------------------
    # 旧格式解析
    # ------------------------------------------------------------------
    def _parse_part(self, data: bytes) -> str | None:
        try:
            root = etree.fromstring(data)
        except etree.XMLSyntaxError:
            return None
        if root.tag != f"{{{self.namespace}}}lineage":
            return None
        node = root.find(f"{{{self.namespace}}}nodeId")
        if node is None or not (node.text or "").strip():
            return None
        return node.text.strip()

    @staticmethod
    def _ensure_override(pkg: DocxPackage, part_name: str, content_type: str) -> None:
        if not pkg.has(PART_CONTENT_TYPES):
            return
        root = pkg.xml(PART_CONTENT_TYPES)
        part = "/" + part_name
        for override in root.findall(qn("ct:Override")):
            if override.get("PartName") == part:
                override.set("ContentType", content_type)
                return
        node = etree.SubElement(root, qn("ct:Override"))
        node.set("PartName", part)
        node.set("ContentType", content_type)
