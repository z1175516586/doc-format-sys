"""在 docx 中写入、读取转换树节点身份。

用 Custom XML 部件保存 node_id。Word 另存时通常会保留 customXml，
比改 settings.xml 更稳。不认识的文档返回 None，不抛错。
"""

from __future__ import annotations

from pathlib import Path

from lxml import etree

from app.core.constants import lineage_config
from app.core.namespaces import PART_CONTENT_TYPES, qn
from app.domain.ooxml.package import DocxPackage
from app.core.errors import DocumentCorruptedError


class LineageStamp:
    """文档身份戳：一个类，只负责 node_id 的读写。"""

    def __init__(self) -> None:
        config = lineage_config()
        self.part_name = str(config["part_name"])
        self.props_name = str(config["props_name"])
        self.rels_name = str(config["rels_name"])
        self.namespace = str(config["namespace"])
        self.datastore_ns = str(config["datastore_ns"])
        self.item_guid = str(config["item_guid"])
        self.content_type = str(config["content_type"])
        self.props_content_type = str(config["props_content_type"])

    def read(self, path: str | Path) -> str | None:
        """从文档里取出 node_id。不是 docx 或没有身份时返回 None。"""
        try:
            with DocxPackage.open(path) as pkg:
                return self.read_package(pkg)
        except DocumentCorruptedError:
            return None

    def read_package(self, pkg: DocxPackage) -> str | None:
        if pkg.has(self.part_name):
            node_id = self._parse_part(pkg.read(self.part_name))
            if node_id:
                return node_id
        for name in pkg.iter_parts(prefix="customXml/", suffix=".xml"):
            node_id = self._parse_part(pkg.read(name))
            if node_id:
                return node_id
        return None

    def write(self, path: str | Path, node_id: str) -> None:
        """把 node_id 写进文档。已有身份则覆盖。"""
        with DocxPackage.open(path) as pkg:
            self.write_package(pkg, node_id)
            pkg.save(path)

    def write_package(self, pkg: DocxPackage, node_id: str) -> None:
        pkg.add_part(self.part_name, self._lineage_xml(node_id))
        pkg.add_part(self.props_name, self._props_xml())
        pkg.add_part(self.rels_name, self._rels_xml())
        self._ensure_override(pkg, self.part_name, self.content_type)
        self._ensure_override(pkg, self.props_name, self.props_content_type)

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

    def _lineage_xml(self, node_id: str) -> bytes:
        root = etree.Element(f"{{{self.namespace}}}lineage", nsmap={None: self.namespace})
        child = etree.SubElement(root, f"{{{self.namespace}}}nodeId")
        child.text = node_id
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    def _props_xml(self) -> bytes:
        root = etree.Element(
            f"{{{self.datastore_ns}}}datastoreItem",
            nsmap={"ds": self.datastore_ns},
        )
        root.set(f"{{{self.datastore_ns}}}itemID", self.item_guid)
        refs = etree.SubElement(root, f"{{{self.datastore_ns}}}schemaRefs")
        ref = etree.SubElement(refs, f"{{{self.datastore_ns}}}schemaRef")
        ref.set(f"{{{self.datastore_ns}}}uri", self.namespace)
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    def _rels_xml(self) -> bytes:
        target = Path(self.props_name).name
        root = etree.Element(
            qn("pr:Relationships"),
            nsmap={None: "http://schemas.openxmlformats.org/package/2006/relationships"},
        )
        rel = etree.SubElement(root, qn("pr:Relationship"))
        rel.set("Id", "rId1")
        rel.set(
            "Type",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps",
        )
        rel.set("Target", target)
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    @staticmethod
    def _ensure_override(pkg: DocxPackage, part_name: str, content_type: str) -> None:
        if not pkg.has(PART_CONTENT_TYPES):
            return
        root = pkg.xml(PART_CONTENT_TYPES)
        part = "/" + part_name
        for override in root.findall(qn("ct:Override")):
            if override.get("PartName") == part:
                override.set("ContentType", content_type)
                pkg.mark_dirty(PART_CONTENT_TYPES)
                return
        node = etree.SubElement(root, qn("ct:Override"))
        node.set("PartName", part)
        node.set("ContentType", content_type)
        pkg.mark_dirty(PART_CONTENT_TYPES)
