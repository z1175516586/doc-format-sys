"""DOCX 包读写器。

这是整个 OOXML 层的复用基石。原实现中每个功能都重复以下模式：

    zipfile.ZipFile(path).extractall(temp_dir)
    ... 改文件 ...
    os.walk(temp_dir) 重新 zip

由此带来三类问题：临时目录泄漏、部件顺序变化、XML 声明/命名空间被破坏。
DocxPackage 改为在内存中持有部件字典，只在最后一次性原子写回。
"""

from __future__ import annotations

import os
import zipfile
from collections import OrderedDict
from pathlib import Path
from typing import Iterator

from lxml import etree

from app.core.errors import DocumentCorruptedError
from app.core.namespaces import NSMAP, PART_DOCUMENT, qn

_XML_PARSER = etree.XMLParser(
    remove_blank_text=False,
    resolve_entities=False,
    huge_tree=True,
)


def _is_xml_part(name: str) -> bool:
    return name.endswith(".xml") or name.endswith(".rels")


class DocxPackage:
    """一个可读写的 docx（OPC 包）。

    用法::

        with DocxPackage.open(path) as pkg:
            root = pkg.xml(PART_DOCUMENT)
            ...
            pkg.save(output_path)
    """

    def __init__(self, path: Path, parts: "OrderedDict[str, bytes]") -> None:
        self.path = Path(path)
        self._parts = parts
        self._xml_cache: dict[str, etree._Element] = {}
        self._dirty: set[str] = set()

    # ------------------------------------------------------------------
    # 构造
    # ------------------------------------------------------------------
    @classmethod
    def open(cls, path: str | Path) -> "DocxPackage":
        src = Path(path)
        if not src.exists():
            raise DocumentCorruptedError(f"文件不存在: {src}")

        if not zipfile.is_zipfile(src):
            raise DocumentCorruptedError("文件不是合法的 docx（zip 结构异常），可能是 .doc 或已损坏")

        parts: "OrderedDict[str, bytes]" = OrderedDict()
        try:
            with zipfile.ZipFile(src, "r") as zf:
                for item in zf.infolist():
                    if item.is_dir():
                        continue
                    parts[item.filename] = zf.read(item.filename)
        except zipfile.BadZipFile as exc:
            raise DocumentCorruptedError(f"docx 解包失败: {exc}") from exc

        if PART_DOCUMENT not in parts:
            raise DocumentCorruptedError("docx 缺少 word/document.xml，不是有效的 Word 文档")

        return cls(src, parts)

    # ------------------------------------------------------------------
    # 部件访问
    # ------------------------------------------------------------------
    def __contains__(self, name: str) -> bool:
        return name in self._parts

    def has(self, name: str) -> bool:
        return name in self._parts

    @property
    def part_names(self) -> list[str]:
        return list(self._parts.keys())

    def iter_parts(self, prefix: str = "", suffix: str = "") -> Iterator[str]:
        for name in self._parts:
            if name.startswith(prefix) and name.endswith(suffix):
                yield name

    def read(self, name: str) -> bytes:
        try:
            return self._parts[name]
        except KeyError as exc:
            raise DocumentCorruptedError(f"缺少部件: {name}") from exc

    def text(self, name: str) -> str:
        return self.read(name).decode("utf-8", errors="ignore")

    # ------------------------------------------------------------------
    # XML 访问（就地修改）
    # ------------------------------------------------------------------
    def xml(self, name: str) -> etree._Element:
        """解析并缓存部件 XML；对返回值的修改会在 save 时写回。"""
        if name in self._xml_cache:
            return self._xml_cache[name]

        data = self.read(name)
        try:
            root = etree.fromstring(data, parser=_XML_PARSER)
        except etree.XMLSyntaxError as exc:
            raise DocumentCorruptedError(f"部件 {name} XML 解析失败: {exc}") from exc

        self._xml_cache[name] = root
        return root

    def mark_dirty(self, name: str) -> None:
        """显式标记部件需要重新序列化。"""
        if name in self._xml_cache:
            self._dirty.add(name)

    def replace_xml_bytes(self, name: str, data: bytes) -> None:
        """直接替换部件原始字节，同时清掉解析缓存。"""
        self._parts[name] = data
        self._xml_cache.pop(name, None)
        self._dirty.discard(name)

    def add_part(self, name: str, data: bytes) -> None:
        self._parts[name] = data
        self._xml_cache.pop(name, None)
        self._dirty.discard(name)

    def remove_part(self, name: str) -> bool:
        existed = name in self._parts
        self._parts.pop(name, None)
        self._xml_cache.pop(name, None)
        self._dirty.discard(name)
        return existed

    # ------------------------------------------------------------------
    # 序列化
    # ------------------------------------------------------------------
    def _serialize(self, name: str, root: etree._Element) -> bytes:
        return etree.tostring(
            root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

    def _flush(self) -> None:
        """把已解析（可能已修改）的 XML 回写到部件字节。

        save / clone 前必须调用：否则 clone 出来的包会丢掉内存中的修改，
        这是"改了 A 却得到旧 A"这类问题的根源。
        """
        for name, root in self._xml_cache.items():
            self._parts[name] = self._serialize(name, root)

    def save(self, dest: str | Path | None = None) -> Path:
        """原子写回。dest 为 None 时覆盖自身。"""
        self._flush()

        target = Path(dest) if dest is not None else self.path
        target.parent.mkdir(parents=True, exist_ok=True)

        tmp_path = target.with_name(target.name + ".tmp")
        try:
            with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zf:
                for name, data in self._parts.items():
                    zf.writestr(name, data)
            os.replace(tmp_path, target)
        finally:
            if tmp_path.exists():
                tmp_path.unlink(missing_ok=True)

        self.path = target
        return target

    # ------------------------------------------------------------------
    # 快捷查询
    # ------------------------------------------------------------------
    @property
    def body(self) -> etree._Element:
        """返回 w:body 元素。"""
        root = self.xml(PART_DOCUMENT)
        body = root.find(qn("w:body"))
        if body is None:
            raise DocumentCorruptedError("document.xml 缺少 w:body")
        return body

    def paragraphs(self) -> list[etree._Element]:
        """正文顶层段落（不含表格内段落），与 python-docx 的 doc.paragraphs 对齐。"""
        return [el for el in self.body if el.tag == qn("w:p")]

    def tables(self) -> list[etree._Element]:
        """正文顶层表格。"""
        return [el for el in self.body if el.tag == qn("w:tbl")]

    def section_properties(self) -> list[etree._Element]:
        """所有 w:sectPr，按文档顺序。

        注意：节属性有两种存在形式——段落内 w:pPr/w:sectPr 和 body 末尾的 w:sectPr，
        这里用文档顺序 + body 末尾统一返回，调用方无需关心差异。
        """
        root = self.xml(PART_DOCUMENT)
        return root.xpath(".//w:sectPr", namespaces=NSMAP)

    # ------------------------------------------------------------------
    # 上下文管理
    # ------------------------------------------------------------------
    def clone(self) -> "DocxPackage":
        """复制出一个独立的包实例（部件字节深拷贝）。

        用于文档拆分：以同一份源文档为基底生成封面包与正文包，
        互不影响，也避免重复解包磁盘文件。
        """
        self._flush()
        parts: "OrderedDict[str, bytes]" = OrderedDict(
            (name, bytes(data)) for name, data in self._parts.items()
        )
        return DocxPackage(self.path, parts)

    def __enter__(self) -> "DocxPackage":
        return self

    def __exit__(self, exc_type, exc, tb) -> None:  # noqa: ANN001
        self._xml_cache.clear()


def paragraph_text(paragraph: etree._Element) -> str:
    """提取段落纯文本（含所有 run 与嵌套内容）。"""
    return "".join(paragraph.itertext()).strip()


def element_text(element: etree._Element) -> str:
    """提取任意元素的纯文本。"""
    return "".join(element.itertext()).strip()


def make_element(tag: str, **attrs: str) -> etree._Element:
    """创建一个空元素，属性名支持 'w:val' 形式。"""
    el = etree.Element(qn(tag))
    for key, value in attrs.items():
        el.set(qn(key), value)
    return el


def sub_element(parent: etree._Element, tag: str, **attrs: str) -> etree._Element:
    """在父元素下创建子元素，属性名支持 'w:val' 形式。"""
    el = etree.SubElement(parent, qn(tag))
    for key, value in attrs.items():
        el.set(qn(key), value)
    return el
