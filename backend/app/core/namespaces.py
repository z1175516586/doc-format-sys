"""OOXML 命名空间常量。

集中定义，避免各模块重复写长字符串（原代码中散落在 10+ 个文件里）。
"""

from __future__ import annotations

# ---- 核心命名空间 ----
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
V = "urn:schemas-microsoft-com:vml"
O = "urn:schemas-microsoft-com:office:office"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
XML = "http://www.w3.org/XML/1998/namespace"

# OPC 包结构命名空间（.rels / [Content_Types].xml），注意它们的元素没有前缀
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
PKG_CT = "http://schemas.openxmlformats.org/package/2006/content-types"

# ---- lxml 查询用前缀映射 ----
NSMAP: dict[str, str] = {
    "w": W,
    "r": R,
    "wp": WP,
    "a": A,
    "pic": PIC,
    "v": V,
    "o": O,
    "mc": MC,
    "wps": WPS,
    "w14": W14,
    "xml": XML,
    "pr": PKG_REL,
    "ct": PKG_CT,
}

# ---- 包内部件路径 ----
PART_DOCUMENT = "word/document.xml"
PART_STYLES = "word/styles.xml"
PART_NUMBERING = "word/numbering.xml"
PART_SETTINGS = "word/settings.xml"
PART_CONTENT_TYPES = "[Content_Types].xml"
PART_DOCUMENT_RELS = "word/_rels/document.xml.rels"


def qn(tag: str) -> str:
    """把 'w:p' 形式的标签转换为 Clark notation '{ns}p'。

    用于 lxml Element 的 tag/find/set 操作，避免每次手写完整命名空间。
    """
    prefix, _, local = tag.partition(":")
    if not local:
        raise ValueError(f"标签必须带前缀，如 'w:p'，收到: {tag!r}")
    return f"{{{NSMAP[prefix]}}}{local}"


def w_attr(name: str) -> str:
    """返回 w: 命名空间下的属性全名，如 'val' -> '{...}val'。"""
    return f"{{{W}}}{name}"


def r_attr(name: str) -> str:
    """返回 r: 命名空间下的属性全名，用于 rId 引用。"""
    return f"{{{R}}}{name}"
