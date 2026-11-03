"""自动编号还原的单元测试。

还原口径以 Word 实际渲染为准，用 Word COM 读取 ListString 逐段对照过
《玉麒麟竖炉节能增效改造EPC总承包工程-初步设计说明书》全文（270 段有编号，0 处不一致）。
这里的用例把当时踩到的坑固化成断言。
"""

from __future__ import annotations

from pathlib import Path

import pytest
from docx import Document
from lxml import etree

from app.core.namespaces import PART_NUMBERING, PART_STYLES, qn, w_attr
from app.domain.ooxml.numbering import NumberingProcessor, format_number_value
from app.domain.ooxml.package import DocxPackage, paragraph_text

_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

_NUMBERING_TEMPLATE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering {ns}>
  <w:abstractNum w:abstractNumId="2">
    <w:multiLevelType w:val="multilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1"/><w:suff w:val="space"/>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/><w:suff w:val="space"/>
    </w:lvl>
    <w:lvl w:ilvl="2">
      <w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2.%3"/><w:suff w:val="space"/>
    </w:lvl>
    <w:lvl w:ilvl="3">
      <w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2.%3.%4"/><w:suff w:val="space"/>
      {deep_restart}
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="2"/></w:num>
  <w:num w:numId="7">
    <w:abstractNumId w:val="2"/>
    <w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>
    <w:lvlOverride w:ilvl="1"><w:startOverride w:val="1"/></w:lvlOverride>
    <w:lvlOverride w:ilvl="2"><w:startOverride w:val="1"/></w:lvlOverride>
    <w:lvlOverride w:ilvl="3"><w:startOverride w:val="1"/></w:lvlOverride>
  </w:num>
</w:numbering>
"""


def _numbering_xml(*, deep_restart: int | None = None) -> str:
    tag = "" if deep_restart is None else f'<w:lvlRestart w:val="{deep_restart}"/>'
    return _NUMBERING_TEMPLATE.format(ns=_NS, deep_restart=tag)


def _open_with_numbering(tmp_path: Path, numbering_xml: str) -> DocxPackage:
    path = tmp_path / "num.docx"
    doc = Document()
    doc.add_paragraph("占位")
    doc.save(path)
    pkg = DocxPackage.open(path)
    pkg.add_part(PART_NUMBERING, numbering_xml.encode("utf-8"))
    for paragraph in list(pkg.paragraphs()):
        pkg.body.remove(paragraph)
    return pkg


def _add_paragraph(
    pkg: DocxPackage,
    text: str,
    *,
    style_id: str | None = None,
    num_id: str | None = None,
    ilvl: int = 0,
) -> etree._Element:
    paragraph = etree.SubElement(pkg.body, qn("w:p"))
    if style_id or num_id is not None:
        ppr = etree.SubElement(paragraph, qn("w:pPr"))
        if style_id:
            etree.SubElement(ppr, qn("w:pStyle")).set(w_attr("val"), style_id)
        if num_id is not None:
            num_pr = etree.SubElement(ppr, qn("w:numPr"))
            etree.SubElement(num_pr, qn("w:ilvl")).set(w_attr("val"), str(ilvl))
            etree.SubElement(num_pr, qn("w:numId")).set(w_attr("val"), num_id)
    run = etree.SubElement(paragraph, qn("w:r"))
    etree.SubElement(run, qn("w:t")).text = text
    return paragraph


def _texts(pkg: DocxPackage) -> list[str]:
    return [paragraph_text(paragraph) for paragraph in pkg.paragraphs()]


def test_deep_level_keeps_counting_when_lvl_restart_is_zero(tmp_path: Path) -> None:
    """lvlRestart=0（永不重置）：上层变化后子号继续累加，Word 就是这么显示的。

    真实文档里「11.2.3 地质条件」下的四级标题是「11.2.3.7」——
    1~6 早就在第 7 章用掉了。还原成「11.2.3.1」就和 Word 对不上。
    """
    pkg = _open_with_numbering(tmp_path, _numbering_xml(deep_restart=0))
    _add_paragraph(pkg, "甲", num_id="1", ilvl=0)
    _add_paragraph(pkg, "乙", num_id="1", ilvl=1)
    _add_paragraph(pkg, "丙", num_id="1", ilvl=2)
    _add_paragraph(pkg, "丁", num_id="7", ilvl=3)
    _add_paragraph(pkg, "戊", num_id="1", ilvl=2)
    _add_paragraph(pkg, "己", num_id="7", ilvl=3)

    NumberingProcessor(pkg).convert()

    assert _texts(pkg) == [
        "1 甲",
        "1.1 乙",
        "1.1.1 丙",
        "1.1.1.1 丁",
        "1.1.2 戊",
        "1.1.2.2 己",
    ]


def test_deep_level_restarts_by_default(tmp_path: Path) -> None:
    """没有 lvlRestart 时，上层编号一变，更深层级重新从 1 数。"""
    pkg = _open_with_numbering(tmp_path, _numbering_xml())
    _add_paragraph(pkg, "甲", num_id="1", ilvl=0)
    _add_paragraph(pkg, "乙", num_id="1", ilvl=1)
    _add_paragraph(pkg, "丙", num_id="1", ilvl=2)
    _add_paragraph(pkg, "丁", num_id="7", ilvl=3)
    _add_paragraph(pkg, "戊", num_id="1", ilvl=2)
    _add_paragraph(pkg, "己", num_id="7", ilvl=3)

    NumberingProcessor(pkg).convert()

    assert _texts(pkg) == [
        "1 甲",
        "1.1 乙",
        "1.1.1 丙",
        "1.1.1.1 丁",
        "1.1.2 戊",
        "1.1.2.1 己",
    ]


def test_num_id_zero_cancels_style_numbering(tmp_path: Path) -> None:
    """numId=0 是显式取消编号，不能回落到样式上的编号。

    真实文档里「我国现行的设计规范、规程及标准：」就是标题 3 样式 + numId=0，
    Word 里这一条没有编号。若回落到样式编号，后面所有三级标题都会 +1。
    """
    pkg = _open_with_numbering(tmp_path, _numbering_xml())
    styles = pkg.xml(PART_STYLES)
    style = etree.SubElement(styles, qn("w:style"))
    style.set(w_attr("type"), "paragraph")
    style.set(w_attr("styleId"), "FakeHeading3")
    ppr = etree.SubElement(style, qn("w:pPr"))
    num_pr = etree.SubElement(ppr, qn("w:numPr"))
    etree.SubElement(num_pr, qn("w:ilvl")).set(w_attr("val"), "2")
    etree.SubElement(num_pr, qn("w:numId")).set(w_attr("val"), "1")

    _add_paragraph(pkg, "有编号", style_id="FakeHeading3")
    _add_paragraph(pkg, "无编号", style_id="FakeHeading3", num_id="0")

    NumberingProcessor(pkg).convert()

    assert _texts(pkg) == ["1.1.1 有编号", "无编号"]


def test_style_numbering_is_converted_once(tmp_path: Path) -> None:
    """样式上的编号转成文本后必须摘掉样式里的 numPr，否则打开文档会重复编号。"""
    pkg = _open_with_numbering(tmp_path, _numbering_xml())
    styles = pkg.xml(PART_STYLES)
    style = etree.SubElement(styles, qn("w:style"))
    style.set(w_attr("type"), "paragraph")
    style.set(w_attr("styleId"), "FakeHeading1")
    ppr = etree.SubElement(style, qn("w:pPr"))
    num_pr = etree.SubElement(ppr, qn("w:numPr"))
    etree.SubElement(num_pr, qn("w:numId")).set(w_attr("val"), "1")

    _add_paragraph(pkg, "总论", style_id="FakeHeading1")

    assert NumberingProcessor(pkg).convert() == 1
    assert _texts(pkg) == ["1 总论"]
    assert style.find(f"{qn('w:pPr')}/{qn('w:numPr')}") is None


@pytest.mark.parametrize(
    ("value", "num_fmt", "expected"),
    [
        (1, "decimal", "1"),
        (7, "decimal", "7"),
        (1, "decimalEnclosedCircleChinese", "①"),
        (12, "decimalEnclosedCircle", "⑫"),
        (3, "chineseCounting", "三"),
        (11, "chineseCountingThousand", "十一"),
        (24, "chineseCounting", "二十四"),
        (4, "lowerLetter", "d"),
        (28, "upperLetter", "BB"),
        (9, "lowerRoman", "ix"),
        (14, "upperRoman", "XIV"),
        (3, "none", ""),
    ],
)
def test_num_format_rendering(value: int, num_fmt: str, expected: str) -> None:
    assert format_number_value(value, num_fmt) == expected


def test_suff_nothing_keeps_label_glued(tmp_path: Path) -> None:
    """suff=nothing 表示编号与正文紧贴（「（1）配料室」），不再补分隔符。"""
    numbering = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering {_NS}>
  <w:abstractNum w:abstractNumId="5">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/><w:numFmt w:val="decimal"/>
      <w:lvlText w:val="（%1）"/><w:suff w:val="nothing"/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="10"><w:abstractNumId w:val="5"/></w:num>
</w:numbering>
"""
    pkg = _open_with_numbering(tmp_path, numbering)
    _add_paragraph(pkg, "配料室：", num_id="10")

    NumberingProcessor(pkg).convert()

    assert _texts(pkg) == ["（1）配料室："]
