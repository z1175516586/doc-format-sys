# -*- coding: utf-8 -*-

"""
记录并恢复 Word 表格字体

用途：
1. 自动排版前记录每个表格原始字体信息
2. 自动排版后整体恢复表格字体

记录维度：
(table_idx) -> {
    font_name,
    font_size
}
"""

from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from typing import Dict, Optional
from docx.opc.constants import RELATIONSHIP_TYPE as RT

def _get_style_font_size(doc: Document, style_id: str) -> Optional[float]:
    """
    正确：styleId → styles.xml → rPr → sz
    支持 basedOn + DefaultChar
    """

    if not style_id:
        return None

    try:
        style_part = None
        for rel in doc.part.rels.values():
            if "styles" in rel.reltype:
                style_part = rel._target
                break

        if style_part is None:
            return None

        xml = style_part.blob.decode("utf-8", errors="ignore")

    except Exception as e:
        print("[style parse error]", e)
        return None

    from xml.etree import ElementTree as ET
    root = ET.fromstring(xml)

    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

    # =========================
    # 建索引：styleId -> node
    # =========================
    style_map = {}

    for s in root.findall("w:style", ns):
        sid = s.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId")
        style_map[sid] = s

    # =========================
    # 递归查 size
    # =========================
    def resolve_size(sid, depth=0):
        if not sid or depth > 5:
            return None

        node = style_map.get(sid)
        if node is None:
            return None

        rpr = node.find("w:rPr", ns)
        if rpr is not None:
            sz = rpr.find("w:sz", ns)
            if sz is not None:
                val = sz.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")
                if val:
                    print(f"[STYLE HIT] {sid} size={val}")
                    return int(val) / 2

        # basedOn 递归（关键）
        based_on = node.find("w:basedOn", ns)
        if based_on is not None:
            parent = based_on.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")
            return resolve_size(parent, depth + 1)

        return None

    return resolve_size(style_id)

def _get_doc_styles(doc: Document):
    """
    解析 styles.xml（同时支持 rPr + pPr）
    """

    styles = {}

    try:
        style_part = None

        for rel in doc.part.rels.values():
            if "styles" in rel.reltype:
                style_part = rel._target
                break

        if style_part is None:
            return styles

        xml = style_part.blob.decode("utf-8", errors="ignore")

    except Exception:
        return styles

    from xml.etree import ElementTree as ET
    root = ET.fromstring(xml)

    ns = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    }

    for style in root.findall("w:style", ns):
        style_id = style.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId")

        # ⭐ 1. run级字号
        rpr = style.find(".//w:rPr", ns)
        sz = None

        if rpr is not None:
            sz = rpr.find("w:sz", ns)

        # ⭐ 2. paragraph级字号（你现在缺的关键）
        if sz is None:
            ppr = style.find(".//w:pPr", ns)
            if ppr is not None:
                sz = ppr.find("w:sz", ns)

        if sz is not None:
            val = sz.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")
            if val:
                styles[style_id] = int(val) / 2

    return styles

def _get_font_size(run) -> Optional[float]:

    doc = run.part.document
    styles = _get_doc_styles(doc)

    # 1. run级
    if run.font.size:
        return run.font.size.pt

    rPr = run._element.rPr

    # 2. run rPr
    if rPr is not None:
        sz = rPr.find(qn("w:sz"))
        if sz is not None:
            val = sz.get(qn("w:val"))
            if val:
                return int(val) / 2

        sz_cs = rPr.find(qn("w:szCs"))
        if sz_cs is not None:
            val = sz_cs.get(qn("w:val"))
            if val:
                return int(val) / 2

    # 3. paragraph rPr
    p = run._element.getparent().getparent()
    if p is not None:
        pPr = p.find(qn("w:pPr"))
        if pPr is not None:
            p_rPr = pPr.find(qn("w:rPr"))
            if p_rPr is not None:
                sz = p_rPr.find(qn("w:sz"), qn)
                if sz is not None:
                    val = sz.get(qn("w:val"))
                    if val:
                        return int(val) / 2

    # 4. ⭐ style fallback（严格按照 document → style.xml 链路）

    from docx.oxml.text.paragraph import CT_P

    def _get_paragraph(run):
        el = run._element
        while el is not None and not isinstance(el, CT_P):
            el = el.getparent()
        return el

    p = _get_paragraph(run)

    style_id = None

    if p is not None:
        pPr = p.find(qn("w:pPr"))

        if pPr is not None:
            pStyle = pPr.find(qn("w:pStyle"))

            if pStyle is not None:
                style_id = pStyle.get(qn("w:val"))

    # ===== ⭐ 从 styles.xml 精确查 rPr =====
    if style_id:
        size = _get_style_font_size(doc, style_id)
        if size:
            print("[FINAL STYLE SIZE]", size)
            return size

    print("[STEP] run.font.size =", run.font.size)

    print("[STEP] rPr size parsed =", None)
    print("[STEP] pPr size parsed =", None)
    print("[STEP] style_id =", style_id if 'style_id' in locals() else None)

    print("====== DEBUG FONT SIZE ======")
    print("text:", run.text)
    print("run.font.size:", run.font.size)
    print("style_id:", style_id)
    print("=============================")

    return 12.0

def _get_font_name(run) -> Optional[str]:
    """
    获取 run 字体名称
    优先取中文字体 eastAsia
    """

    rPr = run._element.rPr

    if rPr is None or rPr.rFonts is None:
        return None

    east_asia = rPr.rFonts.get(qn("w:eastAsia"))

    if east_asia:
        return east_asia

    if run.font.name:
        return run.font.name

    return None


def _set_run_font(run, font_name: str, font_size: float):

    if font_size is None:
        font_size = 12.0
    else:
        try:
            font_size = float(font_size)
        except Exception:
            font_size = 12.0

    if not font_name:
        font_name = "宋体"

    run.font.name = font_name
    run.font.size = Pt(font_size)

    rPr = run._element.get_or_add_rPr()

    # 字体名
    if rPr.rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    else:
        rFonts = rPr.rFonts

    rFonts.set(qn("w:eastAsia"), font_name)
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)

    # -----------------------------
    # 强制补写 w:sz + w:szCs
    # -----------------------------

    sz_val = str(int(round(font_size * 2)))   # 10.5 -> 21

    sz = rPr.find(qn("w:sz"))
    if sz is None:
        sz = OxmlElement("w:sz")
        rPr.append(sz)
    sz.set(qn("w:val"), sz_val)

    sz_cs = rPr.find(qn("w:szCs"))
    if sz_cs is None:
        sz_cs = OxmlElement("w:szCs")
        rPr.append(sz_cs)
    sz_cs.set(qn("w:val"), sz_val)

def record_table_font_sizes(doc_path: str) -> Dict:
    """
    记录每个表格的字体信息

    Returns:
    {
        table_idx: {
            "font_name": "宋体",
            "font_size": 10.5
        }
    }
    """

    doc = Document(doc_path)

    table_font_map = {}

    for t_idx, table in enumerate(doc.tables):

        font_name = None
        font_size = None

        found = False

        for row in table.rows:
            if found:
                break

            for cell in row.cells:
                if found:
                    break

                for paragraph in cell.paragraphs:
                    if found:
                        break

                    for run in paragraph.runs:

                        text = run.text.strip()

                        if not text:
                            continue

                        font_name = _get_font_name(run) or "宋体"
                        font_size = _get_font_size(run) or 12.0

                        if font_name or font_size:
                            found = True
                            break

        if font_name or font_size:
            table_font_map[t_idx] = {
                "font_name": font_name,
                "font_size": font_size
            }

    print(
        f"📌 已记录 {len(table_font_map)} 个表格字体信息"
    )

    return table_font_map


def restore_table_font_sizes(
        doc_path: str,
        table_font_map: Dict
) -> None:
    """
    按表格整体恢复字体
    """

    doc = Document(doc_path)

    restore_count = 0

    for t_idx, font_info in table_font_map.items():

        if t_idx >= len(doc.tables):
            continue

        table = doc.tables[t_idx]

        font_name = font_info["font_name"]
        font_size = font_info["font_size"]

        if not font_name:
            font_name = "宋体"

        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        _set_run_font(
                            run,
                            font_name,
                            font_size
                        )
                        restore_count += 1

    doc.save(doc_path)

    print(
        f"✅ 已恢复表格字体，共恢复 {restore_count} 个 run"
    )


if __name__ == "__main__":

    test_file_A = r"D:\ASE_develop_file\doc-convertor\test\2026年5月25日\A.docx"
    test_file_B = r"D:\ASE_develop_file\doc-convertor\test\2026年5月25日\B.docx"

    font_map = record_table_font_sizes(test_file_A)

    print(font_map)

    restore_table_font_sizes(
        test_file_B,
        font_map
    )

    print("测试完成")