# -*- coding: utf-8 -*-
"""
将“原文件”中的所有表格按在正文中的位置顺序，
替换到“被修改的文件”的相同位置上。

依赖：
    pip install python-docx

使用方法：
    1) 修改下方路径常量 ORIGINAL_DOC_PATH / EDITED_DOC_PATH / OUTPUT_DOC_PATH
    2) 运行脚本：python replace_tables_by_position.py
"""

import copy
from pathlib import Path
from typing import Any, Iterator

# 导入Document类（忽略静态分析工具的警告）
# pyright: reportAttributeAccessIssue=false
try:
    from docx import Document
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P
    DOCX_AVAILABLE = True
except ImportError:
    Document = None
    CT_Tbl = type(None)  # 使用一个实际的类型而不是None
    CT_P = type(None)    # 使用一个实际的类型而不是None
    DOCX_AVAILABLE = False

# ====== 路径配置（按需修改）======
ORIGINAL_DOC_PATH = "input/test(1).docx"   # “原文件”（想要抽取表格的文件）
EDITED_DOC_PATH   = "input/step2_pandoc_转换成功_20250917_211258.docx"     # “被修改的文件”（要被替换表格的文件）
OUTPUT_DOC_PATH   = "output/replaced.docx"  # 输出文件
# =================================


def iter_body_children(doc) -> Iterator[Any]:
    """
    迭代正文 body 的直接子节点，保持原有顺序。
    只区分段落(CT_P)与表格(CT_Tbl)，其余节点直接返回原 oxml 以免改动顺序。
    """
    if not DOCX_AVAILABLE or doc is None:
        return iter([])
    
    body = doc.element.body
    for child in body.iterchildren():
        yield child


def collect_tables_in_body(doc) -> list:
    """
    按正文顺序收集表格的底层 OOXML 节点 (CT_Tbl)。
    不进入页眉/页脚/文本框/形状等，仅处理 document.body。
    """
    if not DOCX_AVAILABLE or doc is None:
        return []
    
    tables = []
    for child in iter_body_children(doc):
        if DOCX_AVAILABLE and CT_Tbl is not type(None) and isinstance(child, CT_Tbl):
            tables.append(child)
    return tables

def remove_table_objects(table_xml):
    """
    删除表格中的OLE对象和嵌入对象

    清理:
        w:object
        o:OLEObject
        v:shape
    """

    from lxml import etree

    # 删除 w:object
    objects = table_xml.xpath(
        ".//w:object"
    )

    count = 0

    for obj in objects:
        parent = obj.getparent()

        if parent is not None:
            parent.remove(obj)
            count += 1


    # 删除 OLE关系节点残留
    ole_objects = table_xml.xpath(
        ".//*[local-name()='OLEObject']"
    )

    print("删除表格中的OLE对象")

    for ole in ole_objects:
        parent = ole.getparent()
        if parent is not None:
            parent.remove(ole)


    return count

def replace_tables_by_index(original_path, edited_path, output_path):
    """
    将原文件中的所有表格按在正文中的位置顺序，替换到被修改的文件的相同位置上。
    并在插入表格时，将紧邻的表题注段落（如“表 10.4.1 主要设备表”或“表 10-4 主要设备表”）强制设置为正文。
    """
    import re
    import copy
    from pathlib import Path
    from docx import Document
    from docx.shared import Pt
    from docx.table import Table

    # ⚡ 修改：表题注识别正则，支持 . 或 -
    TABLE_CAPTION_PATTERN = re.compile(
        r'^表\s*\d+(?:[.-]\d+){0,3}\s*.+$'
    )

    if not DOCX_AVAILABLE or Document is None:
        raise ImportError("python-docx库不可用，请安装 python-docx")

    if not Path(original_path).exists() or not Path(edited_path).exists():
        raise FileNotFoundError("请确认 ORIGINAL_DOC_PATH 与 EDITED_DOC_PATH 文件存在。")

    print(f"📄 加载原文件: {original_path}")
    doc_original = Document(original_path)

    print(f"📄 加载被修改文件: {edited_path}")
    doc_edited = Document(edited_path)

    print("🟡 开始收集原文件表格")
    try:
        orig_tables = collect_tables_in_body(doc_original)
    except Exception as e:
        print("❌ 原文件 collect_tables_in_body失败")
        import traceback
        traceback.print_exc()
        raise

    print("🟡 开始收集修改文件表格")
    try:
        edited_tables = collect_tables_in_body(doc_edited)
    except Exception as e:
        print("❌ 修改文件 collect_tables_in_body失败")
        import traceback
        traceback.print_exc()
        raise

    print(f"🔎 原文件表格数量: {len(orig_tables)}")
    print(f"🔎 被修改文件表格数量: {len(edited_tables)}")

    if len(orig_tables) == 0:
        doc_edited.save(output_path)
        return True

    n = min(len(orig_tables), len(edited_tables))
    body = doc_edited.element.body

    def reset_prev_paragraph_if_table_caption(inserted_tbl):
        """
        ⚡ 新增：
        如果表格紧邻的上一个段落是“表 x.x 表名”或“表 x-x 表名”形式，
        则强制将其样式设为 Normal
        """
        parent = inserted_tbl.getparent()
        children = list(parent)

        idx = children.index(inserted_tbl)
        if idx == 0:
            return

        prev = children[idx - 1]

        # 只处理段落
        from docx.oxml.text.paragraph import CT_P
        if not isinstance(prev, CT_P):
            return

        from docx.text.paragraph import Paragraph
        para = Paragraph(prev, doc_edited)

        text = para.text.strip()
        if TABLE_CAPTION_PATTERN.match(text):
            para.style = 'Normal'
            para.paragraph_format.space_before = Pt(0)
            para.paragraph_format.space_after = Pt(0)
            print(f"⚠️ 表题注重置为正文: {text}")

    def find_i_th_table_and_replace(i, new_tbl_oxml):
        idx = -1
        for child in body.iterchildren():
            from docx.oxml.table import CT_Tbl
            if isinstance(child, CT_Tbl):
                idx += 1
                if idx == i:
                    insert_at = list(body).index(child)

                    inserted_tbl = copy.deepcopy(new_tbl_oxml)

                    # ===============================
                    # 清理Word OLE对象
                    # ===============================
                    print("🟡 开始清理表格对象")
                    try:
                        removed_objects = remove_table_objects(inserted_tbl)
                    except Exception as e:
                        print("❌ remove_table_objects失败")
                        print(type(e))
                        print(e)
                        import traceback
                        traceback.print_exc()
                        raise

                    if removed_objects:
                        print(
                            f"🧹 清理表格OLE对象: 第{i + 1}个表格 删除 {removed_objects} 个"
                        )
                    body.insert(insert_at, inserted_tbl)
                    body.remove(child)

                    # ⚡ 处理表格上方的表题注
                    reset_prev_paragraph_if_table_caption(inserted_tbl)

                    # 原有逻辑：表格内部全部设为 Normal
                    table_obj = Table(inserted_tbl, doc_edited)
                    for row in table_obj.rows:
                        for cell in row.cells:
                            for para in cell.paragraphs:
                                para.style = 'Normal'
                                para.paragraph_format.space_before = Pt(0)
                                para.paragraph_format.space_after = Pt(0)

                    return True
        return False

    replaced = 0
    for i in range(n):
        if find_i_th_table_and_replace(i, orig_tables[i]):
            replaced += 1
            print(f"✅ 已替换第 {i + 1} 个表格")

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    doc_edited.save(output_path)

    print(f"\n🎉 完成！共替换 {replaced}/{n} 个表格")
    print(f"📁 输出文件: {output_path}")
    return True


if __name__ == "__main__":
    replace_tables_by_index(ORIGINAL_DOC_PATH, EDITED_DOC_PATH, OUTPUT_DOC_PATH)