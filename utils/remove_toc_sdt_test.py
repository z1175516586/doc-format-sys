# -*- coding: utf-8 -*-

import os
import zipfile
import tempfile
import shutil
from lxml import etree


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
}


def unwrap_sdt_keep_content(sdt):
    """
    删除 <w:sdt> 标签，保留里面内容
    """

    parent = sdt.getparent()
    if parent is None:
        return 0

    sdt_content = sdt.find("w:sdtContent", namespaces=NS)
    if sdt_content is None:
        return 0

    insert_index = parent.index(sdt)

    children = list(sdt_content)

    for child in children:
        sdt_content.remove(child)
        parent.insert(insert_index, child)
        insert_index += 1

    parent.remove(sdt)

    return 1


def is_toc_container(sdt):
    """
    判断是否是目录总控件
    依据：
    内部是否包含 “目录”
    """

    texts = sdt.xpath(".//w:t/text()", namespaces=NS)
    all_text = "".join(texts).strip()

    return all_text.startswith("目录")


def remove_toc_sdt_from_xml(xml_bytes):

    parser = etree.XMLParser(remove_blank_text=False)
    root = etree.fromstring(xml_bytes, parser)

    removed_count = 0

    # 找所有 sdt（非常关键）
    all_sdts = root.xpath(".//w:sdt", namespaces=NS)

    # 找目录总容器
    toc_sdt = None

    for sdt in all_sdts:
        if is_toc_container(sdt):
            toc_sdt = sdt
            break

    if toc_sdt is None:
        print("未找到目录内容控件")
        return xml_bytes

    print("找到目录内容控件，开始清理...")

    # 倒序删，避免 parent 结构变化影响 xpath
    nested_sdts = toc_sdt.xpath(".//w:sdt", namespaces=NS)

    for sdt in reversed(nested_sdts):
        removed_count += unwrap_sdt_keep_content(sdt)

    # 最外层也删掉
    removed_count += unwrap_sdt_keep_content(toc_sdt)

    print(f"共删除 {removed_count} 个目录内容控件")

    return etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone="yes"
    )


def remove_toc_sdt(docx_path):
    """
    原地覆盖 docx
    """

    temp_dir = tempfile.mkdtemp()

    try:
        with zipfile.ZipFile(docx_path, "r") as zip_ref:
            zip_ref.extractall(temp_dir)

        document_xml = os.path.join(
            temp_dir,
            "word",
            "document.xml"
        )

        with open(document_xml, "rb") as f:
            xml_data = f.read()

        new_xml = remove_toc_sdt_from_xml(xml_data)

        with open(document_xml, "wb") as f:
            f.write(new_xml)

        temp_docx = docx_path + ".tmp"

        with zipfile.ZipFile(temp_docx, "w", zipfile.ZIP_DEFLATED) as zf:
            for foldername, _, filenames in os.walk(temp_dir):
                for filename in filenames:
                    file_path = os.path.join(foldername, filename)

                    arcname = os.path.relpath(
                        file_path,
                        temp_dir
                    )

                    zf.write(file_path, arcname)

        os.replace(temp_docx, docx_path)

        print("处理完成（原文件已覆盖）")

    finally:
        shutil.rmtree(temp_dir)


if __name__ == "__main__":

    file_path = r"D:\ASE_develop_file\doc-convertor\test\2026年5月29日 - 解决源文档目录中包含文本框的问题\1#包装机组可研2026.4.17 - test.docx"

    remove_toc_sdt(file_path)