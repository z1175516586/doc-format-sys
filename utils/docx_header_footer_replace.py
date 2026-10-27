"""
DOCX页眉页脚内容提取与替换工具

该模块提供了以下功能：
1. 提取DOCX文档中指定节的页眉页脚内容（不包括页码）
2. 将提取的页眉页脚内容替换到另一个DOCX文档的对应节中
3. 保持目标文档原有的页码、格式等内容不变

使用方法：
- extract_header_footer_content(docx_path, section_index): 提取指定节的页眉页脚内容
- replace_header_footer_content(source_docx_path, target_docx_path, source_section_index, target_section_index, save_path): 
  将源文档的页眉页脚内容替换到目标文档
"""

import os
import re
import zipfile
from lxml import etree
from typing import Dict, Any

# Add the docx import at the top
try:
    from docx import Document
    HAS_PYTHON_DOCX = True
except ImportError:
    HAS_PYTHON_DOCX = False
    print("Warning: python-docx library not found. Using fallback XML parsing method.")

def extract_header_footer_content(docx_path: str, section_index: int = 1) -> Dict[str, Any]:
    """
    提取DOCX文档中指定节的页眉页脚内容（不包括页码）
    
    Args:
        docx_path (str): DOCX文件路径
        section_index (int): 节索引（从1开始）
        
    Returns:
        Dict[str, Any]: 包含页眉和页脚内容的字典
    """
    print(f"🔍 开始提取文档页眉页脚内容: {docx_path}")
    print(f"📝 目标节号: {section_index}")
    
    # If python-docx is available, use the simpler approach
    if HAS_PYTHON_DOCX:
        return _extract_header_footer_with_docx(docx_path, section_index)
    
    # Fallback to the original XML parsing approach
    return _extract_header_footer_with_xml(docx_path, section_index)

from lxml import etree

def extract_text_from_shapes(element):
    """
    从 python-docx 的 element（页眉/页脚）中提取文本，包括：
    - 普通段落
    - w:drawing / wps:txbx 文本框
    - w:pict / v:textbox 文本框
    """
    texts = []

    # 1. 获取 XML 字符串并用 lxml 解析
    xml = element._element.xml
    root = etree.fromstring(xml.encode("utf-8"))

    # 2. 定义命名空间
    nsmap = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
        "v": "urn:schemas-microsoft-com:vml",
        "o": "urn:schemas-microsoft-com:office:office",
        "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    }

    # 3. 抓取普通段落文字
    for t in root.xpath(".//w:p//w:r//w:t", namespaces=nsmap):
        if t.text and t.text.strip():
            texts.append(t.text.strip())

    # 4. w:drawing -> wps:txbx -> w:txbxContent -> w:p -> w:r -> w:t
    for t in root.xpath(".//w:drawing//wps:txbx//w:t", namespaces=nsmap):
        if t.text and t.text.strip():
            texts.append(t.text.strip())

    # 5. w:pict -> v:textbox -> w:txbxContent -> w:p -> w:r -> w:t
    for t in root.xpath(".//w:pict//w:txbxContent//w:t", namespaces=nsmap):
        if t.text and t.text.strip():
            texts.append(t.text.strip())

    return texts


def _extract_header_footer_with_docx(docx_path: str, section_index: int) -> Dict[str, Any]:
    """
    使用python-docx库提取页眉页脚内容（支持文本框）
    """
    from docx.oxml import OxmlElement
    from docx.document import Document as _Document

    def get_element_safe(obj):
        """
        安全获取 _element，兼容 CT_HdrFtr / Header / Footer
        """
        if hasattr(obj, "_element"):
            return obj._element
        return obj  # 如果已经是 CT_HdrFtr

    try:
        doc = Document(docx_path)

        if len(doc.sections) < section_index:
            raise ValueError(f"文档共有 {len(doc.sections)} 节，不能操作第 {section_index} 节")

        result = {"headers": {}, "footers": {}}

        sections_to_check = [doc.sections[section_index - 1]]
        sections_to_check.extend(doc.sections[section_index:])

        print(f"🔍 将检查 {len(sections_to_check)} 个节以查找页眉页脚内容")

        for i, section in enumerate(sections_to_check):
            # 页眉
            header = section.header
            header_texts = [p.text for p in getattr(header, "paragraphs", []) if p.text.strip()]
            if not header_texts:
                header_texts.extend(extract_text_from_shapes(get_element_safe(header)))

            if header_texts:
                result["headers"]["default"] = header_texts[0]
                print(f"  📄 找到页眉内容: {result['headers']['default']}")

            # 页脚
            footer = section.footer
            footer_texts = [p.text for p in getattr(footer, "paragraphs", []) if p.text.strip()]
            if not footer_texts:
                footer_texts.extend(extract_text_from_shapes(get_element_safe(footer)))

            if footer_texts:
                result["footers"]["default"] = "\n".join(footer_texts)
                print(f"  📄 找到页脚内容: {result['footers']['default']}")

            if result["headers"] and result["footers"]:
                break

        if result["headers"] or result["footers"]:
            print("✅ 页眉页脚内容提取完成 (使用python-docx + XML解析文本框)")
            return result

        raise ValueError("未能在文档的任何节中找到页眉或页脚内容")

    except Exception as e:
        print(f"❌ 使用python-docx提取页眉页脚内容失败: {e}")
        return _extract_header_footer_with_xml(docx_path, section_index)

def _extract_header_footer_with_xml(docx_path: str, section_index: int) -> Dict[str, Any]:
    """
    原始的XML解析方式提取页眉页脚内容（作为备选方案）
    """
    try:
        # 解压 docx 到内存
        with zipfile.ZipFile(docx_path, 'r') as zin:
            filelist = zin.namelist()
            files = {name: zin.read(name) for name in filelist}

        # 找到 document.xml，定位节与 header/footer 的对应关系
        root = etree.fromstring(files["word/document.xml"])
        nsmap = root.nsmap
        sects = root.xpath(".//w:sectPr", namespaces=nsmap)
        
        print(f"📊 找到 {len(sects)} 个节")

        if len(sects) < section_index:
            raise ValueError(f"文档共有 {len(sects)} 节，不能操作第 {section_index} 节")

        target_sect = sects[section_index - 1]
        
        # 查找该节的 headerReference 和 footerReference
        header_refs = target_sect.xpath("./w:headerReference", namespaces=nsmap)
        footer_refs = target_sect.xpath("./w:footerReference", namespaces=nsmap)
        
        print(f"📋 找到 {len(header_refs)} 个页眉引用, {len(footer_refs)} 个页脚引用")
        
        result = {
            "headers": {},
            "footers": {}
        }
        
        # 提取页眉内容（保留格式信息）
        for i, href in enumerate(header_refs):
            rid = href.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            header_type = href.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type", "default")
            
            # 根据关系文件找到 headerX.xml
            rels_name = "word/_rels/document.xml.rels"
            if rels_name in files:
                rels_root = etree.fromstring(files[rels_name])
                header_target = rels_root.xpath(f".//rel:Relationship[@Id='{rid}']",
                                                namespaces={"rel": "http://schemas.openxmlformats.org/package/2006/relationships"})
                if header_target:
                    header_file = "word/" + header_target[0].get("Target")
                    if header_file in files:
                        hroot = etree.fromstring(files[header_file])
                        print(f"  📄 header_fiel是: {header_file}")
                        print(f"  📄 hroot是: {hroot}")
                        # 提取文本内容和格式信息，排除页码域
                        header_content1 = extract_formatted_content(hroot, nsmap)
                        print(f"  📄 提取headrcontent ({header_type}): {header_content1}")
                        parts = header_content1.strip().split()
                        header_content = f"{parts[0]}\t{parts[1]}" if len(parts) >= 2 else header_content1
                        result["headers"][header_type] = header_content
                        print(f"  📄 提取页眉 ({header_type}): {header_content}")

        # 提取页脚内容（保留格式信息）
        for i, fref in enumerate(footer_refs):
            rid = fref.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            footer_type = fref.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type", "default")
            
            # 根据关系文件找到 footerX.xml
            rels_name = "word/_rels/document.xml.rels"
            if rels_name in files:
                rels_root = etree.fromstring(files[rels_name])
                footer_target = rels_root.xpath(f".//rel:Relationship[@Id='{rid}']",
                                                namespaces={"rel": "http://schemas.openxmlformats.org/package/2006/relationships"})
                if footer_target:
                    footer_file = "word/" + footer_target[0].get("Target")
                    if footer_file in files:
                        froot = etree.fromstring(files[footer_file])
                        
                        # 提取文本内容和格式信息，排除页码域
                        footer_content = extract_formatted_content(froot, nsmap)
                        result["footers"][footer_type] = footer_content
                        print(f"  📄 提取页脚 ({footer_type}): {footer_content}")
        
        print(f"✅ 页眉页脚内容提取完成")
        return result
        
    except Exception as e:
        print(f"❌ 提取页眉页脚内容失败: {e}")
        raise


def extract_formatted_content(root, nsmap):
    """
    提取带有格式信息的内容（保留制表符、换行符等）
    
    Args:
        root: XML根节点
        nsmap: 命名空间映射
        
    Returns:
        str: 格式化的内容
    """
    content_parts = []
    
    # 遍历所有段落
    for p_elem in root.xpath(".//w:p", namespaces=nsmap):
        para_parts = []
        
        # 遍历段落中的所有运行（run）
        for r_elem in p_elem.xpath(".//w:r", namespaces=nsmap):
            # 检查是否有制表符
            tabs = r_elem.xpath(".//w:tab", namespaces=nsmap)
            for _ in tabs:
                para_parts.append("\t")
            
            # 检查是否有文本
            for t_elem in r_elem.xpath(".//w:t", namespaces=nsmap):
                if t_elem.text:
                    para_parts.append(t_elem.text)
            
            # 检查是否有换行符
            brs = r_elem.xpath(".//w:br", namespaces=nsmap)
            for _ in brs:
                para_parts.append("\n")
        
        # 将段落内容添加到结果中
        content_parts.append("".join(para_parts))
    print(f"  📄 为什么呢: {content_parts}")
    # 用换行符连接所有段落
    return "\n".join(content_parts)


def replace_formatted_content(root, new_content, nsmap):
    """
    替换带有格式的内容（保留原有格式结构）
    修改：优先替换 SDT（内容控件）中的文本；若不存在，再使用原有逻辑
    页码域（PAGE）始终保留
    """

    # ===============================
    # ✅ 第一优先级：替换 SDT 内容
    # ===============================
    sdt_texts = root.xpath(".//w:sdtContent//w:t", namespaces=nsmap)
    if sdt_texts:
        print(f"  📦 命中 SDT 内容控件，替换 {len(sdt_texts)} 个文本节点")
        for t in sdt_texts:
            t.text = new_content
        return  # ⚠️ 关键：SDT 命中后直接返回，不再走下面逻辑

    # ===============================
    # ⬇️ 以下为你原来的逻辑（基本未动）
    # ===============================

    # 获取所有段落
    p_elems = root.xpath(".//w:p", namespaces=nsmap)

    # 按行分割新内容
    lines = new_content.split('\n')

    for p_index, p_elem in enumerate(p_elems):
        if p_index >= len(lines):
            break

        line = lines[p_index]

        # 按制表符分割内容
        parts = []
        current_part = ""
        for char in line:
            if char == '\t':
                parts.append(current_part)
                parts.append('\t')
                current_part = ""
            else:
                current_part += char
        parts.append(current_part)

        # 获取段落中的所有运行
        r_elems = p_elem.xpath(".//w:r", namespaces=nsmap)
        part_index = 0

        for r_elem in r_elems:
            # 检查制表符
            tab_elems = r_elem.xpath(".//w:tab", namespaces=nsmap)
            ptab_elems = r_elem.xpath(".//w:ptab", namespaces=nsmap)

            # 判断是否为页码域
            is_in_page_field = False
            instr_elems = r_elem.xpath(".//w:instrText", namespaces=nsmap)
            for instr_elem in instr_elems:
                if instr_elem.text and "PAGE" in instr_elem.text.upper():
                    is_in_page_field = True
                    break

            if is_in_page_field:
                # 页码域保持不变
                continue

            if tab_elems or ptab_elems:
                while part_index < len(parts) and parts[part_index] != '\t':
                    part_index += 1

                if part_index < len(parts) and parts[part_index] == '\t':
                    if not tab_elems and not ptab_elems:
                        etree.SubElement(
                            r_elem,
                            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tab"
                        )
                    t_elems = r_elem.xpath(".//w:t", namespaces=nsmap)
                    for t_elem in t_elems:
                        t_elem.getparent().remove(t_elem)
                    part_index += 1
            else:
                t_elems = r_elem.xpath(".//w:t", namespaces=nsmap)
                if t_elems and part_index < len(parts):
                    while part_index < len(parts) and parts[part_index] == '\t':
                        part_index += 1
                    if part_index < len(parts):
                        t_elems[0].text = parts[part_index]
                        part_index += 1
                elif t_elems:
                    t_elems[0].text = ""


def replace_header_footer_content(source_docx_path: str, target_docx_path: str, 
                                 source_section_index: int = 1, target_section_index: int = 1,
                                 save_path: str = "", proj_name: str = "", file_type_name: str = "") -> bool:
    """
    将源文档的页眉页脚内容替换到目标文档的指定节中，保持页码和格式不变
    
    Args:
        source_docx_path (str): 源DOCX文件路径（提供页眉页脚内容）
        target_docx_path (str): 目标DOCX文件路径（被替换页眉页脚内容）
        source_section_index (int): 源文档节索引（从1开始）
        target_section_index (int): 目标文档节索引（从1开始）
        save_path (str): 保存路径，如果为None则覆盖目标文件
        
    Returns:
        bool: 操作是否成功
    """
    print(f"🔄 开始替换页眉页脚内容")
    print(f"  源文档: {source_docx_path} (第{source_section_index}节)")
    print(f"  目标文档: {target_docx_path} (第{target_section_index}节)")
    
    try:
        # 提取源文档的页眉页脚内容
        source_content = extract_header_footer_content(source_docx_path, source_section_index)
        
        # 解压源文档到内存，获取段落属性
        with zipfile.ZipFile(source_docx_path, 'r') as zin:
            source_filelist = zin.namelist()
            source_files = {name: zin.read(name) for name in source_filelist}
        
        # 解压目标文档到内存
        with zipfile.ZipFile(target_docx_path, 'r') as zin:
            filelist = zin.namelist()
            files = {name: zin.read(name) for name in filelist}

        # 找到源文档的 document.xml，定位节与 header/footer 的对应关系
        source_document_root = etree.fromstring(source_files["word/document.xml"])
        source_nsmap = source_document_root.nsmap
        source_sects = source_document_root.xpath(".//w:sectPr", namespaces=source_nsmap)
        
        if len(source_sects) < source_section_index:
            raise ValueError(f"源文档共有 {len(source_sects)} 节，不能操作第 {source_section_index} 节")

        source_sect = source_sects[source_section_index - 1]
        source_header_refs = source_sect.xpath("./w:headerReference", namespaces=source_nsmap)
        source_footer_refs = source_sect.xpath("./w:footerReference", namespaces=source_nsmap)
        
        # 找到目标文档的 document.xml，定位节与 header/footer 的对应关系
        document_root = etree.fromstring(files["word/document.xml"])
        nsmap = document_root.nsmap
        sects = document_root.xpath(".//w:sectPr", namespaces=nsmap)
        
        print(f"📊 目标文档找到 {len(sects)} 个节")

        if len(sects) < target_section_index:
            raise ValueError(f"目标文档共有 {len(sects)} 节，不能操作第 {target_section_index} 节")

        target_sect = sects[target_section_index - 1]
        
        # 查找该节的 headerReference 和 footerReference
        header_refs = target_sect.xpath("./w:headerReference", namespaces=nsmap)
        footer_refs = target_sect.xpath("./w:footerReference", namespaces=nsmap)
        print(
            f"📋 目标节 footerReference 数量: {len(footer_refs)}",
            [fref.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type") for fref in footer_refs]
        )
        
        print(f"📋 找到 {len(header_refs)} 个页眉引用, {len(footer_refs)} 个页脚引用")
        
        # 创建一个字典来存储源文档的段落属性
        source_header_properties = {}
        source_footer_properties = {}
        
        # 提取源文档页眉的段落属性
        for i, href in enumerate(source_header_refs):
            rid = href.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            header_type = href.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type", "default")
            
            # 根据关系文件找到 headerX.xml
            rels_name = "word/_rels/document.xml.rels"
            if rels_name in source_files:
                rels_root = etree.fromstring(source_files[rels_name])
                header_target = rels_root.xpath(f".//rel:Relationship[@Id='{rid}']",
                                                namespaces={"rel": "http://schemas.openxmlformats.org/package/2006/relationships"})
                if header_target:
                    header_file = "word/" + header_target[0].get("Target")
                    if header_file in source_files:
                        hroot = etree.fromstring(source_files[header_file])
                        header_nsmap = hroot.nsmap
                        
                        # 提取段落属性
                        p_elems = hroot.xpath(".//w:p", namespaces=header_nsmap)
                        properties = []
                        for j, p_elem in enumerate(p_elems):
                            p_pr_elems = p_elem.xpath("./w:pPr", namespaces=header_nsmap)
                            if p_pr_elems:
                                p_pr_elem = p_pr_elems[0]
                                # 提取对齐方式
                                jc_elems = p_pr_elem.xpath("./w:jc", namespaces=header_nsmap)
                                jc_val = jc_elems[0].get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") if jc_elems else None
                                # 提取制表符停止点
                                tabs_elems = p_pr_elem.xpath("./w:tabs", namespaces=header_nsmap)
                                tabs_xml = etree.tostring(tabs_elems[0], encoding="unicode") if tabs_elems else None
                                properties.append({
                                    "jc_val": jc_val,
                                    "tabs_xml": tabs_xml
                                })
                        source_header_properties[header_type] = properties
        
        # 提取源文档页脚的段落属性
        for i, fref in enumerate(source_footer_refs):
            rid = fref.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            footer_type = fref.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type", "default")
            
            # 根据关系文件找到 footerX.xml
            rels_name = "word/_rels/document.xml.rels"
            if rels_name in source_files:
                rels_root = etree.fromstring(source_files[rels_name])
                footer_target = rels_root.xpath(f".//rel:Relationship[@Id='{rid}']",
                                                namespaces={"rel": "http://schemas.openxmlformats.org/package/2006/relationships"})
                if footer_target:
                    footer_file = "word/" + footer_target[0].get("Target")
                    if footer_file in source_files:
                        froot = etree.fromstring(source_files[footer_file])
                        footer_nsmap = froot.nsmap
                        
                        # 提取段落属性
                        p_elems = froot.xpath(".//w:p", namespaces=footer_nsmap)
                        properties = []
                        for j, p_elem in enumerate(p_elems):
                            p_pr_elems = p_elem.xpath("./w:pPr", namespaces=footer_nsmap)
                            if p_pr_elems:
                                p_pr_elem = p_pr_elems[0]
                                # 提取对齐方式
                                jc_elems = p_pr_elem.xpath("./w:jc", namespaces=footer_nsmap)
                                jc_val = jc_elems[0].get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") if jc_elems else None
                                # 提取制表符停止点
                                tabs_elems = p_pr_elem.xpath("./w:tabs", namespaces=footer_nsmap)
                                tabs_xml = etree.tostring(tabs_elems[0], encoding="unicode") if tabs_elems else None
                                properties.append({
                                    "jc_val": jc_val,
                                    "tabs_xml": tabs_xml
                                })
                        source_footer_properties[footer_type] = properties
        
        # 替换页眉内容
        for i, href in enumerate(header_refs):
            rid = href.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            header_type = href.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type", "default")
            
            # 根据关系文件找到 headerX.xml
            rels_name = "word/_rels/document.xml.rels"
            if rels_name in files:
                rels_root = etree.fromstring(files[rels_name])
                header_target = rels_root.xpath(f".//rel:Relationship[@Id='{rid}']",
                                                namespaces={"rel": "http://schemas.openxmlformats.org/package/2006/relationships"})
                if header_target:
                    header_file = "word/" + header_target[0].get("Target")
                    if header_file in files:
                        hroot = etree.fromstring(files[header_file])
                        header_nsmap = hroot.nsmap  # 使用页眉文件的命名空间映射
                        header_text = f"{proj_name}\t{file_type_name}"
                        replace_formatted_content(hroot, header_text, header_nsmap)
                        print(f"  🔄 替换页眉 ({header_type}): {header_text}")
                        # 如果源文档有对应类型的页眉内容，则替换
                        # if header_type in source_content["headers"]:
                        #     header_text = f"{proj_name}\t{file_type_name}" # source_content["headers"][header_type]t"
                        #     if header_text:  # 只有当有内容时才替换
                        #         # 保留原有结构，只替换非页码域的文本内容
                        #         replace_formatted_content(hroot, header_text, header_nsmap)
                        #         print(f"  🔄 替换页眉 ({header_type}): {header_text}")
                        
                        # 应用源文档的段落属性
                        if header_type in source_header_properties:
                            p_elems = hroot.xpath(".//w:p", namespaces=header_nsmap)
                            source_properties = source_header_properties[header_type]
                            for j, p_elem in enumerate(p_elems):
                                if j < len(source_properties):
                                    p_pr_elems = p_elem.xpath("./w:pPr", namespaces=header_nsmap)
                                    if p_pr_elems:
                                        p_pr_elem = p_pr_elems[0]
                                        source_prop = source_properties[j]
                                        
                                        # 应用对齐方式
                                        if source_prop["jc_val"] is not None:
                                            jc_elems = p_pr_elem.xpath("./w:jc", namespaces=header_nsmap)
                                            if jc_elems:
                                                jc_elem = jc_elems[0]
                                            else:
                                                jc_elem = etree.SubElement(p_pr_elem, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}jc")
                                            jc_elem.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", source_prop["jc_val"])
                                        
                                        # 应用制表符停止点
                                        if source_prop["tabs_xml"] is not None:
                                            # 删除现有的制表符停止点
                                            tabs_elems = p_pr_elem.xpath("./w:tabs", namespaces=header_nsmap)
                                            for tabs_elem in tabs_elems:
                                                tabs_elem.getparent().remove(tabs_elem)
                                            # 添加源文档的制表符停止点
                                            try:
                                                tabs_elem = etree.fromstring(source_prop["tabs_xml"])
                                                p_pr_elem.append(tabs_elem)
                                            except Exception as e:
                                                print(f"  ⚠️ 应用制表符停止点时出错: {e}")
                        
                        files[header_file] = etree.tostring(hroot, xml_declaration=True, encoding="UTF-8", standalone="yes")

        # 替换页脚内容
        for i, fref in enumerate(footer_refs):
            rid = fref.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            footer_type = fref.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type", "default")
            
            # 根据关系文件找到 footerX.xml
            rels_name = "word/_rels/document.xml.rels"
            if rels_name in files:
                rels_root = etree.fromstring(files[rels_name])
                footer_target = rels_root.xpath(f".//rel:Relationship[@Id='{rid}']",
                                                namespaces={"rel": "http://schemas.openxmlformats.org/package/2006/relationships"})
                if footer_target:
                    footer_file = "word/" + footer_target[0].get("Target")
                    if footer_file in files:
                        froot = etree.fromstring(files[footer_file])
                        footer_nsmap = froot.nsmap  # 使用页脚文件的命名空间映射
                        footer_text = f"鞍钢集团工程技术有限公司"
                        replace_formatted_content(froot, footer_text, footer_nsmap)
                        print(f"  🔄 替换页脚 ({footer_type}): {footer_text}")
                        # 如果源文档有对应类型的页脚内容，则替换
                        # if footer_type in source_content["footers"]:
                        #     footer_text = f"鞍钢集团工程技术有限公司"
                        #     if footer_text:  # 只有当有内容时才替换
                        #         # 保留原有结构，只替换非页码域的文本内容
                        #         replace_formatted_content(froot, footer_text, footer_nsmap)
                        #         print(f"  🔄 替换页脚 ({footer_type}): {footer_text}")
                        
                        # 应用源文档的段落属性
                        if footer_type in source_footer_properties:
                            p_elems = froot.xpath(".//w:p", namespaces=footer_nsmap)
                            source_properties = source_footer_properties[footer_type]
                            for j, p_elem in enumerate(p_elems):
                                if j < len(source_properties):
                                    p_pr_elems = p_elem.xpath("./w:pPr", namespaces=footer_nsmap)
                                    if p_pr_elems:
                                        p_pr_elem = p_pr_elems[0]
                                        source_prop = source_properties[j]
                                        
                                        # 应用对齐方式
                                        if source_prop["jc_val"] is not None:
                                            jc_elems = p_pr_elem.xpath("./w:jc", namespaces=footer_nsmap)
                                            if jc_elems:
                                                jc_elem = jc_elems[0]
                                            else:
                                                jc_elem = etree.SubElement(p_pr_elem, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}jc")
                                            jc_elem.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", source_prop["jc_val"])
                                        
                                        # 应用制表符停止点
                                        if source_prop["tabs_xml"] is not None:
                                            # 删除现有的制表符停止点
                                            tabs_elems = p_pr_elem.xpath("./w:tabs", namespaces=footer_nsmap)
                                            for tabs_elem in tabs_elems:
                                                tabs_elem.getparent().remove(tabs_elem)
                                            # 添加源文档的制表符停止点
                                            try:
                                                tabs_elem = etree.fromstring(source_prop["tabs_xml"])
                                                p_pr_elem.append(tabs_elem)
                                            except Exception as e:
                                                print(f"  ⚠️ 应用制表符停止点时出错: {e}")
                        
                        files[footer_file] = etree.tostring(froot, xml_declaration=True, encoding="UTF-8", standalone="yes")

        # 保存修改后的文档
        output_path = save_path if save_path else target_docx_path
        output_dir = os.path.dirname(os.path.abspath(output_path))
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
            
        # 先尝试删除目标文件
        if os.path.exists(output_path):
            try:
                os.remove(output_path)
            except Exception as e:
                print(f"  ⚠️ 删除原文件失败: {e}")
        
        # 写入新文件
        with zipfile.ZipFile(output_path, 'w') as zout:
            for name, data in files.items():
                zout.writestr(name, data)
                print(f"  📄 更新文件: {name}")

        print(f"✅ 页眉页脚内容替换完成，文件保存到: {output_path}")
        return True
        
    except Exception as e:
        print(f"❌ 替换页眉页脚内容失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """主函数 - 提供命令行接口"""
    print("🚀 DOCX页眉页脚内容提取与替换工具")
    print("=" * 50)
    print("使用方法:")
    print("1. extract_header_footer_content(docx_path, section_index)")
    print("   - 提取指定节的页眉页脚内容")
    print("2. replace_header_footer_content(source_docx_path, target_docx_path, source_section_index, target_section_index, save_path)")
    print("   - 将源文档的页眉页脚内容替换到目标文档")
    print("=" * 50)


if __name__ == "__main__":
    main()