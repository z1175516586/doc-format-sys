"""
Word文档横竖版处理工具

实现判断Word文档中的横竖版，并将横竖版按连续段落拆分为多个Word文件。
例如：如果文档中是竖版、横版、竖版、横版、竖版的形式，最后会生成5个文档，
分别是竖版的、横版的、竖版的、横版的、竖版的。
如果文档中是竖版、竖版、竖版、横版、竖版，则最后只会生成三个文件，
分别是竖版的、横版的、竖版的。

主要功能：
1. analyze_document_orientation_segments: 分析文档的页面方向段落分布
2. split_document_by_orientation_segments: 按照页面方向的连续段落拆分Word文档
"""

import os
import sys
import shutil
import time
from typing import Optional, List, Dict, Any
from utils.docx_merge import copy_all_to_beginning
from docx import Document
from copy import deepcopy
from lxml import etree
import zipfile
import tempfile

# COM组件可用性检测
try:
    import pythoncom
    import win32com.client as win32
    from win32com.client import constants
    # 使用条件赋值避免重新赋值Final变量
    WinConstants = constants if constants else None
    COM_AVAILABLE = True
except Exception:
    pythoncom = None
    win32 = None
    WinConstants = None
    COM_AVAILABLE = False

# 添加项目路径
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
sys.path.insert(0, parent_dir)

# 定义默认常量值
WD_ORIENT_PORTRAIT = 0
WD_ORIENT_LANDSCAPE = 1
WD_HEADER_FOOTER_PRIMARY = 1
WD_HEADER_FOOTER_FIRST_PAGE = 2
WD_HEADER_FOOTER_EVEN_PAGES = 3


class DocumentOrientationProcessor:
    """文档横竖版处理器"""

    def _count_sections_in_doc(self, doc_path: str, word=None) -> int:
        """
        统计文档节数（使用 COM）。
        """
        if not COM_AVAILABLE or not win32 or not pythoncom:
            return 0
        if not doc_path or not os.path.exists(doc_path):
            return 0

        doc_path = os.path.abspath(doc_path)
        own_word = False
        com_initialized = False
        doc = None

        try:
            if word is None:
                pythoncom.CoInitialize()
                com_initialized = True
                word = win32.DispatchEx("Word.Application")
                word.Visible = False
                word.DisplayAlerts = 0
                try:
                    word.AutomationSecurity = 3
                except Exception:
                    pass
                own_word = True

            doc = word.Documents.Open(doc_path, ReadOnly=True, AddToRecentFiles=False, Visible=False)
            return doc.Sections.Count if doc else 0
        except Exception:
            return 0
        finally:
            if doc is not None:
                try:
                    doc.Close(SaveChanges=0)
                except Exception:
                    pass
            if own_word and word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass


    def __init__(self, temp_dir: Optional[str] = None):
        """初始化处理器"""
        self.temp_dir = temp_dir or os.path.join(parent_dir, 'temp')
        os.makedirs(self.temp_dir, exist_ok=True)

    def _initialize_com_resources(self):
        """
        初始化COM资源

        Returns:
            tuple: (pythoncom_initialized, word_app) - COM初始化状态和Word应用对象
        """
        if not COM_AVAILABLE or not win32 or not pythoncom:
            return False, None

        pythoncom_initialized = False
        word_app = None

        try:
            # 初始化COM库
            pythoncom.CoInitialize()
            pythoncom_initialized = True
            print("  ✅[INFO] COM库成功初始化")

            # 创建Word应用程序对象
            word_app = win32.DispatchEx("Word.Application")
            # word_app = win32.DispatchEx("Kwps.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = False
            print("  ✅[INFO] Word应用程序已成功启动")

            return pythoncom_initialized, word_app
        except Exception as e:
            print(f"  ❌[ERROR] COM资源初始化失败: {str(e)}")
            # 如果初始化失败，清理已初始化的资源
            if pythoncom_initialized and pythoncom is not None:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
            return False, None

    def remove_blank_sections_xml(self, doc_path: str) -> bool:
        """
        删除Word中空白节（XML层面）
        规则：如果某个section内没有任何有效内容（只有sectPr或空段落），则删除该section的sectPr所属段落
        """
        try:
            temp_output = doc_path + ".tmp"

            shutil.copy2(doc_path, temp_output)

            with zipfile.ZipFile(temp_output, 'r') as zin:
                xml_content = zin.read("word/document.xml")
                root = etree.fromstring(xml_content)

                ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
                body = root.find("w:body", ns)

                if body is None:
                    return False

                elements = list(body)

                # 找section边界
                boundaries = []
                start = 0

                sect_indices = []

                for i, elem in enumerate(elements):
                    sectPr = elem.find("./w:pPr/w:sectPr", ns)
                    if sectPr is not None:
                        sect_indices.append(i)

                        boundaries.append((start, i))
                        start = i + 1

                if start < len(elements):
                    boundaries.append((start, len(elements) - 1))

                # 判断每个section是否为空
                remove_indices = set()

                for (s, e) in boundaries:
                    has_content = False

                    for idx in range(s, e + 1):
                        elem = elements[idx]

                        # 跳过sectPr本身
                        if elem.find("./w:pPr/w:sectPr", ns) is not None:
                            continue

                        # 判断是否有文本或非空结构
                        text = "".join(elem.itertext()).strip()
                        if text:
                            has_content = True
                            break

                    # 如果整节没有内容 -> 删除该节最后的sectPr段落
                    if not has_content:
                        for idx in range(e, s - 1, -1):
                            if elements[idx].find("./w:pPr/w:sectPr", ns) is not None:
                                remove_indices.add(idx)
                                break

                # 重建body
                new_body = etree.Element(body.tag, nsmap=body.nsmap)

                for i, elem in enumerate(elements):
                    if i not in remove_indices:
                        new_body.append(deepcopy(elem))

                root.replace(body, new_body)

                new_xml = etree.tostring(
                    root,
                    xml_declaration=True,
                    encoding="UTF-8",
                    standalone="yes"
                )

            # 写回docx
            with zipfile.ZipFile(doc_path, 'w') as zout:
                with zipfile.ZipFile(temp_output, 'r') as zin:
                    for item in zin.infolist():
                        if item.filename == "word/document.xml":
                            continue
                        zout.writestr(item, zin.read(item.filename))

                zout.writestr("word/document.xml", new_xml)

            os.remove(temp_output)
            return True

        except Exception as e:
            print(f"[ERROR] 删除空白节失败: {e}")
            return False


    def _release_com_resources(self, word_app=None, doc=None, pythoncom_initialized=False):
        """
        释放COM资源并等待资源完全释放

        Args:
            word_app: Word应用程序对象
            doc: Word文档对象
            pythoncom_initialized: COM是否已初始化
        """
        try:
            # 关闭文档
            if doc is not None:
                try:
                    doc.Close(SaveChanges=0)  # 不保存更改
                    print("  ✅[INFO] 文档已成功关闭")
                except Exception as e:
                    print(f" ❌[ERROR] 关闭文档时出错: {e}")

            # 退出Word应用程序
            if word_app is not None:
                try:
                    word_app.Quit()
                    print("  ✅[INFO] Word应用程序已成功退出")
                except Exception as e:
                    print(f"  ❌[WARNING] 退出Word应用程序时出错: {e}")

            # 释放COM库资源
            if pythoncom_initialized and pythoncom is not None:
                try:
                    pythoncom.CoUninitialize()
                except Exception as e:
                    print(f"  ❌[WARNING] 释放COM库资源时出错: {e}")
        finally:
            # 添加延迟确保COM资源完全释放
            print("  ✅[INFO] 延迟1秒确保释放中")
            time.sleep(1)

    def _create_orientation_segment_copy_xml(
            self,
            source_content_path: str,
            output_path: str,
            start_section: int,
            end_section: int
    ) -> bool:

        try:

            temp_output = output_path + ".tmp"

            shutil.copy2(source_content_path, temp_output)

            with zipfile.ZipFile(temp_output, 'r') as zin:

                xml_content = zin.read("word/document.xml")

                root = etree.fromstring(xml_content)

                ns = {
                    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                }

                body = root.find("w:body", ns)

                if body is None:
                    return False

                elements = list(body)

                boundaries = self._get_section_boundaries(body)

                if not boundaries:
                    return False

                keep_start = start_section - 1
                keep_end = end_section - 1

                if keep_end >= len(boundaries):
                    return False

                keep_ranges = boundaries[keep_start:keep_end + 1]

                keep_indices = set()

                for s, e in keep_ranges:

                    for idx in range(s, e + 1):
                        keep_indices.add(idx)

                new_body = etree.Element(body.tag, nsmap=body.nsmap)

                for idx in range(len(elements)):
                    if idx in keep_indices:
                        new_body.append(deepcopy(elements[idx]))

                # ⚠️ 必须保证至少有内容
                if len(new_body) == 0:
                    print("[ERROR] 生成段落为空，跳过")
                    return False

                root.replace(body, new_body)

                new_xml = etree.tostring(
                    root,
                    xml_declaration=True,
                    encoding="UTF-8",
                    standalone="yes"
                )

                # 创建新的zip
                with zipfile.ZipFile(output_path, 'w') as zout:

                    for item in zin.infolist():

                        if item.filename == "word/document.xml":
                            continue

                        zout.writestr(item, zin.read(item.filename))

                    # 写入新的document.xml
                    zout.writestr("word/document.xml", new_xml)

            if os.path.exists(temp_output):
                os.remove(temp_output)

            return True

        except Exception as e:

            print(f"[ERROR] XML拆分失败: {e}")

            return False

    def _get_section_boundaries(self, body):
        """
        获取每个 section 的边界节点
        """

        ns = {
            "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        }

        elements = list(body)

        boundaries = []

        current_start = 0

        for i, elem in enumerate(elements):

            sectPr = elem.find("./w:pPr/w:sectPr", ns)

            if sectPr is not None:
                boundaries.append((current_start, i))

                current_start = i + 1

        # 最后一个 section
        if current_start < len(elements):
            boundaries.append((current_start, len(elements) - 1))

        return boundaries

    def analyze_document_orientation_segments(self, doc_path: str) -> List[Dict]:
        """
        基于 XML 分析文档页面方向段落
        不再依赖 COM
        """
        print("--------------------------------------------------")
        print("分析文档页面方向段落(XML模式)")

        try:
            doc = Document(doc_path)

            sections = doc.sections

            total_sections = len(sections)

            print(f"   总节数: {total_sections}")

            section_orientations = []

            for i, section in enumerate(sections):

                sectPr = section._sectPr

                pgSz = sectPr.pgSz

                width = int(pgSz.w)
                height = int(pgSz.h)

                orientation_code = 1 if width > height else 0

                orientation = "横版" if orientation_code == 1 else "竖版"

                section_orientations.append({
                    "section_index": i + 1,
                    "orientation": orientation,
                    "orientation_code": orientation_code
                })

            segments = []

            if section_orientations:

                current_segment = {
                    "start_section": section_orientations[0]["section_index"],
                    "end_section": section_orientations[0]["section_index"],
                    "orientation": section_orientations[0]["orientation"],
                    "orientation_code": section_orientations[0]["orientation_code"]
                }

                for i in range(1, len(section_orientations)):

                    section_info = section_orientations[i]

                    if (
                            section_info["orientation_code"]
                            == current_segment["orientation_code"]
                    ):
                        current_segment["end_section"] = section_info["section_index"]

                    else:

                        segments.append(current_segment)

                        current_segment = {
                            "start_section": section_info["section_index"],
                            "end_section": section_info["section_index"],
                            "orientation": section_info["orientation"],
                            "orientation_code": section_info["orientation_code"]
                        }

                segments.append(current_segment)

            print(f"   段落数: {len(segments)}")

            for i, segment in enumerate(segments):

                if segment["start_section"] == segment["end_section"]:

                    print(
                        f"     段落 {i + 1}: "
                        f"第{segment['start_section']}节 "
                        f"({segment['orientation']})"
                    )

                else:

                    print(
                        f"     段落 {i + 1}: "
                        f"第{segment['start_section']}节 - "
                        f"第{segment['end_section']}节 "
                        f"({segment['orientation']})"
                    )

            return segments

        except Exception as e:

            print(f"分析页面方向段落时发生错误: {str(e)}")

            return []

    ## 合并相关
    def _merge_documents_xml(
            self,
            base_file: str,
            file_to_append: str,
            output_file: str
    ) -> bool:
        """
        基于 XML 的 Word 文档合并
        完全不依赖 COM
        """

        try:

            temp_output = output_file + ".tmp"

            shutil.copy2(base_file, temp_output)

            with zipfile.ZipFile(temp_output, 'r') as base_zip, \
                    zipfile.ZipFile(file_to_append, 'r') as append_zip:

                # 读取 document.xml
                base_xml = etree.fromstring(
                    base_zip.read("word/document.xml")
                )

                append_xml = etree.fromstring(
                    append_zip.read("word/document.xml")
                )

                ns = {
                    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                }

                base_body = base_xml.find("w:body", ns)
                append_body = append_xml.find("w:body", ns)

                if base_body is None or append_body is None:
                    return False

                # ---------------------------------------------------
                # 获取 base 最后的 sectPr
                # Word 最后一个 sectPr 是文档级别 sectPr
                # ---------------------------------------------------

                base_elements = list(base_body)

                last_sectPr = None

                for elem in reversed(base_elements):

                    # 只找 body-level sectPr
                    if elem.tag == "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr":
                        last_sectPr = deepcopy(elem)
                        break

                # ---------------------------------------------------
                # 删除 base 最后的 body-level sectPr
                # 避免重复
                # ---------------------------------------------------

                for elem in reversed(base_elements):

                    if (
                            elem.tag ==
                            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr"
                    ):
                        base_body.remove(elem)
                        break

                # ---------------------------------------------------
                # 插入分页 section break
                # ---------------------------------------------------

                if last_sectPr is not None:
                    p = etree.Element(
                        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"
                    )

                    pPr = etree.SubElement(
                        p,
                        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr"
                    )

                    pPr.append(deepcopy(last_sectPr))

                    base_body.append(p)

                # ---------------------------------------------------
                # 追加 append_body 内容
                # 注意：
                # append 最后一个 body-level sectPr
                # 必须保留
                # ---------------------------------------------------

                append_elements = list(append_body)

                for elem in append_elements:

                    # 跳过 body-level sectPr
                    # 只跳过 body-level sectPr
                    elem_copy = deepcopy(elem)
                    base_body.append(elem_copy)

                    # elem_copy = deepcopy(elem)

                    # 删除段落内部多余 sectPr
                    # 删除段落内部多余 sectPr
                    # sectPrs = elem_copy.findall(
                    #     ".//w:sectPr",
                    #     ns
                    # )

                    # 只保留真正的 section break
                    # 删除 header/footer 引用
                    # for sectPr in sectPrs:
                    #
                    #     for child in list(sectPr):
                    #
                    #         tag_name = child.tag.split("}")[-1]
                    #
                    #         if tag_name in [
                    #             "headerReference",
                    #             "footerReference"
                    #         ]:
                    #             sectPr.remove(child)
                    #
                    # # 只保留真正的 section break
                    # # 删除 header/footer 引用
                    # for sectPr in sectPrs:
                    #
                    #     for child in list(sectPr):
                    #
                    #         tag_name = child.tag.split("}")[-1]
                    #
                    #         if tag_name in [
                    #             "headerReference",
                    #             "footerReference"
                    #         ]:
                    #             sectPr.remove(child)

                    base_body.append(elem_copy)

                # ---------------------------------------------------
                # 重新设置唯一 body-level sectPr（必须保留，否则文档损坏）
                # ---------------------------------------------------

                # ---------------------------------------------------
                # 把原 base 的 body-level sectPr
                # 转成 paragraph-level sectPr
                # 作为 section break 插入
                # ---------------------------------------------------

                if last_sectPr is not None:

                    p = etree.Element(
                        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"
                    )

                    pPr = etree.SubElement(
                        p,
                        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr"
                    )

                    clean_sectPr = deepcopy(last_sectPr)

                    for child in list(clean_sectPr):

                        tag_name = child.tag.split("}")[-1]

                        if tag_name in [
                            "headerReference",
                            "footerReference"
                        ]:
                            clean_sectPr.remove(child)

                    pPr.append(clean_sectPr)

                    base_body.append(p)

                else:
                    # ⚠️ 兜底：必须补一个最基础 sectPr，否则 Word 直接坏
                    sectPr = etree.Element("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr")
                    base_body.append(sectPr)

                # ---------------------------------------------------
                # 写回 XML
                # ---------------------------------------------------

                new_xml = etree.tostring(
                    base_xml,
                    xml_declaration=True,
                    encoding="UTF-8",
                    standalone="yes"
                )

                with zipfile.ZipFile(output_file, 'w') as zout:

                    for item in base_zip.infolist():

                        if item.filename == "word/document.xml":
                            continue

                        zout.writestr(
                            item,
                            base_zip.read(item.filename)
                        )

                    zout.writestr(
                        "word/document.xml",
                        new_xml
                    )

            if os.path.exists(temp_output):
                os.remove(temp_output)

            return True

        except Exception as e:

            print(f"[ERROR] XML合并失败: {e}")

            return False

    def _merge_documents_default(self, base_file: str, file_to_append: str, output_file: str) -> bool:
        """
        默认文档合并方法

        Args:
            base_file: 基础文档路径
            file_to_append: 待追加文档路径
            output_file: 输出文档路径

        Returns:
            bool: 合并是否成功
        """
        try:
            # 使用docx_merge.py中的方法进行文档合并
            copy_all_to_beginning(base_file, file_to_append, output_file)
            return True
        except Exception as e:
            print(f"  ❌[ERROR] 默认合并方法失败: {str(e)}")
            return False

    def split_document_by_orientation_segments(self, source_file: str, output_dir: Optional[str] = None) -> List[Dict]:
        """
        按照页面方向的连续段落拆分Word文档

        Args:
            source_file: 源文档路径
            output_dir: 输出目录路径（可选，默认使用临时目录）

        Returns:
            list: 包含拆分文件信息的列表，每个元素为字典 {"file_path": str, "orientation": str, "index": int, "sections": str}
        """
        print("-" * 50)
        print("[INFO] 按页面方向段落拆分文档")

        # 分析文档的页面方向段落
        segments = self.analyze_document_orientation_segments(source_file)

        try:
            # 如果未指定输出目录，则使用临时目录
            if output_dir is None:
                output_dir = self.temp_dir

            # 确保输出目录存在
            os.makedirs(output_dir, exist_ok=True)

            # 获取源文件基础名称
            base_name = os.path.splitext(os.path.basename(source_file))[0]

            if not segments:
                print("[WARNING] 未检测到有效的页面方向段落，返回完整文档")
                default_path = os.path.join(output_dir, f"{base_name}_完整文档.docx")
                shutil.copy2(source_file, default_path)
                return [{
                    "file_path": default_path,
                    "orientation": "完整文档",
                    "index": 0,
                    "sections": "全部"
                }]

            # 存储拆分结果
            split_files = []

            # 为每个段落创建独立文档
            for i, segment in enumerate(segments):
                # 创建文件名
                orientation_name = segment["orientation"]
                start_section = segment["start_section"]
                end_section = segment["end_section"]
                sections_desc = f"第{start_section}节-第{end_section}节" if start_section != end_section else f"第{start_section}节"

                output_path = os.path.join(output_dir, f"{base_name}_段落{i + 1}_{orientation_name}.docx")

                # 创建段落文档
                success = self._create_orientation_segment_copy_xml(
                    source_content_path=source_file,
                    output_path=output_path,
                    start_section=start_section,
                    end_section=end_section
                )

                if success and os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                    self.remove_blank_sections_xml(output_path)
                    split_files.append({
                        "file_path": output_path,
                        "orientation": orientation_name,
                        "index": i,
                        "sections": sections_desc
                    })
                    print(
                        f"[SUCCESS] 已保存段落文件: {os.path.basename(output_path)} ({sections_desc}, {orientation_name})")
                elif os.path.exists(output_path):
                    # 删除空文件
                    try:
                        os.remove(output_path)
                    except Exception:
                        pass

            # 打印拆分结果
            print(f"\n[INFO] 拆分结果:")
            for item in split_files:
                print(
                    f"   {item['index'] + 1}. {os.path.basename(item['file_path'])} ({item['sections']}, {item['orientation']})")

            return split_files

        except Exception as e:
            print(f"  ❌[ERROR] 按页面方向段落拆分文档过程中发生错误: {str(e)}")
            # 发生错误时，返回包含原始文件的列表
            return [{
                "file_path": source_file,
                "orientation": "未知",
                "index": 0,
                "sections": "全部"
            }]

    ## 合并文档
    # def merge_documents_with_formatting(self, base_file: str, file_to_append: str, output_file: str) -> bool:
    #     """
    #     将 file_to_append 追加到 base_file 末尾，保持各自的页面方向格式
    #
    #     Args:
    #         base_file: 基础文档路径
    #         file_to_append: 待追加文档路径
    #         output_file: 输出文档路径
    #
    #     Returns:
    #         bool: 合并是否成功
    #     """
    #     print("-" * 50)
    #     print("[INFO] 合并文档（保持页面方向）")
    #
    #     # 检查输入文件是否存在
    #     if not os.path.exists(base_file):
    #         print(f"  ❌[ERROR] 基础文档不存在: {base_file}")
    #         return False
    #
    #     if not os.path.exists(file_to_append):
    #         print(f"  ❌[ERROR] 待追加文档不存在: {file_to_append}")
    #         return False
    #
    #     # 确保输出目录存在
    #     output_dir = os.path.dirname(output_file)
    #     if output_dir:
    #         os.makedirs(output_dir, exist_ok=True)
    #
    #     # 分析待追加文档的页面方向
    #     append_segments = self.analyze_document_orientation_segments(file_to_append)
    #
    #     # 如果无法分析待追加文档的方向，则使用默认的合并方法
    #     if not append_segments:
    #         print("[WARNING] 无法分析待追加文档的页面方向，使用默认合并方法")
    #         print("[INFO] 使用默认合并方法")
    #         return self._merge_documents_default(base_file, file_to_append, output_file)
    #
    #     # 获取待追加文档的主要方向（第一个段落的方向）
    #     primary_orientation = append_segments[0]["orientation_code"]
    #     is_landscape = primary_orientation == 1  # 1 表示横向
    #
    #     print(f"[INFO] 基础文档: {os.path.basename(base_file)}")
    #     print(f"[INFO] 待追加文档: {os.path.basename(file_to_append)} (主要方向: {'横向' if is_landscape else '纵向'})")
    #     print(f"[INFO] 输出文档: {os.path.basename(output_file)}")
    #
    #     print("[INFO] 使用 XML 合并模式")
    #     return self._merge_documents_xml(
    #         base_file=base_file,
    #         file_to_append=file_to_append,
    #         output_file=output_file
    #     )

    def merge_documents_with_formatting(self, base_file: str, file_to_append: str, output_file: str) -> bool:
        """
        将 file_to_append 追加到 base_file 末尾，保持各自的页面方向格式

        Args:
            base_file: 基础文档路径
            file_to_append: 待追加文档路径
            output_file: 输出文档路径

        Returns:
            bool: 合并是否成功
        """
        print("-" * 50)
        print("[INFO] 合并文档（保持页面方向）")

        # 检查输入文件是否存在
        if not os.path.exists(base_file):
            print(f"  ❌[ERROR] 基础文档不存在: {base_file}")
            return False

        if not os.path.exists(file_to_append):
            print(f"  ❌[ERROR] 待追加文档不存在: {file_to_append}")
            return False

        # 确保输出目录存在
        output_dir = os.path.dirname(output_file)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)

        # 分析待追加文档的页面方向
        append_segments = self.analyze_document_orientation_segments(file_to_append)

        # 如果无法分析待追加文档的方向，则使用默认的合并方法
        if not append_segments:
            print("[WARNING] 无法分析待追加文档的页面方向，使用默认合并方法")
            print("[INFO] 使用默认合并方法")
            return self._merge_documents_default(base_file, file_to_append, output_file)

        # 获取待追加文档的主要方向（第一个段落的方向）
        primary_orientation = append_segments[0]["orientation_code"]
        is_landscape = primary_orientation == 1  # 1 表示横向

        print(f"[INFO] 基础文档: {os.path.basename(base_file)}")
        print(f"[INFO] 待追加文档: {os.path.basename(file_to_append)} (主要方向: {'横向' if is_landscape else '纵向'})")
        print(f"[INFO] 输出文档: {os.path.basename(output_file)}")

        # 使用专门的合并方法保持页面方向
        print("[INFO] 使用专门的合并方法保持页面方向")
        return self._merge_with_orientation_preserved(base_file, file_to_append, output_file, is_landscape)

    def _merge_with_orientation_preserved(self, base_file: str, file_to_append: str, output_file: str,
                                          is_landscape: bool) -> bool:
        """
        保持页面方向的文档合并方法

        Args:
            base_file: 基础文档路径
            file_to_append: 待追加文档路径
            output_file: 输出文档路径
            is_landscape: 待追加文档是否为横向

        Returns:
            bool: 合并是否成功
        """
        # 分析待追加文档的页面方向段落分布
        append_segments = self.analyze_document_orientation_segments(file_to_append)

        if not append_segments:
            print("[WARNING] 无法分析待追加文档的页面方向")
            print("[INFO] 使用默认合并方法")
            return self._merge_documents_default(base_file, file_to_append, output_file)

        # 使用自定义的合并方法来保持页面方向
        print("[INFO] 用自定义的合并方法来保持页面方向")
        return self._custom_merge_with_orientation(base_file, file_to_append, output_file, append_segments)

    def _custom_merge_with_orientation(self, base_file: str, file_to_append: str, output_file: str,
                                       append_segments: list) -> bool:
        """
        自定义的保持页面方向合并方法

        Args:
            base_file: 基础文档路径
            file_to_append: 待追加文档路径
            output_file: 输出文档路径
            append_segments: 待追加文档的页面方向段落信息

        Returns:
            bool: 合并是否成功
        """
        # 确保输出目录存在
        output_dir = os.path.dirname(output_file)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)

        # 复制基础文件到输出位置
        shutil.copy2(base_file, output_file)

        # 使用COM组件进行文档处理
        if not COM_AVAILABLE or not win32 or not pythoncom:
            print("  ❌[ERROR] COM组件不可用，无法进行页面方向保持合并")
            return False

        pythoncom_initialized = False
        word_app = None
        doc_base = None

        try:
            # 初始化COM资源
            pythoncom_initialized, word_app = self._initialize_com_resources()
            if not pythoncom_initialized or word_app is None:
                return False

            # 打开基础文档
            doc_base = word_app.Documents.Open(output_file, ReadOnly=False)

            # 记录插入前的节(Section)数量
            sections_before = doc_base.Sections.Count
            print(f"基础文档包含 {sections_before} 个节。")

            # 定义常量
            wdSectionBreakNextPage = 2
            wdCollapseEnd = 0
            wdOrientPortrait = getattr(WinConstants, "wdOrientPortrait", 0) if WinConstants else 0
            wdOrientLandscape = getattr(WinConstants, "wdOrientLandscape", 1) if WinConstants else 1
            wdHeaderFooterPrimary = getattr(WinConstants, "wdHeaderFooterPrimary", 1) if WinConstants else 1
            wdHeaderFooterFirstPage = getattr(WinConstants, "wdHeaderFooterFirstPage", 2) if WinConstants else 2
            wdHeaderFooterEvenPages = getattr(WinConstants, "wdHeaderFooterEvenPages", 3) if WinConstants else 3

            # 移动到文档末尾并插入分节符
            print("移动到文档末尾并插入分节符...")
            end_range = doc_base.Content
            end_range.Collapse(wdCollapseEnd)
            end_range.InsertBreak(Type=wdSectionBreakNextPage)

            # 在新节的开头插入待追加文件
            print(f"正在末尾追加文件: {os.path.basename(file_to_append)}")
            end_range.InsertFile(
                FileName=file_to_append,
                Link=False,
                ConfirmConversions=False
            )

            # 记录插入后的节(Section)数量
            sections_after = doc_base.Sections.Count
            print(f"插入后文档变为 {sections_after} 个节。")

            # 遍历所有新添加的节并设置正确的页面方向
            start_section_index = sections_before + 1
            print(f"将从第 {start_section_index} 节开始设置页面方向...")

            # 根据待追加文档的段落信息设置页面方向
            section_index = start_section_index
            for i, segment in enumerate(append_segments):
                start_section = segment["start_section"]
                end_section = segment["end_section"]
                orientation_code = segment["orientation_code"]

                # 计算这个段落包含的节的数量
                segment_section_count = end_section - start_section + 1

                print(
                    f"  处理段落 {i + 1}: 第{start_section}节 - 第{end_section}节 ({segment['orientation']})，共{segment_section_count}节")

                # 为这个段落的所有节设置正确的页面方向
                for j in range(segment_section_count):
                    if section_index <= sections_after:
                        try:
                            new_section = doc_base.Sections(section_index)

                            # 设置页面方向
                            new_section.PageSetup.Orientation = orientation_code
                            print(f"    > 已设置第 {section_index} 节为{segment['orientation']}。")

                            # 强制断开所有页眉页脚链接
                            try:
                                new_section.Headers(wdHeaderFooterPrimary).LinkToPrevious = False
                                new_section.Headers(wdHeaderFooterFirstPage).LinkToPrevious = False
                                new_section.Headers(wdHeaderFooterEvenPages).LinkToPrevious = False
                                new_section.Footers(wdHeaderFooterPrimary).LinkToPrevious = False
                                new_section.Footers(wdHeaderFooterFirstPage).LinkToPrevious = False
                                new_section.Footers(wdHeaderFooterEvenPages).LinkToPrevious = False
                                print(f"    > 已断开第 {section_index} 节的页眉/页脚链接。")
                            except Exception as link_error:
                                print(
                                    f"    [INFO] 断开第 {section_index} 节链接时出现非致命错误 (已忽略): {link_error}")

                        except Exception as section_error:
                            print(f"    ❌[ERROR] 处理第 {section_index} 节时出现错误: {section_error}")

                        section_index += 1

            # 保存文档
            print(f"正在保存到: {output_file}")
            doc_base.Save()
            print("合并成功！")
            return True

        except Exception as e:
            print(f"  ❌[ERROR] 合并过程中发生错误: {str(e)}")
            return False

        finally:
            # 释放COM资源
            self._release_com_resources(word_app, doc_base, pythoncom_initialized)

    def _normalize_section_orientation_by_segments(
        self,
        doc_path: str,
        segments: List[Dict[str, Any]],
        word=None
    ) -> bool:
        """
        根据 segments 顺序强制校正最终文档每个节的方向。
        """
        if not COM_AVAILABLE or not win32 or not pythoncom or WinConstants is None:
            print("⚠️ COM 不可用，无法校正节方向")
            return False
        if not doc_path or not os.path.exists(doc_path):
            print("⚠️ 文档不存在，无法校正节方向")
            return False
        if not segments:
            print("⚠️ segments 为空，跳过节方向校正")
            return False

        own_word = False
        com_initialized = False
        doc = None

        try:
            if word is None:
                pythoncom.CoInitialize()
                com_initialized = True
                word = win32.DispatchEx("Word.Application")
                word.Visible = False
                word.DisplayAlerts = 0
                try:
                    word.AutomationSecurity = 3
                except Exception:
                    pass
                own_word = True

            expected_orientations: List[int] = []
            for seg in segments:
                seg_path = seg.get("file_path")
                orientation = seg.get("orientation", "竖版")
                if not seg_path or not os.path.exists(seg_path):
                    continue
                section_count = self._count_sections_in_doc(seg_path, word=word)
                if section_count <= 0:
                    section_count = 1
                orientation_code = (
                    getattr(WinConstants, "wdOrientLandscape", 1)
                    if orientation == "横版"
                    else getattr(WinConstants, "wdOrientPortrait", 0)
                )
                expected_orientations.extend([orientation_code] * section_count)

            if not expected_orientations:
                print("⚠️ 未生成期望方向序列，跳过校正")
                return False

            doc = word.Documents.Open(doc_path, ReadOnly=False)
            total_sections = doc.Sections.Count if doc else 0
            if total_sections <= 0:
                return False

            apply_count = min(total_sections, len(expected_orientations))
            for idx in range(1, apply_count + 1):
                try:
                    section = doc.Sections(idx)
                    section.PageSetup.Orientation = expected_orientations[idx - 1]
                except Exception:
                    continue

            if total_sections != len(expected_orientations):
                print(
                    f"⚠️ 节数量不一致: 实际={total_sections}, 期望={len(expected_orientations)}，已按前 {apply_count} 节校正"
                )
            else:
                print(f"✅ 已校正全部 {apply_count} 个节的方向")

            doc.Save()
            return True
        except Exception as exc:
            print(f"⚠️ 校正节方向失败: {exc}")
            return False
        finally:
            if doc is not None:
                try:
                    doc.Close(SaveChanges=0)
                except Exception:
                    pass
            if own_word and word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass


def main():
    """主函数：测试拆分 + XML合并"""

    # ---------------------------------------------------
    # 原始文档
    # ---------------------------------------------------

    source_file = r"E:\新建文件夹\tmpctwyqtk9.docx"

    if not os.path.exists(source_file):
        print(f"[ERROR] 源文件不存在: {source_file}")
        return

    # ---------------------------------------------------
    # 输出目录
    # ---------------------------------------------------

    output_dir = os.path.join(
        "E:",
        "新建文件夹",
        "orientation_segment_split_results"
    )

    os.makedirs(output_dir, exist_ok=True)

    # ---------------------------------------------------
    # 创建处理器
    # ---------------------------------------------------

    processor = DocumentOrientationProcessor()

    # ---------------------------------------------------
    # 1. 拆分
    # ---------------------------------------------------

    print("\n================ 拆分测试 ================\n")

    split_files = processor.split_document_by_orientation_segments(
        source_file=source_file,
        output_dir=output_dir
    )

    if not split_files:
        print("[ERROR] 拆分失败")
        return

    print("\n拆分结果：")

    for item in split_files:

        print(
            f"  {item['index'] + 1}. "
            f"{os.path.basename(item['file_path'])} "
            f"({item['orientation']}) "
            f"[{item['sections']}]"
        )

    # ---------------------------------------------------
    # 2. 合并
    # ---------------------------------------------------

    print("\n================ 合并测试 ================\n")

    merged_output = os.path.join(
        output_dir,
        "merged_result.docx"
    )

    # 先复制第一个文件作为基础文件
    shutil.copy2(
        split_files[0]["file_path"],
        merged_output
    )

    print(f"[INFO] 初始合并文件:")
    print(f"  {merged_output}")

    # 依次追加后续文件
    for i in range(1, len(split_files)):

        append_file = split_files[i]["file_path"]

        temp_merged = os.path.join(
            output_dir,
            f"temp_merge_{i}.docx"
        )

        print(f"\n[INFO] 正在追加:")
        print(f"  {os.path.basename(append_file)}")

        success = processor.merge_documents_with_formatting(
            base_file=merged_output,
            file_to_append=append_file,
            output_file=temp_merged
        )

        if not success:
            print(f"[ERROR] 合并失败: {append_file}")
            return

        # 替换 merged_output
        os.remove(merged_output)

        shutil.move(
            temp_merged,
            merged_output
        )

    # ---------------------------------------------------
    # 完成
    # ---------------------------------------------------

    print("\n========================================")
    print("[SUCCESS] 全部测试完成")
    print("========================================")

    print(f"\n原始文件:")
    print(f"  {source_file}")

    print(f"\n拆分目录:")
    print(f"  {output_dir}")

    print(f"\n最终合并文件:")
    print(f"  {merged_output}")

    # ---------------------------------------------------
    # 3. 全局方向校正
    # ---------------------------------------------------

    print("\n================ 方向校正测试 ================\n")

    normalize_success = processor._normalize_section_orientation_by_segments(
        doc_path=merged_output,
        segments=split_files
    )

    if normalize_success:

        print("[SUCCESS] 方向校正成功")

    else:

        print("[WARNING] 方向校正失败")

    # ---------------------------------------------------
    # 4. COM打开测试（检测文档是否损坏）
    # ---------------------------------------------------

    print("\n================ COM打开测试 ================\n")

    try:

        import pythoncom
        import win32com.client as win32

        pythoncom.CoInitialize()

        word = win32.DispatchEx("Word.Application")

        word.Visible = False
        word.DisplayAlerts = False

        print("[INFO] 正在使用 Word 打开合并后的文档...")

        doc = word.Documents.Open(
            os.path.abspath(merged_output),
            ReadOnly=False
        )

        print("[SUCCESS] Word 成功打开文档")

        # 再保存一次
        doc.Save()

        print("[SUCCESS] Word Save 成功")

        # 关闭文档
        doc.Close(False)

        print("[SUCCESS] Word Close 成功")

        # 退出 Word
        word.Quit()

        pythoncom.CoUninitialize()

        print("[SUCCESS] COM测试通过，文档未损坏")

    except Exception as e:

        print(f"[ERROR] COM打开测试失败: {e}")

def main1():
    """测试：依次合并 test 目录下所有 docx 文件"""

    test_dir = r"E:\新建文件夹\test"

    if not os.path.exists(test_dir):
        print(f"[ERROR] 目录不存在: {test_dir}")
        return

    # ---------------------------------------------------
    # 获取所有 docx
    # ---------------------------------------------------

    docx_files = sorted([
        os.path.join(test_dir, f)
        for f in os.listdir(test_dir)
        if (
                f.lower().endswith(".docx")
                and not f.startswith("merged_")
                and not f.startswith("temp_merge_")
        )
    ])

    if not docx_files:
        print("[ERROR] 未找到 docx 文件")
        return

    print("\n================ 待合并文件 ================\n")

    for i, f in enumerate(docx_files, 1):
        print(f"{i}. {os.path.basename(f)}")

    # ---------------------------------------------------
    # 输出文件
    # ---------------------------------------------------

    merged_output = os.path.join(
        test_dir,
        "merged_all.docx"
    )

    # ---------------------------------------------------
    # 创建处理器
    # ---------------------------------------------------

    processor = DocumentOrientationProcessor()

    # ---------------------------------------------------
    # 第一个文件作为基础文件
    # ---------------------------------------------------

    shutil.copy2(
        docx_files[0],
        merged_output
    )

    print(f"\n[INFO] 初始文件:")
    print(f"  {os.path.basename(docx_files[0])}")

    # ---------------------------------------------------
    # 依次合并
    # ---------------------------------------------------

    for i in range(1, len(docx_files)):

        append_file = docx_files[i]

        temp_output = os.path.join(
            test_dir,
            f"temp_merge_{i}.docx"
        )

        print(f"\n[INFO] 正在合并:")
        print(f"  {os.path.basename(append_file)}")

        success = processor.merge_documents_with_formatting(
            base_file=merged_output,
            file_to_append=append_file,
            output_file=temp_output
        )

        if not success:
            print(f"[ERROR] 合并失败: {append_file}")
            return

        # 替换主文件
        os.remove(merged_output)

        shutil.move(
            temp_output,
            merged_output
        )

    # ---------------------------------------------------
    # COM 打开测试
    # ---------------------------------------------------

    print("\n================ COM打开测试 ================\n")

    try:

        import pythoncom
        import win32com.client as win32

        pythoncom.CoInitialize()

        word = win32.DispatchEx("Word.Application")

        word.Visible = False
        word.DisplayAlerts = False

        doc = word.Documents.Open(
            os.path.abspath(merged_output),
            ReadOnly=False
        )

        print("[SUCCESS] Word 成功打开合并文档")

        doc.Save()

        print("[SUCCESS] Word Save 成功")

        doc.Close(False)

        word.Quit()

        pythoncom.CoUninitialize()

        print("[SUCCESS] COM测试通过")

    except Exception as e:

        print(f"[ERROR] COM打开失败: {e}")

    print("\n========================================")
    print("[SUCCESS] 全部合并完成")
    print("========================================")

    print(f"\n输出文件:")
    print(f"  {merged_output}")

if __name__ == "__main__":
    main()
