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
from typing import Optional, List, Dict
from utils.docx_merge import copy_all_to_beginning

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

class DocumentOrientationProcessor:
    """文档横竖版处理器"""
    
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
    
    def analyze_document_orientation_segments(self, doc_path: str) -> List[Dict]:
        """
        分析文档的页面方向段落
        
        Args:
            doc_path: 文档路径
            
        Returns:
            List[Dict]: 页面方向段落信息列表
        """
        print("--------------------------------------------------")
        print("分析文档页面方向段落")
        
        pythoncom_initialized = False
        word_app = None
        doc = None
        
        try:
            # 初始化COM资源
            pythoncom_initialized, word_app = self._initialize_com_resources()
            if not pythoncom_initialized or word_app is None:
                print("  ❌[ERROR] COM资源初始化失败")
                return []
            
            # 打开文档
            doc = word_app.Documents.Open(doc_path)
            
            sections = doc.Sections
            total_sections = sections.Count
            print(f"   总节数: {total_sections}")
            
            # 分析每个节的方向
            section_orientations = []
            for i in range(1, total_sections + 1):
                section = sections(i)
                page_setup = section.PageSetup
                
                # 判断页面方向
                if page_setup.Orientation == 1:  # wdOrientLandscape
                    orientation = "横版"
                    orientation_code = 1
                else:  # wdOrientPortrait
                    orientation = "竖版"
                    orientation_code = 0
                    
                section_orientations.append({
                    "section_index": i,
                    "orientation": orientation,
                    "orientation_code": orientation_code  # 添加orientation_code字段
                })
            
            # 按方向分组相邻的节
            segments = []
            if section_orientations:
                current_segment = {
                    "start_section": section_orientations[0]["section_index"],
                    "end_section": section_orientations[0]["section_index"],
                    "orientation": section_orientations[0]["orientation"],
                    "orientation_code": section_orientations[0]["orientation_code"]  # 添加orientation_code字段
                }
                
                for i in range(1, len(section_orientations)):
                    section_info = section_orientations[i]
                    
                    # 如果方向相同且节号连续，则扩展当前段落
                    if (section_info["orientation"] == current_segment["orientation"] and 
                        section_info["section_index"] == current_segment["end_section"] + 1):
                        current_segment["end_section"] = section_info["section_index"]
                    else:
                        # 方向不同或不连续，保存当前段落并开始新段落
                        segments.append(current_segment)
                        current_segment = {
                            "start_section": section_info["section_index"],
                            "end_section": section_info["section_index"],
                            "orientation": section_info["orientation"],
                            "orientation_code": section_info["orientation_code"]  # 添加orientation_code字段
                        }
                
                # 添加最后一个段落
                segments.append(current_segment)
            
            print(f"   段落数: {len(segments)}")
            for i, segment in enumerate(segments):
                if segment["start_section"] == segment["end_section"]:
                    print(f"     段落 {i+1}: 第{segment['start_section']}节 ({segment['orientation']})")
                else:
                    print(f"     段落 {i+1}: 第{segment['start_section']}节 - 第{segment['end_section']}节 ({segment['orientation']})")
            
            return segments
            
        except Exception as e:
            print(f"分析页面方向段落时发生错误: {str(e)}")
            return []
        finally:
            # 释放COM资源
            self._release_com_resources(word_app, doc, pythoncom_initialized)
    
    ## 拆分文档

    def _create_orientation_segment_copy(self, source_content_path: str, output_path: str, start_section: int, end_section: int) -> bool:
        """
        使用Word COM仅保留指定节范围并另存为新文档
        
        Args:
            source_content_path: 源文档路径
            output_path: 输出文档路径
            start_section: 起始节（从1开始）
            end_section: 结束节（从1开始）
            
        Returns:
            bool: 是否成功创建
        """
        if not COM_AVAILABLE or not win32 or not pythoncom:
            return False

        pythoncom_initialized = False
        word_app = None
        cloned_doc = None
        remaining_sections = 0  # 初始化变量

        source_content_path = os.path.abspath(source_content_path)
        output_path = os.path.abspath(output_path)

        if not os.path.exists(source_content_path):
            return False

        try:
            shutil.copy2(source_content_path, output_path)
        except Exception as copy_error:
            print(f"[WARNING] 正文克隆失败: {copy_error}")
            return False

        try:
            # 初始化COM资源
            pythoncom_initialized, word_app = self._initialize_com_resources()
            if not pythoncom_initialized or word_app is None:
                return False

            cloned_doc = word_app.Documents.Open(output_path, ReadOnly=False)
            sections_count = cloned_doc.Sections.Count if cloned_doc else 0

            # 删除指定范围之外的节
            # 先删除后面的节（从后往前删除避免索引变化）
            for index in range(sections_count, end_section, -1):
                if index <= cloned_doc.Sections.Count:
                    section = cloned_doc.Sections(index)
                    section.Range.Delete()

            # 再删除前面的节（从前往后删除避免索引变化）
            for index in range(start_section - 1, 0, -1):
                if cloned_doc.Sections.Count > 0:
                    section = cloned_doc.Sections(1)  # 总是删除第一节
                    section.Range.Delete()

            remaining_sections = cloned_doc.Sections.Count if cloned_doc else 0

            if remaining_sections > 0:
                # 使用WinConstants或默认值
                wdHeaderFooterPrimary = getattr(WinConstants, "wdHeaderFooterPrimary", WD_HEADER_FOOTER_PRIMARY) if WinConstants else WD_HEADER_FOOTER_PRIMARY
                wdHeaderFooterFirstPage = getattr(WinConstants, "wdHeaderFooterFirstPage", WD_HEADER_FOOTER_FIRST_PAGE) if WinConstants else WD_HEADER_FOOTER_FIRST_PAGE
                wdHeaderFooterEvenPages = getattr(WinConstants, "wdHeaderFooterEvenPages", WD_HEADER_FOOTER_EVEN_PAGES) if WinConstants else WD_HEADER_FOOTER_EVEN_PAGES
                
                header_types = [
                    wdHeaderFooterPrimary,
                    wdHeaderFooterFirstPage,
                    wdHeaderFooterEvenPages,
                ]
                footer_types = [
                    wdHeaderFooterPrimary,
                    wdHeaderFooterFirstPage,
                    wdHeaderFooterEvenPages,
                ]
                try:
                    for sec_index in range(1, remaining_sections + 1):
                        section = cloned_doc.Sections(sec_index)
                        for header_type in header_types:
                            try:
                                header = section.Headers(header_type)
                                if header is not None and getattr(header, "Exists", True):
                                    header.LinkToPrevious = False
                            except Exception:
                                continue
                        for footer_type in footer_types:
                            try:
                                footer = section.Footers(footer_type)
                                if footer is not None and getattr(footer, "Exists", True):
                                    footer.LinkToPrevious = False
                            except Exception:
                                continue
                except Exception as link_error:
                    print(f"[INFO] 取消页眉页脚链接时出现问题: {link_error}")

            content_length = cloned_doc.Content.StoryLength if cloned_doc else 0

            if remaining_sections > 0 and content_length > 1:
                cloned_doc.Save()
                return True

            if os.path.exists(output_path):
                try:
                    os.remove(output_path)
                except Exception:
                    pass

            return False
        except Exception as exc:
            print(f"[WARNING] 页面方向段落内容裁剪失败: {exc}")
            return False
        finally:
            # 释放COM资源
            self._release_com_resources(word_app, cloned_doc, pythoncom_initialized)

            if (remaining_sections <= 0 or not os.path.exists(output_path) or os.path.getsize(output_path) == 0):
                if os.path.exists(output_path):
                    try:
                        os.remove(output_path)
                    except Exception:
                        pass

    ## 合并相关
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

    def _merge_with_orientation_preserved(self, base_file: str, file_to_append: str, output_file: str, is_landscape: bool) -> bool:
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

    def _custom_merge_with_orientation(self, base_file: str, file_to_append: str, output_file: str, append_segments: list) -> bool:
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
                
                print(f"  处理段落 {i+1}: 第{start_section}节 - 第{end_section}节 ({segment['orientation']})，共{segment_section_count}节")
                
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
                                print(f"    [INFO] 断开第 {section_index} 节链接时出现非致命错误 (已忽略): {link_error}")
                                
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

    ## 直接使用的方法
    ## 拆分文档
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
                
                output_path = os.path.join(output_dir, f"{base_name}_段落{i+1}_{orientation_name}.docx")
                
                # 创建段落文档
                success = self._create_orientation_segment_copy(
                    source_content_path=source_file,
                    output_path=output_path,
                    start_section=start_section,
                    end_section=end_section
                )
                
                if success and os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                    split_files.append({
                        "file_path": output_path,
                        "orientation": orientation_name,
                        "index": i,
                        "sections": sections_desc
                    })
                    print(f"[SUCCESS] 已保存段落文件: {os.path.basename(output_path)} ({sections_desc}, {orientation_name})")
                elif os.path.exists(output_path):
                    # 删除空文件
                    try:
                        os.remove(output_path)
                    except Exception:
                        pass
            
            # 打印拆分结果
            print(f"\n[INFO] 拆分结果:")
            for item in split_files:
                print(f"   {item['index'] + 1}. {os.path.basename(item['file_path'])} ({item['sections']}, {item['orientation']})")
                
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

    ## 删除连续的两个分节符，保留前面一个
    def remove_empty_sections(self, source_file: str, output_file: str) -> bool:
        """
        删除文档中连续的两个分节符之间没有任何内容的后面那个分节符
        
        Args:
            source_file: 源文档路径
            output_file: 输出文档路径
            
        Returns:
            bool: 处理是否成功
        """
        print("-" * 50)
        print("[INFO] 删除空节")
        
        # 检查输入文件是否存在
        if not os.path.exists(source_file):
            print(f"  ❌[ERROR] 源文档不存在: {source_file}")
            return False
            
        # 确保输出目录存在
        output_dir = os.path.dirname(output_file)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
            
        # 检查COM组件是否可用
        if not COM_AVAILABLE or not win32 or not pythoncom:
            print("[ERROR] COM组件不可用，无法处理空节")
            return False
            
        pythoncom_initialized = False
        word_app = None
        doc = None
        
        try:
            # 初始化COM资源
            pythoncom_initialized, word_app = self._initialize_com_resources()
            if not pythoncom_initialized or word_app is None:
                return False
            
            # 打开文档
            doc = word_app.Documents.Open(source_file, ReadOnly=False)
            
            # 获取总节数
            total_sections = doc.Sections.Count if doc else 0
            print(f"📊 文档总节数: {total_sections}")
            
            if total_sections <= 1:
                print("[INFO] 文档节少于2个，无需处理")
                # 直接复制文件
                doc.SaveAs(output_file)
                return True
            
            # 从后往前遍历节，避免删除节后索引变化的问题
            sections_to_remove = []
            
            # 检查每两个连续的节之间是否为空
            for i in range(total_sections - 1, 0, -1):  # 从倒数第二节开始检查
                current_section = doc.Sections(i)
                next_section = doc.Sections(i + 1)
                
                # 检查当前节的内容是否为空
                # 获取当前节的范围（不包括分节符）
                current_range = current_section.Range
                # 获取下一节的范围（不包括分节符）
                next_range = next_section.Range
                
                # 检查下一节是否为空（没有文字、图片等内容）
                next_content = next_range.Text.strip()
                
                # 检查是否有图片或其他对象
                has_content = len(next_content) > 0 or next_range.InlineShapes.Count > 0 or next_range.ShapeRange.Count > 0
                
                if not has_content:
                    print(f"[INFO] 发现空节: 第{i+1}节为空，标记为待删除")
                    sections_to_remove.append(i + 1)  # 记录要删除的节索引
            
            # 删除标记的空节（从后往前删除避免索引变化）
            removed_count = 0
            for section_index in sorted(sections_to_remove, reverse=True):
                if section_index <= doc.Sections.Count:
                    try:
                        # 删除指定节的范围
                        section_to_remove = doc.Sections(section_index)
                        section_to_remove.Range.Delete()
                        removed_count += 1
                        print(f"[SUCCESS] 已删除第{section_index}节")
                    except Exception as e:
                        print(f"[WARNING] 删除第{section_index}节时出错: {str(e)}")
            
            print(f"📊 共删除了 {removed_count} 个空节")
            
            # 保存处理后的文档
            doc.SaveAs(output_file)
            print(f"💾 文档已保存到: {output_file}")
            return True
            
        except Exception as e:
            print(f"  ❌[ERROR] 处理空节时发生错误: {str(e)}")
            return False
            
        finally:
            # 释放COM资源
            self._release_com_resources(word_app, doc, pythoncom_initialized)

def main():
    """主函数，用于测试功能"""
    # 源文档路径
    source_file = r"C:\Users\yanha\Desktop\数字总师\文档\可行性报告（test）.docx"
    
    # 检查源文件是否存在
    if not os.path.exists(source_file):
        print(f"  ❌[ERROR] 源文件不存在: {source_file}")
        return
    
    # 创建处理器实例
    processor = DocumentOrientationProcessor()
    
    # 按页面方向段落拆分文档
    print("\n[INFO] 按页面方向段落拆分文档...")
    output_dir = os.path.join(current_dir, "orientation_segment_split_results")
    split_files = processor.split_document_by_orientation_segments(source_file, output_dir)
    
    print(f"\n[INFO] 拆分结果:")
    for item in split_files:
        print(f"   {item['index'] + 1}. {os.path.basename(item['file_path'])} ({item['sections']}, {item['orientation']})")
    
    print(f"\n[SUCCESS] 拆分完成，结果保存在: {output_dir}")


if __name__ == "__main__":
    main()