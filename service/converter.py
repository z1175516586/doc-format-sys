"""
完整文档格式化转换器

实现从源文档到格式化文档的完整转换流程:
1. 文档拆分 - 分离封面和正文
2. Pandoc转换 - 使用模板格式化正文(如果没有指定，默认使用template目录下的reference.docx)
3. 表格格式化 - 使用mcp服务格式化正文的表格（还没做到这一步，现在不管）
4. 文档合并 - 重新合并为完整文档
5. 标题修改 - 修改合并后的文档的目录标题
6. 图片格式化 - 图片居中，单倍行距
7. 补充处理 - 将库号右靠齐
8. 在目录之后插入分节符
9. 处理文档节的页码设置
10. 删除文件中中所有的突出显示

# 页眉替换 - 设置指定页眉内容 // 直接替换模板文件的，现在不管

"""

import os
import sys
import tempfile
import shutil
import time
import zipfile
import subprocess
import json
from datetime import datetime
from typing import List, Optional, Tuple, Dict, Any
from pathlib import Path
from docx import Document

# 导入所需的工具模块
from utils.pandoc_converter import PandocConverter
from utils.docx_split import DocxSplitProcessor
from utils.docx_merge import copy_all_to_beginning
from utils.docx_update_toc_title import update_toc_title_xml
from utils.convert_word_equations import replace_ole_with_images
from utils.check_history_logo import print_dochistoryid
from docx.oxml.ns import qn
from docx.shared import Pt, Cm

# 导入各步骤中使用的模块
from utils.document_orientation_processor import DocumentOrientationProcessor
from utils.docx_header_footer_replace import replace_header_footer_content, extract_header_footer_content
from utils.cover_replace import replace_content_in_cover_auto, replace_content_in_cover_by_dify, apply_character_spacing_to_cover, force_cover_title_spacing_in_docx
from utils.cover_table_replace import replace_table_after_marker
from utils.docx_table_excel import convert_embedded_excels_inplace
from utils.docx_table_replace import replace_tables_by_index
from utils.docx_update_toc_title import update_toc_title
from utils.docx_picture import format_pictures_with_advanced_settings
from utils.docx_supplement import format_library_number_advanced
from utils.docx_supplement import insert_section_break_after_toc
from utils.docx_supplement import cancel_section_link_com
from utils.docx_supplement import process_section2_docx
from utils.docx_supplement import process_section3_docx
from utils.cover_number_format import process_document
from service.dify_service import DifyService

# COM 组件可用性检测（用于页面方向拆分等功能）
try:
    import pythoncom
    import win32com.client as win32
    from win32com.client import constants as C

    COM_AVAILABLE = True
except Exception:
    pythoncom = None
    win32 = None
    C = None
    COM_AVAILABLE = False


WD_ALIGN_PARAGRAPH_CENTER = getattr(C, "wdAlignParagraphCenter", 1) if C else 1

# 添加项目路径
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
sys.path.insert(0, parent_dir)

# 添加lxml导入用于去除突出显示
try:
    from lxml import etree

    LXML_AVAILABLE = True
except ImportError:
    LXML_AVAILABLE = False
    etree = None


def _kill_existing_winword_processes():
    if os.name != "nt":
        return
    cmd = ["taskkill", "/f", "/im", "WINWORD.EXE"]
    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    try:
        subprocess.run(
            cmd,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
            creationflags=creationflags,
        )
    except Exception:
        pass


def _wait_for_file_release_for_step(path: str, timeout: float = 10.0, interval: float = 0.3) -> bool:
    """
    Wait until the file can be opened for read/write. Returns True on success.
    """
    if not path:
        return True
    target = Path(path)
    deadline = time.time() + timeout
    while True:
        try:
            if target.exists():
                with open(target, "a+b"):
                    return True
            else:
                return True
        except OSError:
            if time.time() >= deadline:
                return False
            time.sleep(interval)


class DocumentConverter:
    """文档格式化转换器"""

    def __init__(self, document_type: int = 1):
        """初始化转换器"""
        self.temp_dir = None
        self.pandoc_converter = None
        self.intermediate_files = {}  # 保存中间文件路径
        self.debug_output_dir = os.path.join(parent_dir, 'temp')  # 中间文件保存目录
        self.save_intermediate_files = False  # 是否保存中间文件的开关
        self.document_type = document_type  # 文档类型参数
        self.extracted_footer_content = None  # 存储提取的页脚内容
        self.step0_processed_source = None  # 记录步骤0/1.8的文件
        self.step1_8_image_dir = None  # 步骤1.8图片目录
        self.step1_8_output = None  # 步骤1.8输出文档

    def __enter__(self):
        """上下文管理器入口"""
        _kill_existing_winword_processes()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """上下文管理器出口，清理临时文件"""
        # self.cleanup()

    def cleanup(self):
        """清理临时文件"""
        if self.temp_dir and os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir, ignore_errors=True)

    @staticmethod
    def _normalize_toc_title(text: str) -> str:
        return "".join(text.split()).lower()

    def _center_toc_title_with_com(self, docx_file: str, toc_title: str) -> bool:
        """使用 COM 将目录标题段落设置为居中。"""
        if not COM_AVAILABLE or not win32 or not pythoncom:
            print("⚠️ COM 不可用，跳过目录标题居中设置")
            return False
        if not os.path.exists(docx_file):
            print(f"⚠️ 文档不存在，无法设置目录标题居中: {docx_file}")
            return False

        candidates = {
            toc_title,
            "目录",
            "目 录",
            "目  录",
            "Table of Contents",
            "Contents",
            "TABLE OF CONTENTS",
        }
        normalized_candidates = {
            self._normalize_toc_title(item) for item in candidates if item and item.strip()
        }

        word = None
        doc = None
        com_initialized = False
        try:
            pythoncom.CoInitialize()
            com_initialized = True
            word = win32.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False
            doc = word.Documents.Open(docx_file, ReadOnly=False)

            max_check = min(doc.Paragraphs.Count, 300)
            for i in range(1, max_check + 1):
                para = doc.Paragraphs(i)
                text = para.Range.Text.strip()
                if not text:
                    continue
                normalized_text = self._normalize_toc_title(text)
                if normalized_text in normalized_candidates:
                    para.Range.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_CENTER
                    para.Range.Font.Bold = True
                    para.Range.Font.NameFarEast = "宋体"
                    para.Range.Font.Name = "宋体"
                    para.Range.Font.Size = 14
                    doc.Save()
                    print("✅ 目录标题已设置为居中")
                    return True

            print("ℹ️ 未定位到目录标题段落，跳过居中设置")
            return False
        except Exception as exc:
            print(f"⚠️ 设置目录标题居中失败: {exc}")
            return False
        finally:
            if doc is not None:
                try:
                    doc.Close(SaveChanges=0)
                except Exception:
                    pass
            if word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    def _append_document_with_formatting(
        self,
        base_file: str,
        file_to_append: str,
        output_file: str,
    ) -> bool:
        """将 file_to_append 追加到 base_file 末尾（保留页面方向）。"""
        try:
            from utils.merge_docs_preserve_orientation import (
                append_file_with_formatting as merge_append,
            )
        except ImportError as err:
            print(f"[warn] 无法导入 merge_docs_preserve_orientation: {err}")
            return False

        try:
            merge_append(base_file, file_to_append, output_file)
            return True
        except Exception as exc:
            print(f"[warn] 横版文档追加失败: {exc}")
            import traceback

            traceback.print_exc()
            return False


    def validate_input_files(self, source_file: str, template_file: str) -> bool:
        """
        验证输入文件的有效性

        Args:
            source_file: 源文档路径
            template_file: 模板文档路径

        Returns:
            bool: 文件是否有效
        """
        if os.name == 'nt':
            _kill_existing_winword_processes()

        if not os.path.exists(source_file):
            print(f"❌ 源文档不存在: {source_file}")
            return False

        if not os.path.exists(template_file):
            print(f"❌ 模板文档不存在: {template_file}")
            return False

        # 检查文件是否为DOCX格式
        if not source_file.lower().endswith('.docx'):
            print(f"❌ 源文档不是DOCX格式: {source_file}")
            return False

        if not template_file.lower().endswith('.docx'):
            print(f"❌ 模板文档不是DOCX格式: {template_file}")
            return False

        return True

    def _save_intermediate_file(self, source_path: str, step_name: str, file_description: str = "") -> None:
        """
        保存中间文件到指定目录便于查看和调试

        Args:
            source_path: 源文件路径
            step_name: 步骤名称（如 step1_split, step2_pandoc 等）
            file_description: 文件描述（如封面, 正文内容 等）
        """
        # 如果不保存中间文件，则直接返回
        if not self.save_intermediate_files:
            return

        try:
            # 确保输出目录存在
            if not os.path.exists(self.debug_output_dir):
                os.makedirs(self.debug_output_dir, exist_ok=True)
                print(f"📁 创建调试输出目录: {self.debug_output_dir}")

            # 生成带时间戳的文件名
            from datetime import datetime
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

            base_name = os.path.splitext(os.path.basename(source_path))[0]
            if file_description:
                debug_filename = f"{step_name}_{file_description}_{timestamp}.docx"
            else:
                debug_filename = f"{step_name}_{base_name}_{timestamp}.docx"

            debug_path = os.path.join(self.debug_output_dir, debug_filename)

            # 复制文件
            shutil.copy2(source_path, debug_path)

            print(f"   💾 已保存调试文件: {debug_filename}")

        except Exception as e:
            print(f"   ⚠️ 保存调试文件失败: {str(e)}")

    # def step0_preprocess_headings_new(self, source_file: str) -> str:
    #     """
    #     步骤-1: 预处理标题样式（严格编号 + 标题文本规则 + 顺序比较 + 反向纠错）
    #     """
    #     import os
    #     import shutil
    #     import re
    #     from docx import Document
    #     from docx.shared import Pt

    #     print("-" * 50)
    #     print("📝 步骤0: 预处理标题样式（严格编号规则 + 顺序比较 + 反向纠错）")

    #     try:
    #         if not self.temp_dir:
    #             raise ValueError("临时目录未初始化")

    #         base_name = os.path.splitext(os.path.basename(source_file))[0]
    #         preprocessed_path = os.path.join(self.temp_dir, f"{base_name}_预处理标题.docx")

    #         shutil.copy2(source_file, preprocessed_path)
    #         doc = Document(preprocessed_path)

    #         # 严格编号 + 标题文本规则正则
    #         heading_pattern = re.compile(
    #             r'^(\d+(?:\.\d+){0,3})[\s\t]*(?![)）])(.+)$'
    #         )

    #         # ---------------- 内部函数：标题序号比较（增强版 + 跳号约束） ----------------
    #         def is_valid_heading_order(prev_num: str, curr_num: str) -> bool:
    #             # print(f"   比较: {prev_num} → {curr_num}")
    #             A = [int(x) for x in prev_num.split('.')]
    #             B = [int(x) for x in curr_num.split('.')]

    #             max_len = max(len(A), len(B))
    #             A += [0] * (max_len - len(A))
    #             B += [0] * (max_len - len(B))

    #             for i in range(max_len):
    #                 if A[i] < B[i]:
    #                     diff = B[i] - A[i]
    #                     # 只允许 +0 或 +1
    #                     if diff in (0, 1):
    #                         return True
    #                     return False

    #                 if A[i] > B[i]:
    #                     # 特例：1.2.4 → 1.3
    #                     if i > 0 and A[i - 1] + 1 == B[i - 1]:
    #                         # 要求 B[i:] 全为 0
    #                         if sum(B[i:]) == 0:
    #                             return True
    #                     return False

    #             return True

    #         # ---------------- 正向识别标题 ----------------
    #         prev_heading_number = None

    #         for para in doc.paragraphs:
    #             text = para.text.strip()
    #             if not text:
    #                 continue

    #             match = heading_pattern.match(text)
    #             if not match:
    #                 continue

    #             curr_number = match.group(1)
    #             level = curr_number.count('.') + 1

    #             # ---------- 顺序判断 ----------
    #             if prev_heading_number is None:
    #                 is_valid = True
    #             else:
    #                 is_valid = is_valid_heading_order(prev_heading_number, curr_number)

                
    #             if is_valid:
    #                 print(
    #                 f"   🔍 发现标题: {curr_number} {match.group(2)} (级别: {level})"
    #                 + (" ✅")
    #             )
    #             else:
    #                 print(f"   ❌ 不是标题 {curr_number} {match.group(2)} 原因："
    #                         + "不符合标题格式 或 前一个段落末尾为软回车" 
    #                         + f" 或 与前一个标题 {prev_heading_number} 顺序不匹配。"
    #                         + "请手动修改！")

    #             if not is_valid:
    #                 continue

    #             # ---------- 设置标题样式 ----------
    #             if level == 1:
    #                 para.style = 'Heading 1'
    #             elif level == 2:
    #                 para.style = 'Heading 2'
    #             elif level == 3:
    #                 para.style = 'Heading 3'
    #             else:
    #                 para.style = 'Heading 4'

    #             para.paragraph_format.space_before = Pt(0)
    #             para.paragraph_format.space_after = Pt(0)

    #             prev_heading_number = curr_number

    #         # ---------------- 反向纠错 ----------------
    #         heading_styles = ['Heading 1', 'Heading 2', 'Heading 3', 'Heading 4']
    #         for para in doc.paragraphs:
    #             if para.style.name in heading_styles:
    #                 text = para.text.strip()
    #                 if not heading_pattern.match(text):
    #                     para.style = 'Normal'
    #                     para.paragraph_format.space_before = Pt(0)
    #                     para.paragraph_format.space_after = Pt(0)
    #                     if text:
    #                         print(f"⚠️ 原文档格式可疑处: '{text[:30]}...'")

    #         doc.save(preprocessed_path)
    #         print("✅ 标题预处理完成!")

    #         if self.save_intermediate_files:
    #             self._save_intermediate_file(
    #                 preprocessed_path, "step0_preprocess", "标题预处理"
    #             )

    #         return preprocessed_path

    #     except Exception as e:
    #         print(f"❌ 标题预处理过程中发生错误: {str(e)}")
    #         return source_file

    def step0_preprocess_headings_new(self, source_file: str) -> str:
        """
        步骤0: 预处理标题样式（严格编号规则 + 顺序比较 + 反向纠错）

        处理规则：
        - 识别“目录”
        - 目录后遇到的第一个“1”级标题忽略
        - 从第二个“1”级标题开始，正式处理正文标题（第二个 1 本身必须被处理）
        - 标题编号必须以数字结尾，"." 后必须接数字，禁止 "3 . xxx" 误判
        """
        import os
        import shutil
        import re
        from docx import Document
        from docx.shared import Pt

        print("-" * 50)
        print("📝 步骤0: 预处理标题样式（正文范围内）")

        try:
            if not self.temp_dir:
                raise ValueError("临时目录未初始化")

            base_name = os.path.splitext(os.path.basename(source_file))[0]
            preprocessed_path = os.path.join(
                self.temp_dir, f"{base_name}_预处理标题.docx"
            )

            shutil.copy2(source_file, preprocessed_path)
            doc = Document(preprocessed_path)

            # =========================================================
            # ① 标题正则
            # =========================================================
            heading_pattern = re.compile(
                r'^(\d+(?:\.\d+){0,4})[\s\t]*(?![)）\.、-])(.+)$'
            )

            # =========================================================
            # ② 标题序号顺序判断
            # =========================================================
            def is_valid_heading_order(prev_num: str, curr_num: str) -> bool:
                A = [int(x) for x in prev_num.split('.')]
                B = [int(x) for x in curr_num.split('.')]

                max_len = max(len(A), len(B))
                A += [0] * (max_len - len(A))
                B += [0] * (max_len - len(B))

                for i in range(max_len):
                    if A[i] < B[i]:
                        diff = B[i] - A[i]
                        return diff in (0, 1)
                    if A[i] > B[i]:
                        if i > 0 and A[i - 1] + 1 == B[i - 1]:
                            if sum(B[i:]) == 0:
                                return True
                        return False
                return True

            # =========================================================
            # ③ 扫描目录
            # =========================================================
            toc_regex = re.compile(r'目[\s\t]*录')
            after_toc = False
            level1_count = 0
            processing_enabled = False
            prev_heading_number = None

            print("🔎 正在扫描“目录”关键字…")

            for para in doc.paragraphs:
                text = para.text.strip()
                if not text:
                    continue

                # ------------------ 扫描目录 ------------------
                if not after_toc:
                    if toc_regex.search(text):
                        after_toc = True
                        print(f"✅ 识别到目录段落: '{text}'")
                    # 目录前完全不处理标题
                    continue

                # ------------------ 目录之后，正文处理 ------------------
                # ===== 新增：排除含 pict 的段落 =====
                p_xml = para._element.xml
                if "<w:pict" in p_xml:
                    print(f"   🚫 含 pict 标签，忽略标题: '{text}'")
                    continue

                match = heading_pattern.match(text)
                if not match:
                    continue

                # -------- 新增括号平衡校验 --------
                num_right = text.count(')') + text.count('）')
                num_left = text.count('(') + text.count('（')
                if num_right - num_left >= 1:
                    # 不认为是标题，直接跳过
                    print(f"   ⚠️ 段落括号不平衡，忽略标题: '{text}'")
                    continue

                curr_number = match.group(1)
                title_text = match.group(2)
                level = curr_number.count('.') + 1

                # 第1个1级标题忽略，第2个开始处理
                if not processing_enabled:
                    if curr_number == "1" and level == 1:
                        level1_count += 1
                        if level1_count == 1:
                            print(f"⏭️ 遇到第 1 个 1 级标题，暂不处理: '{text}'")
                            continue
                        elif level1_count == 2:
                            processing_enabled = True
                            prev_heading_number = None
                            print(f"📌 已定位正文起点（第二个 1 级标题）: '{text}'")
                            # ⚠️ 不 continue，让第二个1级标题自己参与处理
                    else:
                        # 第1个1级标题和第2个1级标题之间的段落不处理
                        continue

                # ---------- 顺序判断 ----------
                if prev_heading_number is None:
                    is_valid = True
                else:
                    is_valid = is_valid_heading_order(prev_heading_number, curr_number)

                if not is_valid:
                    print(
                        f"   ❌ 不是标题 {curr_number} {title_text}，"
                        f"与前一个标题 {prev_heading_number} 顺序不匹配"
                    )
                    # ✅ 将顺序不合法的段落重置为正文样式
                    para.style = 'Normal'
                    para.paragraph_format.space_before = Pt(0)
                    para.paragraph_format.space_after = Pt(0)
                    continue

                print(
                    f"   🔍 发现标题: {curr_number} {title_text} "
                    f"(级别: {level}) ✅"
                )

                # ---------- 设置标题样式 ----------
                if level == 1:
                    para.style = 'Heading 1'
                elif level == 2:
                    para.style = 'Heading 2'
                elif level == 3:
                    para.style = 'Heading 3'
                elif level == 4:
                    para.style = 'Heading 4'
                else:
                    para.style = 'Heading 5'

                para.paragraph_format.space_before = Pt(0)
                para.paragraph_format.space_after = Pt(0)

                prev_heading_number = curr_number

            # =========================================================
            # ④ 反向纠错
            # =========================================================
            heading_styles = ['Heading 1', 'Heading 2', 'Heading 3', 'Heading 4', 'Heading 5']

            for para in doc.paragraphs:
                if para.style.name in heading_styles:
                    text = para.text.strip()
                    if not heading_pattern.match(text):
                        para.style = 'Normal'
                        para.paragraph_format.space_before = Pt(0)
                        para.paragraph_format.space_after = Pt(0)
                        if text:
                            print(f"⚠️ 原文档格式可疑处: '{text[:30]}...'")

            doc.save(preprocessed_path)
            print("✅ 标题预处理完成!")

            if self.save_intermediate_files:
                self._save_intermediate_file(
                    preprocessed_path, "step0_preprocess", "标题预处理"
                )

            return preprocessed_path

        except Exception as e:
            print(f"❌ 标题预处理过程中发生错误: {str(e)}")
            return source_file

    def step0_preprocess_headings(self, source_file: str) -> str:
        """
        步骤-1: 预处理标题样式 - 使用Dify服务获取标题信息并设置标题级别

        Args:
            source_file: 源文档路径

        Returns:
            str: 预处理后的文档路径
        """
        print("-" * 50)
        print("📝 步骤0: 预处理标题样式")

        try:
            # 确保临时目录存在
            if not self.temp_dir:
                raise ValueError("临时目录未初始化")

            # 生成预处理后的文件路径
            base_name = os.path.splitext(os.path.basename(source_file))[0]
            preprocessed_path = os.path.join(self.temp_dir, f"{base_name}_预处理标题.docx")

            # 复制源文件到临时文件
            shutil.copy2(source_file, preprocessed_path)

            # 使用Dify服务获取标题信息
            try:
                # 初始化Dify服务，使用环境变量或默认值
                dify_service = DifyService(api_key="app-rvIPHVXm4X4AjxtE2K5cwUhs", base_url="http://172.23.109.63:8088/v1")
                # api_key = os.getenv("DIFY_API_KEY", "app-COuiz5apqunoO2B1f0oaUNZ2")
                # dify_service = DifyService(api_key=api_key)

                # 提取标题信息
                titles_json = dify_service.extract_outputs_content(source_file)
                # print(titles_json)

                # 检查返回的内容是否为错误信息
                if titles_json.startswith("JSON解析错误:") or titles_json.startswith("处理文档时发生错误:") or titles_json.startswith("处理过程中发生错误:"):
                    raise Exception(titles_json)

                # 解析标题信息
                titles = json.loads(titles_json)

                print(f"✅ 成功获取到 {len(titles)} 个标题")

                # 创建标题文本到级别的映射
                title_mapping = {}
                for item in titles:
                    title_text = item.get("text", "").strip()
                    title_level = item.get("level", 1)
                    title_mapping[title_text] = title_level

                    # 同时创建不带换行符的版本（处理可能的格式差异）
                    title_text_no_newlines = title_text.replace('\n', ' ').strip()
                    title_mapping[title_text_no_newlines] = title_level

                print("📋 标题映射表:")
                for title, level in list(title_mapping.items())[:5]:  # 只显示前5个
                    print(f"   '{title}' -> 级别 {level}")
                if len(title_mapping) > 5:
                    print(f"   ... 还有 {len(title_mapping) - 5} 个标题")

            except Exception as dify_error:
                print(f"⚠️ 使用Dify服务获取标题信息失败: {str(dify_error)}")
                print("   将使用原有的基于编号规则的标题处理方式")
                titles = []
                title_mapping = {}

            # 加载文档并处理标题
            # 导入Document类（忽略静态分析工具的警告）
            try:
                doc = Document(preprocessed_path)
            except ImportError:
                print("❌ 无法导入python-docx库，请确保已安装：pip install python-docx")
                return source_file

            # 如果成功获取了标题信息，则使用Dify服务的结果
            if title_mapping:
                print("🔧 使用Dify服务获取的标题信息设置标题级别")

                for para in doc.paragraphs:
                    # 获取段落文本
                    text = para.text.strip()

                    # 查找匹配的标题
                    matched_level = None
                    matched_title = None

                    # 精确匹配
                    if text in title_mapping:
                        matched_level = title_mapping[text]
                        matched_title = text
                    else:
                        # 模糊匹配：检查是否是某个标题的一部分
                        for title, level in title_mapping.items():
                            # 处理可能包含换行符的情况
                            if text.replace('\n', ' ').strip() == title.replace('\n', ' ').strip():
                                matched_level = level
                                matched_title = title
                                break

                    # 如果找到匹配的标题，则设置相应的样式
                    if matched_level is not None:
                        print(f"   匹配标题: '{matched_title}' -> 级别 {matched_level}")

                        # 根据层级应用相应的样式
                        if matched_level == 1:
                            para.style = 'Heading 1'  # 一级标题
                        elif matched_level == 2:
                            para.style = 'Heading 2'  # 二级标题
                        elif matched_level == 3:
                            para.style = 'Heading 3'  # 三级标题
                        elif matched_level == 4:
                            para.style = 'Heading 4'  # 四级标题
                        else:
                            # 对于其他级别，默认使用Heading 4
                            para.style = 'Heading 4'
                        
                        # 设置段前段后间距为0
                        para.paragraph_format.space_before = Pt(0)
                        para.paragraph_format.space_after = Pt(0)
            else:
                # 如果没有成功获取标题信息，则使用原有的基于编号规则的处理方式
                print("🔧 使用原有的基于编号规则的标题处理方式")

                for para in doc.paragraphs:
                    # 获取段落文本
                    text = para.text.strip()

                    # 判断标题的层级（通过编号规则进行判断）
                    if text and text[0].isdigit() and '.' in text:  # 以数字开始且包含.
                        if '，' not in text and '。' not in text and '；' not in text:  # 不包含，或者。
                            # 按点分割序号，判断层级
                            level = text.count('.') + 1

                            # 根据层级应用相应的样式
                            if level == 1:
                                para.style = 'Heading 1'  # 一级标题
                            elif level == 2:
                                para.style = 'Heading 2'  # 二级标题
                            elif level == 3:
                                para.style = 'Heading 3'  # 三级标题
                            elif level == 4:
                                para.style = 'Heading 4'  # 四级标题
                            
                            # 设置段前段后间距为0
                            para.paragraph_format.space_before = Pt(0)
                            para.paragraph_format.space_after = Pt(0)

            # 保存修改后的文档
            doc.save(preprocessed_path)
            print("✅ 标题预处理完成!")

            # 保存中间文件
            if self.save_intermediate_files:
                self._save_intermediate_file(preprocessed_path, "step0_preprocess", "标题预处理")

            return preprocessed_path

        except Exception as e:
            print(f"❌ 标题预处理过程中发生错误: {str(e)}")
            return source_file

    def step0_replace_header_footer(self, source_file: str, template_file: str, proj_name: str, file_type_name: str) -> str:
        """
        步骤0: 页眉页脚替换 - 将源文档的页眉页脚内容替换到模板文档中

        Args:
            source_file: 源文档路径（提供页眉页脚内容）
            template_file: 模板文档路径（被替换页眉页脚内容）

        Returns:
            str: 替换页眉页脚后的模板文件路径
        """
        print("-" * 50)
        print("📑 步骤0: 页眉页脚替换")

        try:
            # 确保临时目录存在
            if not self.temp_dir:
                raise ValueError("临时目录未初始化")

            # 生成输出文件路径
            base_name = os.path.splitext(os.path.basename(template_file))[0]
            updated_template_path = os.path.join(self.temp_dir, f"{base_name}_页眉页脚替换后.docx")

            print(f"📄 源文档: {os.path.basename(source_file)}")
            print(f"📄 模板文档: {os.path.basename(template_file)}")
            print(f"📤 更新后模板: {os.path.basename(updated_template_path)}")

            # 使用docx_header_footer_replace.py中的方法进行页眉页脚替换
            try:
                from utils.docx_header_footer_replace import replace_header_footer_content, \
                    extract_header_footer_content
                # 提取源文档的页眉页脚内容
                source_content = extract_header_footer_content(source_file, 2)  # 从源文档第2节提取

                # 保存提取的页脚内容，用于后续步骤
                if "footers" in source_content and source_content["footers"]:
                    # 获取默认类型的页脚内容
                    default_footer = source_content["footers"].get("default", "")
                    even_footer = source_content["footers"].get("even", "")
                    first_footer = source_content["footers"].get("first", "")

                    # 优先使用默认页脚，如果没有则尝试其他类型
                    self.extracted_footer_content = default_footer or even_footer or first_footer
                    print(f"📄 提取的页脚内容: {self.extracted_footer_content}")
                else:
                    print("⚠️ 未提取到页脚内容，将使用默认页脚格式")
                    self.extracted_footer_content = None
                # 执行页眉页脚替换
                success = replace_header_footer_content(
                    source_docx_path=source_file,
                    target_docx_path=template_file,
                    source_section_index=2,  # 从源文档第1节提取
                    target_section_index=1,  # 替换到模板文档第1节
                    save_path=updated_template_path,
                    proj_name=proj_name,
                    file_type_name=file_type_name
                )

                if success and os.path.exists(updated_template_path):
                    print("✅ 页眉页脚替换成功!")

                    # 保存中间文件到指定目录便于查看调试
                    if self.save_intermediate_files:
                        print(f"   更新后模板: {os.path.basename(updated_template_path)}")
                        print(f"📁 正在保存step0中间文件到: {self.debug_output_dir}")
                        self._save_intermediate_file(updated_template_path, "step0_header_footer", "替换后模板")

                    return updated_template_path
                else:
                    print("❌ 页眉页脚替换失败，使用原始模板文件")
                    return template_file

            except Exception as replace_error:
                print(f"❌ 页眉页脚替换过程中发生错误: {str(replace_error)}")
                print("   继续使用原始模板文件")
                return template_file

        except Exception as e:
            print(f"❌ 页眉页脚替换过程中发生错误: {str(e)}")
            return template_file

    from typing import List, Dict, Any, Tuple
    def step1_8_convert_equations(self, source_file: str, word=None) ->  Tuple[str, List[Dict[str, Any]]]:
        """
        Step 1.8: 将竖版正文中的 OLE 公式转换为图片，供后续步骤使用。
        """
        print("-" * 50)
        print("步骤1.8: 公式图片化处理")

        formula_info_list = []

        try:
            if not self.temp_dir:
                raise ValueError("临时目录未初始化")

            base_name = os.path.splitext(os.path.basename(source_file))[0]
            output_path = os.path.join(self.temp_dir, f"{base_name}_step1_8_equation_images.docx")
            image_dir = Path(self.temp_dir) / f"{base_name}_formula_images"
            self.step1_8_image_dir = str(image_dir)

            try:
                images, formula_info_list = replace_ole_with_images(
                    docx_path=source_file,
                    output_docx_path=output_path,
                    output_img_dir=image_dir,
                    word_visible=False,
                    word=word
                )
            except Exception as e:
                if word is not None:
                    print(f"步骤1.8: 使用现有Word实例失败 ({e})，尝试重新创建实例...")
                    images, formula_info_list = replace_ole_with_images(
                        docx_path=source_file,
                        output_docx_path=output_path,
                        output_img_dir=image_dir,
                        word_visible=False,
                        word=None
                    )
                else:
                    raise e

            for idx, info in enumerate(formula_info_list):
                info["file"] = images[idx]
                info["uid"] = f"FORMULA_{idx+1}"

            if images:
                print(f"步骤1.8: 成功转换 {len(images)} 张公式图片")
            else:
                print("步骤1.8: 未检测到需要转换的公式，已生成副本文档继续后续流程。")

            if self.save_intermediate_files and os.path.exists(output_path):
                self._save_intermediate_file(output_path, "step1_8_equations", "equation_images")
                if image_dir.exists():
                    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                    archive_name = f"step1_8_equations_images_{timestamp}.zip"
                    archive_path = Path(self.debug_output_dir) / archive_name
                    try:
                        with zipfile.ZipFile(archive_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                            for img_file in image_dir.glob('**/*'):
                                if img_file.is_file():
                                    zf.write(img_file, arcname=img_file.name)
                        print(f"   已保存公式图片压缩包: {archive_name}")
                    except Exception as zip_err:
                        print(f"   压缩公式图片失败: {zip_err}")

            if os.path.exists(output_path):
                self.step1_8_output = output_path
                self.intermediate_files['step1_8_output'] = output_path
            if images:
                self.intermediate_files['step1_8_images'] = str(image_dir)

            return output_path, formula_info_list

        except Exception as err:
            print(f"步骤1.8: 图片转换发生错误，继续使用原文件。详细信息：{err}")
            return source_file, formula_info_list

    def step1_9_extract_and_map(self, source_file: str, word=None) -> str:
        """
        Step 1.9: 运行 extract_and_map（提取 Shapes 为图片 + 生成图注映射 + 输出清理版文档）。

        - 输入：source_file（按需求为 step1.8 的输入文件）
        - 输出：图片/映射/清理版 docx 均落在 self.temp_dir
        """
        print("-" * 50)
        print("步骤1.9: 绘图提取与映射（extract_and_map）")

        if os.name != "nt":
            print("⚠️ 当前系统非Windows，跳过 extract_and_map")
            return source_file

        try:
            from utils.extract_and_map import extract_and_generate_map
        except Exception as import_err:
            print(f"⚠️ extract_and_map 不可用，跳过。原因: {import_err}")
            return source_file

        try:
            if not self.temp_dir:
                raise ValueError("临时目录未初始化")

            base_name = os.path.splitext(os.path.basename(source_file))[0]
            os.makedirs(self.debug_output_dir, exist_ok=True)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            run_dir = os.path.join(self.debug_output_dir, f"step1_9_extract_and_map_{timestamp}_{base_name}")
            os.makedirs(run_dir, exist_ok=True)
            output_doc = os.path.join(run_dir, f"{base_name}_clean.docx")

            try:
                clean_doc_path, map_json_path, image_map = extract_and_generate_map(
                    src_path=source_file,
                    output_dir=run_dir,
                    output_doc_path=output_doc,
                    word_visible=True,
                    word=word
                )
            except Exception as e:
                if word is not None:
                    print(f"步骤1.9: 使用现有Word实例失败 ({e})，尝试重新创建实例...")
                    clean_doc_path, map_json_path, image_map = extract_and_generate_map(
                        src_path=source_file,
                        output_dir=run_dir,
                        output_doc_path=output_doc,
                        word_visible=True,
                        word=None
                    )
                else:
                    raise e

            if clean_doc_path and os.path.exists(clean_doc_path):
                self.intermediate_files["step1_9_extract_clean_doc"] = clean_doc_path
            if map_json_path and os.path.exists(map_json_path):
                self.intermediate_files["step1_9_image_map_json"] = map_json_path
            self.intermediate_files["step1_9_run_dir"] = run_dir
            self.intermediate_files["step1_9_image_map_count"] = str(len(image_map or {}))

            if clean_doc_path and os.path.exists(clean_doc_path):
                print(f"✅ extract_and_map 完成: 输出目录={run_dir}")
                return clean_doc_path

            print("⚠️ extract_and_map 未生成清理版文档，继续使用原文件")
            return source_file

        except Exception as err:
            print(f"⚠️ extract_and_map 运行失败，继续使用原文件。详细信息：{err}")
            return source_file

    def _table_signature(self, table) -> tuple:
        signature = []

        try:
            for row in table.rows:
                row_text = []
                for cell in row.cells:
                    row_text.append(cell.text.strip())
                signature.append(tuple(row_text))

        except Exception as e:
            print("=" * 80)
            print("❌ 表格解析失败")
            print("错误原因:", e)

            # 输出当前表格 XML
            try:
                print("\n📌 异常表格 XML:")
                print(table._element.xml[:5000])  # 防止太长，只打印前5000字符
            except Exception as xml_err:
                print("XML输出失败:", xml_err)

            print("=" * 80)

            # 继续抛出，让外层捕获
            raise e

        return tuple(signature)
    
    def step1_6_operation_document(self, source_file: str, template_file: Optional[str] = None):
        """
        步骤1.6: 操作文档 - 按页面方向拆分文档，对竖版使用Pandoc转换，横版保持原样，
        然后按顺序合并文档并删除空节
        
        Args:
            source_file: 源文档路径
            template_file: 模板文件路径（可选）
            
        Returns:
            str: 处理后的文件路径，如果处理失败则返回None
        """
        print("-" * 50)
        print("🛠️ 步骤1.6: 操作文档")
        
        try:
            # 确保临时目录存在
            if not self.temp_dir:
                import tempfile
                self.temp_dir = tempfile.mkdtemp(prefix="doc_converter_")
                print(f"📁 临时目录: {self.temp_dir}")
            
            # 如果未指定模板文件，则使用默认模板
            if template_file is None:
                # 获取项目根目录
                current_dir = os.path.dirname(os.path.abspath(__file__))
                parent_dir = os.path.dirname(current_dir)
                template_file = os.path.join(parent_dir, 'template', 'reference_content.docx')
                print(f"信息: 未指定模板文件，使用默认模板: {template_file}")
            
            # 创建输出目录
            output_dir = os.path.join(self.temp_dir, "DocumentOrientationProcessor")
            os.makedirs(output_dir, exist_ok=True)
            
            # 创建 DocumentOrientationProcessor 实例
            processor = DocumentOrientationProcessor()
            
            # 拆分文档
            split_files = processor.split_document_by_orientation_segments(source_file, output_dir)
            
            # 处理拆分后的文件
            print("\n开始处理拆分后的文件...")
            
            # 确保有拆分的文件
            if not split_files:
                print("没有拆分的文件可以处理")
                return None
            
            # 按索引排序，确保按正确顺序处理
            split_files.sort(key=lambda x: x['index'])
            
            # 创建处理后的文件目录
            processed_output_dir = os.path.join(self.temp_dir, "processed")
            os.makedirs(processed_output_dir, exist_ok=True)
            
            # 创建 PandocConverter 实例
            pandoc_converter = PandocConverter()
            
            # 处理每个拆分的文件
            processed_files = []
            first_vertical_processed = False  # 标记是否已处理第一个竖版文件
            
            for i, file_info in enumerate(split_files):
                original_file_path = file_info['file_path']
                orientation = file_info['orientation']
                file_name = os.path.basename(original_file_path)
                
                print(f"\n  处理第 {i+1} 个文件: {file_name} ({orientation})")
                
                # 如果是竖版文件，使用 Pandoc 处理
                if orientation == "竖版":
                    # 生成处理后的文件路径
                    processed_file_path = os.path.join(processed_output_dir, f"processed_{file_name}")
                    
                    # 检查是否是第一个竖版文件
                    if not first_vertical_processed:
                        # 第一个竖版文件使用 --table-of-contents 参数
                        print(f"    → 这是第一个竖版文件，使用目录参数处理")
                        success = pandoc_converter.convert_with_template(
                            input_file=original_file_path,
                            output_file=processed_file_path,
                            template_file=template_file,
                            additional_args=[
                                "--table-of-contents"
                            ]
                        )
                        first_vertical_processed = True
                    else:
                        # 后续竖版文件不使用目录参数
                        print(f"    → 这是后续竖版文件，不使用目录参数处理")
                        success = pandoc_converter.convert_with_template(
                            input_file=original_file_path,
                            output_file=processed_file_path,
                            template_file=template_file
                        )
                    
                    if success and os.path.exists(processed_file_path):
                        print(f"    ✓ 竖版文件处理成功: {os.path.basename(processed_file_path)}")
                        processed_files.append({
                            'file_path': processed_file_path,
                            'orientation': orientation,
                            'index': file_info['index'],
                            'sections': file_info['sections']
                        })
                    else:
                        print(f"    ✗ 竖版文件处理失败，使用原始文件: {file_name}")
                        processed_files.append(file_info)
                else:
                    # 横版文件保持原样
                    print(f"    → 横版文件保持原样: {file_name}")
                    processed_file_path = os.path.join(processed_output_dir, f"processed_{file_name}")
                    shutil.copy2(original_file_path, processed_file_path)
                    processed_files.append({
                        'file_path': processed_file_path,
                        'orientation': orientation,
                        'index': file_info['index'],
                        'sections': file_info['sections']
                    })
            
            # 输出处理结果信息
            print(f"\n处理完成，共有 {len(processed_files)} 个文件:")
            for i, file_info in enumerate(processed_files):
                print(f"  {i+1}. {os.path.basename(file_info['file_path'])} ({file_info['orientation']})")
            
            # 合并处理后的文件
            print("\n开始合并处理后的文件...")
            
            # 创建合并输出目录
            merge_output_dir = os.path.join(self.temp_dir, "merged")
            os.makedirs(merge_output_dir, exist_ok=True)
            
            # 初始化最终输出文件路径
            final_output_path = os.path.join(merge_output_dir, "merged_document.docx")
            
            # 合并文件
            # 如果只有一个文件，直接复制
            if len(processed_files) == 1:
                shutil.copy2(processed_files[0]['file_path'], final_output_path)
                print(f"只有一个处理后的文件，已复制到: {final_output_path}")
            else:
                # 多个文件需要合并
                # 以第一个文件作为基础文件
                base_file = processed_files[0]['file_path']
                
                # 逐个合并后续文件
                current_merged_file = base_file
                temp_files = []
                
                del processor

                for i in range(1, len(processed_files)):
                    file_to_append = processed_files[i]['file_path']
                    temp_output_file = os.path.join(merge_output_dir, f"temp_merged_{i}.docx")
                    temp_files.append(temp_output_file)
                    
                    processor = DocumentOrientationProcessor()
                    # 合并文件并保持格式
                    success = processor.merge_documents_with_formatting(
                        current_merged_file, 
                        file_to_append, 
                        temp_output_file
                    )
                    
                    if success:
                        print(f"成功合并第 {i} 个文件: {os.path.basename(file_to_append)}")
                        current_merged_file = temp_output_file
                    else:
                        print(f"合并第 {i} 个文件失败: {os.path.basename(file_to_append)}")
                        break
                
                # 将最终结果复制到目标位置
                if os.path.exists(current_merged_file):
                    shutil.copy2(current_merged_file, final_output_path)
                    print(f"文档合并完成，输出文件: {final_output_path}")
                
                # 清理临时文件
                for temp_file in temp_files:
                    if os.path.exists(temp_file):
                        try:
                            os.remove(temp_file)
                        except Exception:
                            pass

            # 最后清理空节
            final_cleaned_path = os.path.join(merge_output_dir, "merged_document_no_empty_sections.docx")
            if os.path.exists(final_output_path):
                # 重新创建 DocumentOrientationProcessor 实例以确保可用
                try:
                    processor = DocumentOrientationProcessor()
                    success = processor.remove_empty_sections(final_output_path, final_cleaned_path)
                    if success and os.path.exists(final_cleaned_path):
                        print(f"✅ 清理空节完成，输出文件: {final_cleaned_path}")
                        # 保存中间文件
                        if self.save_intermediate_files:
                            self._save_intermediate_file(final_cleaned_path, "step1_6_operation", "清理空节后")
                        return final_cleaned_path
                    else:
                        print("⚠️ 清理空节失败，使用未清理的合并文件")
                        # 保存中间文件
                        if self.save_intermediate_files:
                            self._save_intermediate_file(final_output_path, "step1_6_operation", "合并后")
                        return final_output_path
                except Exception as clean_error:
                    print(f"⚠️ 清理空节时发生错误: {str(clean_error)}")
                    print("⚠️ 跳过清理空节步骤，使用未清理的合并文件")
                    # 保存中间文件
                    if self.save_intermediate_files:
                        self._save_intermediate_file(final_output_path, "step1_6_operation", "合并后")
                    return final_output_path
            else:
                print("合并文件不存在，跳过清理空节步骤")
                return None
                
        except Exception as e:
            print(f"❌ 步骤1.6处理过程中发生错误: {str(e)}")
            import traceback
            traceback.print_exc()
            return None

    def step12_remove_duplicate_tables(self, doc_path: str) -> str:
        """
        Step 12 预处理：删除与首张表格内容相同的后续表格。
        """
        print("-" * 50)
        print("步骤12-预处理: 删除重复表格")

        if not doc_path or not os.path.exists(doc_path):
            print("步骤12-预处理: 未找到目标文档，跳过表格去重。")
            return doc_path

        try:
            import docx
        except ImportError:
            print("步骤12-预处理: 未安装 python-docx，无法执行表格去重。")
            return doc_path

        try:
            document = docx.Document(doc_path)
        except Exception as err:
            print(f"步骤12-预处理: 打开文档失败，原因: {err}")
            return doc_path

        tables = document.tables
        if len(tables) <= 1:
            print("步骤12-预处理: 表格数量不足，无需去重。")
            return doc_path

        try:
            first_signature = self._table_signature(tables[0])
        except Exception as err:
            print(f"步骤12-预处理: 生成首张表格指纹失败，原因: {err}")
            return doc_path

        removed_count = 0
        for table in list(tables)[1:]:
            try:
                if self._table_signature(table) == first_signature:
                    parent = table._element.getparent()
                    parent.remove(table._element)
                    removed_count += 1
            except Exception as err:
                print(f"步骤12-预处理: 表格去重时发生错误，原因: {err}")

        if removed_count == 0:
            print("步骤12-预处理: 未发现重复表格。")
            return doc_path

        try:
            document.save(doc_path)
            print(f"步骤12-预处理: 已删除 {removed_count} 个重复表格。")
        except Exception as err:
            print(f"步骤12-预处理: 保存文档失败，原因: {err}")
            return doc_path

        self.intermediate_files['step12_table_dedup'] = doc_path
        if self.save_intermediate_files:
            self._save_intermediate_file(doc_path, "step12_table", "去重后")

        return doc_path

    def step12_remove_duplicate_phrase(self, doc_path: str, phrase: str = "各专业参加设计人员名单") -> str:
        """
        Step 12 预处理：删除重复出现的特定段落。
        """
        print("-" * 50)
        print("步骤12-预处理: 删除重复段落")

        if not doc_path or not os.path.exists(doc_path):
            print("步骤12-预处理: 未找到目标文档，跳过段落去重。")
            return doc_path

        try:
            import docx
        except ImportError:
            print("步骤12-预处理: 未安装 python-docx，无法执行段落去重。")
            return doc_path

        try:
            document = docx.Document(doc_path)
        except Exception as err:
            print(f"步骤12-预处理: 打开文档失败，原因: {err}")
            return doc_path

        seen = False
        removed = 0

        for para in list(document.paragraphs):
            if para.text.strip() != phrase:
                continue

            if not seen:
                seen = True
                continue

            parent = para._element.getparent()
            parent.remove(para._element)
            removed += 1

        if removed == 0:
            print("步骤12-预处理: 未发现重复段落。")
            return doc_path

        try:
            document.save(doc_path)
            print(f"步骤12-预处理: 已删除 {removed} 个重复段落。")
        except Exception as err:
            print(f"步骤12-预处理: 保存文档失败，原因: {err}")
            return doc_path

        self.intermediate_files['step12_phrase_dedup'] = doc_path
        if self.save_intermediate_files:
            self._save_intermediate_file(doc_path, "step12_phrase", "段落去重后")

        return doc_path

    def step13_format_landscape_headers(
            self,
            doc_path: str,
            header_left_text: str,
            header_right_text: str
    ) -> bool:
        """
        步骤13: 处理所有横板页眉，设置同一行左右对齐文本。
        """
        print("-" * 50)
        print("步骤13: 横板页眉左右对齐")

        if not doc_path or not os.path.exists(doc_path):
            print("步骤13: 未找到目标文档，跳过横板页眉处理。")
            return False

        try:
            from docx import Document
            from docx.enum.section import WD_ORIENT
            from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
            from docx.oxml import OxmlElement
            from docx.oxml.ns import qn
        except ImportError as import_error:
            print(f"步骤13: 未安装 python-docx，跳过横板页眉处理。原因: {import_error}")
            return False

        import re

        try:
            document = Document(doc_path)
        except Exception as err:
            print(f"步骤13: 打开文档失败，原因: {err}")
            return False

        landscape_headers = []
        for section in document.sections:
            is_landscape = False
            try:
                if section.orientation == WD_ORIENT.LANDSCAPE:
                    is_landscape = True
                elif section.page_width > section.page_height:
                    is_landscape = True
            except Exception:
                is_landscape = False

            if not is_landscape:
                continue

            headers = [("default", section.header)]
            if getattr(section, "different_first_page_header_footer", False):
                headers.append(("first", section.first_page_header))
            if getattr(section, "odd_and_even_pages_header_footer", False):
                headers.append(("even", section.even_page_header))
            landscape_headers.append((section, headers))

        if not landscape_headers:
            print("步骤13: 未找到横板节，跳过横板页眉处理。")
            return False

        def _normalize_header_text(text: str) -> str:
            if not text:
                return ""
            cleaned = re.sub(r"[\s\u00A0\u3000]+", " ", text)
            return cleaned.strip()

        left_text = _normalize_header_text(header_left_text)
        right_text = _normalize_header_text(header_right_text)
        if not left_text and not right_text:
            print("步骤13: 未提供页眉左右文本，尝试自动提取（表格/Tab/空格）。")
            for _, headers in landscape_headers:
                for header_name, header in headers:
                    table_left = ""
                    table_right = ""
                    for table in header.tables:
                        if not table.rows:
                            continue
                        cells = table.rows[0].cells
                        if len(cells) < 2:
                            continue
                        table_left = (cells[0].text or "").strip()
                        table_right = (cells[1].text or "").strip()
                        if table_left or table_right:
                            break

                    tab_left = ""
                    tab_right = ""
                    for paragraph in header.paragraphs:
                        text = paragraph.text or ""
                        if "\t" not in text:
                            continue
                        parts = text.split("\t", 1)
                        tab_left = parts[0].strip()
                        tab_right = parts[1].strip()
                        if tab_left or tab_right:
                            break

                    print(
                        f"步骤13: {header_name}页眉-表格提取: '{table_left}' | '{table_right}'"
                    )
                    print(
                        f"步骤13: {header_name}页眉-Tab提取: '{tab_left}' | '{tab_right}'"
                    )

                    space_left = ""
                    space_right = ""
                    for paragraph in header.paragraphs:
                        text = (paragraph.text or "").strip()
                        if not text:
                            continue
                        parts = re.split(r"\s{2,}", text, maxsplit=1)
                        if len(parts) != 2:
                            continue
                        space_left = parts[0].strip()
                        space_right = parts[1].strip()
                        if space_left or space_right:
                            break

                    print(
                        f"步骤13: {header_name}页眉-空格提取: '{space_left}' | '{space_right}'"
                    )

                    if not left_text and not right_text:
                        if table_left or table_right:
                            left_text = _normalize_header_text(table_left)
                            right_text = _normalize_header_text(table_right)
                        elif tab_left or tab_right:
                            left_text = _normalize_header_text(tab_left)
                            right_text = _normalize_header_text(tab_right)
                        elif space_left or space_right:
                            left_text = _normalize_header_text(space_left)
                            right_text = _normalize_header_text(space_right)

                if left_text or right_text:
                    break

            if not left_text and not right_text:
                print("步骤13: 自动提取失败，跳过横板页眉处理。")
                return False

        updated_headers = 0
        for section, headers in landscape_headers:
            for _, header in headers:
                if header.is_linked_to_previous:
                    header.is_linked_to_previous = False

                target_paragraph = None
                for paragraph in header.paragraphs:
                    text = (paragraph.text or "").strip()
                    if not text:
                        continue
                    if left_text in text or right_text in text:
                        target_paragraph = paragraph
                        break

                if target_paragraph is None:
                    print("步骤13: 未找到包含目标文本的页眉段落，跳过该页眉。")
                    continue

                for paragraph in list(header.paragraphs):
                    if paragraph is target_paragraph:
                        continue
                    has_drawing = bool(
                        paragraph._element.xpath(
                            ".//*[local-name()='drawing'] | .//*[local-name()='pict']"
                        )
                    )
                    if not has_drawing:
                        paragraph._element.getparent().remove(paragraph._element)

                for child in list(target_paragraph._element):
                    if child.tag != qn("w:pPr"):
                        target_paragraph._element.remove(child)

                target_paragraph.paragraph_format.space_before = 0
                target_paragraph.paragraph_format.space_after = 0
                target_paragraph.paragraph_format.line_spacing = 1
                target_paragraph.paragraph_format.left_indent = 0
                target_paragraph.paragraph_format.right_indent = 0
                target_paragraph.paragraph_format.first_line_indent = 0
                target_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

                ppr_list = target_paragraph._element.xpath(".//*[local-name()='pPr']")
                if ppr_list:
                    ppr = ppr_list[0]
                    for elem_name in ("tabs", "ind", "jc"):
                        for elem in list(ppr.xpath(f".//*[local-name()='{elem_name}']")):
                            ppr.remove(elem)
                else:
                    ppr = OxmlElement("w:pPr")
                    target_paragraph._element.insert(0, ppr)

                usable_width = section.page_width - section.left_margin - section.right_margin
                tab_pos = int(usable_width.twips) if hasattr(usable_width, "twips") else int(usable_width)
                tabs = OxmlElement("w:tabs")
                tab = OxmlElement("w:tab")
                tab.set(qn("w:val"), "right")
                tab.set(qn("w:pos"), str(tab_pos))
                tabs.append(tab)
                ppr.append(tabs)

                if left_text:
                    run = OxmlElement("w:r")
                    text = OxmlElement("w:t")
                    text.text = left_text
                    run.append(text)
                    target_paragraph._element.append(run)

                if right_text:
                    run_tab = OxmlElement("w:r")
                    run_tab.append(OxmlElement("w:tab"))
                    target_paragraph._element.append(run_tab)

                    run = OxmlElement("w:r")
                    text = OxmlElement("w:t")
                    text.text = right_text
                    run.append(text)
                    target_paragraph._element.append(run)

                updated_headers += 1

        if updated_headers == 0:
            print("步骤13: 未找到横板页眉，跳过横板页眉处理。")
            return False

        try:
            if not self.temp_dir:
                raise ValueError("临时目录未初始化")
            base_name = os.path.splitext(os.path.basename(doc_path))[0]
            formatted_output = os.path.join(self.temp_dir, f"{base_name}_step13_landscape_headers.docx")
            document.save(formatted_output)
        except Exception as err:
            print(f"步骤13: 保存文档失败，原因: {err}")
            return False

        self.intermediate_files["step13_landscape_headers"] = formatted_output
        if self.save_intermediate_files:
            self._save_intermediate_file(formatted_output, "step13_landscape_headers", "横板页眉左右对齐")

        try:
            shutil.copy2(formatted_output, doc_path)
        except Exception as err:
            print(f"步骤13: 覆盖原文档失败，原因: {err}")
            return False

        print(f"步骤13: 横板页眉处理完成，更新页眉数: {updated_headers}")
        return True

    def step1_split_document(self, source_file: str) -> Tuple[Optional[str], Optional[str], Optional[str], Optional[str]]:
        """
        步骤1: 文档拆分 - 将源文档分离为封面和正文

        Args:
            source_file: 源文档路径

        Returns:
            Tuple[str, str]: (封面文件路径, 不含目录的正文内容文件路径)

        说明:
            - 封面文档：只包含目录之前的内容（如封面、标题等），不包含目录本身
            - 正文内容文档：包含从目录之后开始的所有内容（不包括目录）
        """
        print("-" * 50)
        print("📑 步骤1: 文档拆分")

        try:
            # 确保临时目录存在
            project_name = None
            document_type_info = None
            if not self.temp_dir:
                raise ValueError("临时目录未初始化")

            # 生成输出文件路径
            base_name = os.path.splitext(os.path.basename(source_file))[0]
            cover_toc_path = os.path.join(self.temp_dir, f"{base_name}_封面.docx")
            content_no_toc_path = os.path.join(self.temp_dir, f"{base_name}_正文内容.docx")

            print(f"📄 源文档: {os.path.basename(source_file)}")
            print(f"📤 封面输出: {os.path.basename(cover_toc_path)}")
            print(f"📤 正文内容输出: {os.path.basename(content_no_toc_path)}")

            # 使用整合的拆分方法
            processor = DocxSplitProcessor()
            success = processor.split_document_for_conversion(
                source_path=source_file,
                cover_toc_output=cover_toc_path,
                content_no_toc_output=content_no_toc_path,
                toc_keywords=['目录', '目 录', '目  录', '目   录', '目    录', '目     录', '目      录',
                              '目       录', '目        录', '目         录', 'Contents', 'TABLE OF CONTENTS']
            )

            if success and os.path.exists(cover_toc_path) and os.path.exists(content_no_toc_path):
                print("✅ 文档拆分成功!")

                # 使用cover_replace.py中的便捷函数处理封面文档
                try:
                    # 导入cover_replace模块
                    from utils.cover_replace import replace_content_in_cover_auto
                    import json
                    from datetime import datetime  # 导入datetime模块

                    # 配置文件路径
                    parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

                    # 生成处理后的封面路径
                    processed_cover_path = os.path.join(self.temp_dir, f"{base_name}_封面_处理后.docx")

                    # 使用自动选择模板和配置文件的函数处理封面
                    print("🔧 使用cover_replace_auto处理封面文档...")
                    # actual_path = replace_content_in_cover_auto(
                    #     source_docx_path=cover_toc_path,  # 使用拆分后的封面作为源
                    #     output_docx_path=processed_cover_path,
                    #     document_type=self.document_type,  # 使用文档类型参数
                    #     save_file=self.save_intermediate_files  # 与convert_document的save_intermediate参数关联
                    # )

                    actual_path, project_name, document_type_info = replace_content_in_cover_by_dify(
                        source_docx_path=cover_toc_path,  # 使用拆分后的封面作为源
                        output_docx_path=processed_cover_path,
                        document_type=self.document_type,  # 使用文档类型参数
                        save_file=self.save_intermediate_files  # 与convert_document的save_intermediate参数关联
                    )                    

                    # 如果处理成功，更新cover_toc_path指向处理后的文件
                    if os.path.exists(actual_path):
                        cover_toc_path = actual_path
                        print(f"✅ 封面文档处理成功，使用处理后的文件: {os.path.basename(cover_toc_path)}")

                        # 如果需要保存中间文件，也将处理后的封面文件复制到调试目录
                        if self.save_intermediate_files:
                            processed_cover_filename = f"step1_split_封面处理后_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
                            processed_cover_debug_path = os.path.join(self.debug_output_dir, processed_cover_filename)
                            shutil.copy2(cover_toc_path, processed_cover_debug_path)
                            print(f"   💾 已保存处理后的封面文件: {processed_cover_filename}")
                    else:
                        print("⚠️ 封面文档处理失败，使用原始拆分文件")
                except Exception as cover_error:
                    print(f"⚠️ 封面文档处理过程中发生错误: {str(cover_error)}")
                    print("   继续使用原始拆分文件")

                # 使用cover_table_replace.py中的函数替换处理后封面中的表格
                try:
                    from utils.cover_table_replace import replace_table_after_marker
                    from datetime import datetime  # 导入datetime模块

                    # 生成表格替换后的封面路径
                    table_replaced_cover_path = os.path.join(self.temp_dir, f"{base_name}_封面_表格替换后.docx")

                    # 使用源文档作为表格来源，处理后的封面作为目标进行表格替换
                    print("🔧 使用cover_table_replace替换处理后封面中的表格...")
                    replaced_path = replace_table_after_marker(
                        source_path=source_file,  # 使用原始源文档作为表格来源
                        target_path=cover_toc_path,  # 使用处理后的封面作为替换目标
                        marker="设计人员名单",  # 2026年1月23日改 适配 “各专业主要参加设计人员名单” 和 “各专业参加设计人员名单”
                        save_path=table_replaced_cover_path
                    )

                    # 如果替换成功，更新cover_toc_path指向表格替换后的文件
                    if os.path.exists(replaced_path):
                        cover_toc_path = replaced_path
                        print(f"✅ 封面表格替换成功，使用表格替换后的文件: {os.path.basename(cover_toc_path)}")

                        # 如果需要保存中间文件，也将表格替换后的封面文件复制到调试目录
                        if self.save_intermediate_files:
                            table_replaced_cover_filename = f"step1_split_封面_表格替换后_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
                            table_replaced_cover_debug_path = os.path.join(self.debug_output_dir,
                                                                           table_replaced_cover_filename)
                            shutil.copy2(cover_toc_path, table_replaced_cover_debug_path)
                            print(f"   💾 已保存表格替换后的封面文件: {table_replaced_cover_filename}")
                    else:
                        print("⚠️ 封面表格替换失败，使用处理后的封面文件")
                except Exception as table_error:
                    print(f"⚠️ 封面表格替换过程中发生错误: {str(table_error)}")
                    print("   继续使用处理后的封面文件")

                # 对正文内容文档中的Excel表格进行转换处理
                try:
                    from utils.docx_table_excel import convert_embedded_excels_inplace
                    from datetime import datetime  # 导入datetime模块

                    # 生成处理后的正文内容路径
                    processed_content_path = os.path.join(self.temp_dir, f"{base_name}_正文内容_表格处理后.docx")

                    # 使用docx_table_excel处理正文中的Excel表格
                    print("🔧 使用docx_table_excel处理正文内容中的Excel表格...")
                    try:
                        converted_count = convert_embedded_excels_inplace(
                            source_docx=content_no_toc_path,
                            output_docx=processed_content_path,
                            placeholder_when_no_pandas=True
                        )

                        # 如果处理成功，更新content_no_toc_path指向处理后的文件
                        if os.path.exists(processed_content_path) and converted_count >= 0:
                            content_no_toc_path = processed_content_path
                            print(f"✅ 正文内容中的Excel表格处理成功，共转换 {converted_count} 个表格")

                            # 如果需要保存中间文件，也将处理后的正文文件复制到调试目录
                            if self.save_intermediate_files:
                                processed_content_filename = f"step1_split_正文内容_表格处理后_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
                                processed_content_debug_path = os.path.join(self.debug_output_dir,
                                                                            processed_content_filename)
                                shutil.copy2(content_no_toc_path, processed_content_debug_path)
                                print(f"   💾 已保存处理后的正文内容文件: {processed_content_filename}")
                        else:
                            print("⚠️ 正文内容中的Excel表格处理失败，使用原始正文文件")
                    except Exception as convert_error:
                        print(f"⚠️ 正文内容中的Excel表格转换失败: {str(convert_error)}")
                        print("   继续使用原始正文文件")
                except Exception as convert_init_error:
                    print(f"⚠️ 正文内容中的Excel表格处理初始化失败: {str(convert_init_error)}")
                    print("   继续使用原始正文文件")
                # 无论 Dify 成功与否，都要对最终的 cover_toc_path 执行字间距调整
                # 注意：过大的字间距会导致“可行性研究/初步设计”被拉开。
                try:
                    from docx import Document
                    doc = Document(cover_toc_path)
                    # 保持封面模板中的字间距（默认 7 磅），不再重置为 0
                    apply_character_spacing_to_cover(doc, spacing_pt=7)
                    doc.save(cover_toc_path)
                    print("✅ 封面字间距调整完成(已重置为0)")
                except Exception as e:
                    print(f"⚠️ 封面字间距调整失败: {e}")

                # 保存中间文件到指定目录便于查看调试
                if self.save_intermediate_files:
                    print(f"   封面: {os.path.basename(cover_toc_path)}")
                    print(f"   正文内容: {os.path.basename(content_no_toc_path)}")
                    print(f"📁 正在保存step1中间文件到: {self.debug_output_dir}")
                    self._save_intermediate_file(cover_toc_path, "step1_split", "封面")
                    self._save_intermediate_file(content_no_toc_path, "step1_split", "正文内容")

                # 保存中间文件路径
                self.intermediate_files['cover_toc'] = cover_toc_path
                self.intermediate_files['original_content'] = content_no_toc_path

                return cover_toc_path, content_no_toc_path, project_name, document_type_info
            else:
                print("❌ 文档拆分失败")
                return None, None, None, None

        except Exception as e:
            print(f"❌ 文档拆分过程中发生错误: {str(e)}")
            return None, None, None, None

    def _create_orientation_specific_copy(
        self,
        source_content_path: str,
        output_path: str,
        target_orientation: int,
        word: None
    ) -> bool:
        """使用 Word COM 仅保留指定方向的节并另存为新文档。"""
        if not COM_AVAILABLE or not win32 or not pythoncom or C is None:
            return False

        # word_app = word
        cloned_doc = None
        com_initialized = False
        kept_sections = 0

        source_content_path = os.path.abspath(source_content_path)
        output_path = os.path.abspath(output_path)

        if not os.path.exists(source_content_path):
            return False

        try:
            shutil.copy2(source_content_path, output_path)
        except Exception as copy_error:
            print(f"⚠️ 正文克隆失败: {copy_error}")
            return False

        try:
            pythoncom.CoInitialize()
            com_initialized = True

            # word_app = win32.DispatchEx("Word.Application")
            # word_app.Visible = False
            # word_app.DisplayAlerts = False

            cloned_doc = word.Documents.Open(output_path, ReadOnly=False)
            sections_count = cloned_doc.Sections.Count if cloned_doc else 0

            for index in range(sections_count, 0, -1):
                section = cloned_doc.Sections(index)
                orientation = section.PageSetup.Orientation
                if orientation != target_orientation:
                    section.Range.Delete()
                else:
                    kept_sections += 1

            remaining_sections = cloned_doc.Sections.Count if cloned_doc else 0

            if kept_sections > 0 and remaining_sections > 0:
                header_types = [
                    getattr(C, "wdHeaderFooterPrimary", 1),
                    getattr(C, "wdHeaderFooterFirstPage", 2),
                    getattr(C, "wdHeaderFooterEvenPages", 3),
                ]
                footer_types = [
                    getattr(C, "wdHeaderFooterPrimary", 1),
                    getattr(C, "wdHeaderFooterFirstPage", 2),
                    getattr(C, "wdHeaderFooterEvenPages", 3),
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
                    print(f"ℹ️ 取消页眉页脚链接时出现问题: {link_error}")

            content_length = cloned_doc.Content.StoryLength if cloned_doc else 0

            if kept_sections > 0 and content_length > 1:
                cloned_doc.Save()
                return True

            if os.path.exists(output_path):
                try:
                    os.remove(output_path)
                except Exception:
                    pass

            return False
        except Exception as exc:
            print(f"⚠️ 页面方向内容裁剪失败: {exc}")
            return False
        finally:
            if cloned_doc is not None:
                try:
                    cloned_doc.Close(SaveChanges=True)
                except Exception:
                    pass
            # if word_app is not None:
            #     try:
            #         word_app.Quit()
            #     except Exception:
            #         pass
            # if COM_AVAILABLE and com_initialized:
            #     try:
            #         pythoncom.CoUninitialize()
            #     except Exception:
            #         pass

            if (kept_sections <= 0 or not os.path.exists(output_path) or os.path.getsize(output_path) == 0):
                if os.path.exists(output_path):
                    try:
                        os.remove(output_path)
                    except Exception:
                        pass

    def _split_content_by_orientation_com(
        self,
        source_content_path: str,
        portrait_output_path: str,
        landscape_output_path: str,
        word=None
    ) -> Tuple[bool, bool]:
        """基于 Word COM 按页面方向拆分正文内容。"""
        print("🔧 开始基于 COM API 按页面方向拆分正文内容...")

        if not COM_AVAILABLE:
            print("❌ COM 库不可用，无法进行页面方向拆分")
            return False, False

        if not source_content_path or not os.path.exists(source_content_path):
            print("⚠️ 正文内容文件不存在，跳过页面方向拆分")
            return False, False

        portrait_success = self._create_orientation_specific_copy(
            source_content_path=source_content_path,
            output_path=portrait_output_path,
            target_orientation=getattr(C, "wdOrientPortrait", 0),
            word=word
        )
        # landscape_success = True
        landscape_success = self._create_orientation_specific_copy(
            source_content_path=source_content_path,
            output_path=landscape_output_path,
            target_orientation=getattr(C, "wdOrientLandscape", 1),
            word=word
        )

        if not portrait_success and os.path.exists(portrait_output_path):
            try:
                os.remove(portrait_output_path)
            except Exception:
                pass

        if not landscape_success and os.path.exists(landscape_output_path):
            try:
                os.remove(landscape_output_path)
            except Exception:
                pass

        if portrait_success:
            print(f"   ✅ 已保存竖版文件: {os.path.basename(portrait_output_path)}")
        else:
            print("ℹ️ 竖版节不存在或处理失败，继续使用原始正文文件")

        if landscape_success:
            print(f"   ✅ 已保存横版文件: {os.path.basename(landscape_output_path)}")
        else:
            print("ℹ️ 横版节不存在或处理失败，不生成横版文件")

        return portrait_success, landscape_success

    def _split_document_sections(self, source_path: str, prefix: str, word=None) -> List[str]:
        """
        使用 Word COM 将文档按节拆分成多个子文档。
        """
        results: List[str] = []
        if not (COM_AVAILABLE and pythoncom and win32 and C is not None):
            return results
        if not source_path:
            return results

        source_path = os.path.abspath(source_path)
        if not os.path.exists(source_path):
            return results
        if not self.temp_dir:
            return results

        # word_app = None
        original_doc = None
        com_initialized = False

        try:
            # pythoncom.CoInitialize()
            # com_initialized = True

            # word_app = win32.DispatchEx("Word.Application")
            # word_app.Visible = False
            # word_app.DisplayAlerts = False

            original_doc = word.Documents.Open(source_path, ReadOnly=True)
            sections_count = original_doc.Sections.Count if original_doc else 0
            if original_doc:
                original_doc.Close(SaveChanges=0)
                original_doc = None

            for idx in range(1, sections_count + 1):
                target_path = os.path.join(self.temp_dir, f"{prefix}_{idx:02d}.docx")
                try:
                    shutil.copy2(source_path, target_path)
                except Exception as copy_err:
                    print(f"[warn] step1.5: failed to copy section document: {copy_err}")
                    if os.path.exists(target_path):
                        try:
                            os.remove(target_path)
                        except Exception:
                            pass
                    continue

                cloned_doc = None
                try:
                    cloned_doc = word.Documents.Open(target_path, ReadOnly=False)
                    for sec_idx in range(cloned_doc.Sections.Count, 0, -1):
                        if sec_idx != idx:
                            cloned_doc.Sections(sec_idx).Range.Delete()

                    if cloned_doc.Content.StoryLength > 1:
                        cloned_doc.Save()
                        results.append(target_path)
                    else:
                        cloned_doc.Close(SaveChanges=0)
                        cloned_doc = None
                        if os.path.exists(target_path):
                            try:
                                os.remove(target_path)
                            except Exception:
                                pass
                        continue
                except Exception as sec_err:
                    print(f"[warn] step1.5: section split error: {sec_err}")
                    if cloned_doc is not None:
                        try:
                            cloned_doc.Close(SaveChanges=0)
                        except Exception:
                            pass
                    if os.path.exists(target_path):
                        try:
                            os.remove(target_path)
                        except Exception:
                            pass
                else:
                    if cloned_doc is not None:
                        try:
                            cloned_doc.Close(SaveChanges=0)
                        except Exception:
                            pass
        except Exception as exc:
            print(f"[warn] step1.5: unexpected error while splitting sections: {exc}")
            for path_to_remove in results:
                if os.path.exists(path_to_remove):
                    try:
                        os.remove(path_to_remove)
                    except Exception:
                        pass
            results.clear()
        finally:
            if original_doc is not None:
                try:
                    original_doc.Close(SaveChanges=0)
                except Exception:
                    pass
            # if word_app is not None:
            #     try:
            #         word_app.Quit()
            #     except Exception:
            #         pass
            # if com_initialized:
            #     try:
            #         pythoncom.CoUninitialize()
            #     except Exception:
            #         pass

        return results

    def step1_5_split_by_orientation(self, content_file: str, word=None) -> Tuple[str, List[str]]:
        """
        Step 1.5: Split content into portrait and landscape parts using COM automation.

        Args:
            content_file: Path to the content document produced by step 1.

        Returns:
            Tuple[str, List[str]]: (portrait content path, list of landscape section paths)
        """
        print("-" * 50)
        print("🧭 步骤1.5: 按页面方向拆分正文内容")

        portrait_result = content_file
        landscape_sections: List[str] = []

        try:
            if not content_file or not os.path.exists(content_file):
                print("⚠️ 正文内容文件不存在，跳过页面方向拆分")
                return portrait_result, landscape_sections

            if not self.temp_dir:
                raise ValueError("临时目录未初始化")

            base_name = os.path.splitext(os.path.basename(content_file))[0]
            portrait_path = os.path.join(self.temp_dir, f"{base_name}_竖版内容.docx")
            landscape_path = os.path.join(self.temp_dir, f"{base_name}_横版内容.docx")

            portrait_success, landscape_success = self._split_content_by_orientation_com(
                source_content_path=content_file,
                portrait_output_path=portrait_path,
                landscape_output_path=landscape_path,
                word=word
            )

            if portrait_success and os.path.exists(portrait_path) and os.path.getsize(portrait_path) > 0:
                portrait_result = portrait_path
                print(f"✅ 将竖版文件用于后续处理: {os.path.basename(portrait_result)}")
                if self.save_intermediate_files:
                    self._save_intermediate_file(portrait_result, "step1_5_split", "竖版内容")
            else:
                if os.path.exists(portrait_path) and os.path.getsize(portrait_path) == 0:
                    try:
                        os.remove(portrait_path)
                    except Exception:
                        pass
                print("⚠️ 竖版拆分失败或无竖版内容，继续使用原始正文文件")

            if landscape_success and os.path.exists(landscape_path) and os.path.getsize(landscape_path) > 0:
                print(f"[info] step1.5: landscape file saved -> {os.path.basename(landscape_path)}")
                if self.save_intermediate_files:
                    self._save_intermediate_file(landscape_path, "step1_5_split", "landscape_full")
                sections = self._split_document_sections(
                    landscape_path,
                    f"{base_name}_landscape",
                    word=word
                )
                if sections:
                    landscape_sections = sections
                    if self.save_intermediate_files:
                        for idx, section_path in enumerate(landscape_sections, 1):
                            self._save_intermediate_file(section_path, "step1_5_split", f"landscape_section_{idx}")
                else:
                    landscape_sections.append(landscape_path)
            else:
                if os.path.exists(landscape_path) and os.path.getsize(landscape_path) == 0:
                    try:
                        os.remove(landscape_path)
                    except Exception:
                        pass
                print("[warn] step1.5: no landscape content found")


        except Exception as exc:
            print(f"⚠️ 步骤1.5 处理过程中发生错误: {exc}")
            portrait_result = content_file
            landscape_sections = []
        finally:
            self.intermediate_files['portrait_content'] = portrait_result
            if landscape_sections:
                self.intermediate_files['landscape_sections'] = ';'.join(landscape_sections)
            else:
                self.intermediate_files.pop('landscape_sections', None)

        return portrait_result, landscape_sections

    def step1_5_split_by_orientation_segments_pipeline(
        self,
        content_file: str,
        template_file: str,
        word=None
    ) -> Tuple[str, List[Dict[str, Any]], str, Dict[str, Dict[str, Any]]]:
        """
        Step 1.5 (new): Split by orientation segments, process portrait segments only,
        then merge back in original order.

        - Portrait segments: step1.8 -> step1.9 -> step2 -> step3
        - Landscape segments: skip step1.8/1.9/2/3, keep original
        """
        err_msg = ""
        print("-" * 50)
        print("🧭 步骤1.5: 按横竖段落拆分 + 分段处理 + 按序合并")

        import time
        t_global_start = time.time()

        if not content_file or not os.path.exists(content_file):
            print("⚠️ 正文内容文件不存在，跳过步骤1.5")
            return "", [], "正文内容文件不存在", {}

        if not self.temp_dir:
            raise ValueError("临时目录未初始化")

        segment_dir = os.path.join(self.temp_dir, "step1_5_orientation_segments")
        processed_dir = os.path.join(self.temp_dir, "step1_5_orientation_processed")
        merge_dir = os.path.join(self.temp_dir, "step1_5_orientation_merged")
        os.makedirs(segment_dir, exist_ok=True)
        os.makedirs(processed_dir, exist_ok=True)
        os.makedirs(merge_dir, exist_ok=True)

        t0 = time.time()
        processor = DocumentOrientationProcessor()
        segments = processor.split_document_by_orientation_segments(content_file, segment_dir)
        print(f"⏱ split_document_by_orientation_segments 耗时: {time.time() - t0:.2f}s")

        if not segments:
            print("⚠️ 未识别到横竖段落，使用原始正文文件作为竖版段落处理")
            segments = [{
                "file_path": content_file,
                "orientation": "竖版",
                "index": 0,
                "sections": "全部"
            }]
            return "", [], "未识别到横竖段落", {}

        segments.sort(key=lambda x: x.get("index", 0))

        # 保存拆分后的原始段落到调试目录
        if self.save_intermediate_files and segments:
             for i, seg in enumerate(segments, 1):
                s_path = seg.get("file_path")
                s_orient = seg.get("orientation", "unknown")
                if s_path and os.path.exists(s_path):
                     self._save_intermediate_file(
                        s_path,
                        "step1_5_raw",
                        f"segment_{i}_{s_orient}"
                    )

        processed_segments: List[Dict[str, Any]] = []
        formula_info_all: List[Dict[str, Any]] = []
        formula_uid_counter = 0
        step1_9_run_dirs: List[str] = []
        pic_uid_counter = 0
        picture_info_all = {}
        for i, segment in enumerate(segments, 1):
            t_seg_start = time.time()
            seg_path = segment.get("file_path")
            orientation = segment.get("orientation", "未知")
            seg_index = segment.get("index", i - 1)

            if not seg_path or not os.path.exists(seg_path):
                print(f"⚠️ 段落文件不存在，跳过: index={seg_index}")
                return "", [], f"段落文件不存在", {}

            print(f"\n  处理段落 {i}: {os.path.basename(seg_path)} ({orientation})")

            if orientation == "横版":
                # 横版跳过 Pandoc 和表格替换
                target_path = os.path.join(processed_dir, f"segment_{i:02d}_横版.docx")
                try:
                    if os.path.abspath(seg_path) != os.path.abspath(target_path):
                        shutil.copy2(seg_path, target_path)
                    print("    → 横版段落保留原样")
                    processed_path = target_path
                except Exception as copy_error:
                    print(f"    ⚠️ 横版段落复制失败，使用原文件: {copy_error}")
                    processed_path = seg_path
                    return "", [], f"横版段落复制失败", {}
            else:
                # 竖版段落处理：1.8 -> 1.9 -> 2 -> 3
                step1_8_input_file = seg_path
                processed_path = seg_path

                try:
                    t1 = time.time()
                    processed_path, formula_info_list = self.step1_8_convert_equations(step1_8_input_file, word=word)
                    print(f"    ⏱ step1.8 耗时: {time.time() - t1:.2f}s")
                    if formula_info_list:
                        for info in formula_info_list:
                            formula_uid_counter += 1
                            info = dict(info)
                            info["uid"] = f"FORMULA_{formula_uid_counter}"
                            formula_info_all.append(info)
                except Exception as step18_error:
                    print(f"    ⚠️ step1.8 处理失败，继续使用原文件: {step18_error}")
                    processed_path = step1_8_input_file
                    return "", [], f"公式图片处理失败", {}


                try:
                    t2 = time.time()
                    extracted_doc = self.step1_9_extract_and_map(step1_8_input_file, word=word)
                    print(f"    ⏱ step1.9 耗时: {time.time() - t2:.2f}s")

                    # 收集 step1.9 的运行目录以便后续合并
                    current_run_dir = self.intermediate_files.get("step1_9_run_dir")
                    if current_run_dir and os.path.exists(current_run_dir):
                        step1_9_run_dirs.append(current_run_dir)

                    if extracted_doc and os.path.exists(extracted_doc) and extracted_doc != step1_8_input_file:
                        processed_path = extracted_doc
                    else:
                        print("    ⚠️ extract_and_map 未生成有效输出，继续使用 step1.8 输出文件")
                        return "", [], f"工艺流程图输出异常", {}
                except Exception as step19_error:
                    print(f"    ⚠️ step1.9 处理失败，继续使用 step1.8 输出文件: {step19_error}")
                    return "", [], f"工艺流程图处理失败", {}

                print("!!!!!!!!!!!!!!!!!!!!!ceshi!!!!!!!!!!!!!!!!!!")
                # from utils.picture_format_keeper import extract_and_mark_pictures
                # processed_path, picture_info_list, pic_uid_counter = extract_and_mark_pictures(
                #     processed_path,
                #     start_index=pic_uid_counter + 1
                # )
                #
                # picture_info_all.update(picture_info_list)
                # print(picture_info_all)
                print("!!!!!!!!!!!!!!!!!!!!!ceshi!!!!!!!!!!!!!!!!!!")

                # Step2: Pandoc
                # 强制不生成目录，留待合并后统一生成
                t3 = time.time()
                pandoc_file = self.step2_pandoc_convert(processed_path, template_file, include_toc=False)
                print(f"    ⏱ step2(Pandoc) 耗时: {time.time() - t3:.2f}s")
                if not pandoc_file:
                    print("    ⚠️ Pandoc转换失败，使用原始竖版段落继续")
                    pandoc_file = processed_path
                    return "", [], f"Pandoc转换失败", picture_info_all

                # Step3: 表格格式化
                t4 = time.time()
                table_formatted_file = self.step3_format_tables(
                    content_file=pandoc_file,
                    template_file=template_file,
                    original_content_file=seg_path
                )
                print(f"    ⏱ step3(表格) 耗时: {time.time() - t4:.2f}s")
                if not table_formatted_file:
                    print("    ⚠️ 表格格式化失败，使用Pandoc输出文件继续")
                    table_formatted_file = pandoc_file
                    return "", [], f"表格格式化失败", {}

                processed_path = table_formatted_file

            if self.save_intermediate_files:
                self._save_intermediate_file(
                    processed_path,
                    "step1_5_processed",
                    f"segment_{i}_{orientation}"
                )

            processed_segments.append({
                "file_path": processed_path,
                "orientation": orientation,
                "index": seg_index,
            })

            print(f"⏱ 段落 {i} 总耗时: {time.time() - t_seg_start:.2f}s")

        if not processed_segments:
            print("⚠️ 未生成有效段落文件，回退使用原始正文文件")
            return "", [], "未生成有效段落文件", {}

        processed_segments.sort(key=lambda x: x.get("index", 0))

        # 合并处理后的段落，保持原顺序
        if len(processed_segments) == 1:
            merged_path = os.path.join(merge_dir, "segment_merged.docx")
            try:
                if os.path.abspath(processed_segments[0]["file_path"]) != os.path.abspath(merged_path):
                    shutil.copy2(processed_segments[0]["file_path"], merged_path)
                print(f"✅ 仅一个段落，直接作为合并结果: {merged_path}")
                return merged_path, formula_info_all, "", picture_info_all
            except Exception as copy_error:
                print(f"⚠️ 复制单段落文件失败，使用原文件: {copy_error}")
                return "", [], f"复制单段落文件失败", {}

        t_merge = time.time()

        base_file = processed_segments[0]["file_path"]
        current_merged = base_file
        temp_files: List[str] = []

        for i in range(1, len(processed_segments)):
            next_file = processed_segments[i]["file_path"]
            temp_output = os.path.join(merge_dir, f"segment_merge_{i:02d}.docx")
            temp_files.append(temp_output)

            success = processor.merge_documents_with_formatting(
                current_merged,
                next_file,
                temp_output
            )

            if not success:
                print(f"⚠️ 段落合并失败，尝试默认合并: {os.path.basename(next_file)}")
                success = processor._merge_documents_default(current_merged, next_file, temp_output)
                return "", [], f"段落合并失败", {}

            if success and os.path.exists(temp_output):
                current_merged = temp_output
                print(f"    ✓ 已合并段落 {i + 1}/{len(processed_segments)}")
            else:
                print(f"    ✗ 合并段落失败，停止合并: {os.path.basename(next_file)}")
                return "", [], f"段落合并失败", {}
                break

        # 清理临时合并文件（保留最终结果）
        for temp_path in temp_files[:-1]:
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except Exception:
                    pass

        print(f"⏱ 所有段落合并耗时: {time.time() - t_merge:.2f}s")

        if current_merged and os.path.exists(current_merged):
            final_merged = os.path.join(merge_dir, "segment_merged.docx")
            if os.path.abspath(current_merged) != os.path.abspath(final_merged):
                try:
                    shutil.copy2(current_merged, final_merged)
                except Exception:
                    final_merged = current_merged
            print(f"✅ 段落按序合并完成: {final_merged}")

            # 合并 step1.9 的 image_map 和图片目录
            if step1_9_run_dirs:
                t_res_merge = time.time()
                print("🔄 正在合并 step1.9 的图片映射资源...")
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                merged_run_dir = os.path.join(self.debug_output_dir, f"step1_9_merged_resources_{timestamp}")
                os.makedirs(merged_run_dir, exist_ok=True)

                merged_map = {}
                copied_images_count = 0

                for r_dir in step1_9_run_dirs:
                    # 合并 map
                    map_file = os.path.join(r_dir, "image_map.json")
                    if os.path.exists(map_file):
                        try:
                            with open(map_file, "r", encoding="utf-8") as f:
                                part_map = json.load(f)
                                merged_map.update(part_map)
                        except Exception as e:
                            print(f"    ⚠️ 读取映射文件失败: {map_file}, {e}")
                            return "", [], f"读取映射文件失败", {}

                    # 复制图片
                    try:
                        for fname in os.listdir(r_dir):
                            if fname.lower().endswith(".png"):
                                src_img = os.path.join(r_dir, fname)
                                dst_img = os.path.join(merged_run_dir, fname)
                                shutil.copy2(src_img, dst_img)
                                copied_images_count += 1
                    except Exception as e:
                        print(f"    ⚠️ 复制图片失败: {r_dir}, {e}")
                        return "", [], f"复制图片失败", {}

                # 保存合并后的 map
                merged_map_path = os.path.join(merged_run_dir, "image_map.json")
                print(f"⏱ step1.9资源合并耗时: {time.time() - t_res_merge:.2f}s")
                try:
                    with open(merged_map_path, "w", encoding="utf-8") as f:
                        json.dump(merged_map, f, ensure_ascii=False, indent=4)

                    # 更新 intermediate_files，使 step11.5 能找到合并后的资源
                    self.intermediate_files["step1_9_run_dir"] = merged_run_dir
                    print(f"    ✅ 资源合并完成: {copied_images_count} 张图片, Map条目 {len(merged_map)}")
                except Exception as e:
                    print(f"    ⚠️ 保存合并映射失败: {e}")
                    return "", [], f"保存合并映射失败", {}

            print("FINAL sections:", len(Document(final_merged).sections))

            # 合并后全局校正方向（避免横竖顺序错乱）
            try:
                t_norm = time.time()
                normalized = self._normalize_section_orientation_by_segments(
                    final_merged,
                    processed_segments,
                    word=word
                )
                print(f"⏱ 横竖版校正耗时: {time.time() - t_norm:.2f}s")
                if not normalized:
                    print("⚠️ 全局方向校正未执行或失败，继续使用当前文档")
                    return "", [], f"全局方向校正未执行或失败", {}
            except Exception as normalize_error:
                print(f"⚠️ 全局方向校正异常，继续使用当前文档: {normalize_error}")
                return "", [], f"全局方向校正异常", {}

        # 合并后不再清理空节，避免破坏横竖版分界

            # 合并后为横版节强制复制页眉页脚（从上一节拷贝）
            try:
                t_header = time.time()
                copied = self._copy_headers_footers_to_landscape_sections(final_merged, word=word)
                print(f"⏱ 页眉页脚复制耗时: {time.time() - t_header:.2f}s")
                if not copied:
                    print("⚠️ 横版页眉页脚复制未执行或失败，继续使用当前文档")
                    return "", [], f"横版页眉页脚复制未执行或失败", {}
            except Exception as copy_error:
                print(f"⚠️ 横版页眉页脚复制异常，继续使用当前文档: {copy_error}")
                return "", [], f"横版页眉页脚复制异常", {}

            return final_merged, formula_info_all, "", picture_info_all
        print(f"🚀 step1.5 总耗时: {time.time() - t_global_start:.2f}s")
        print("⚠️ 合并未生成有效文件，回退使用原始正文文件")
        return "", [], f"合并未生成有效文件", {}

    def _copy_headers_footers_to_landscape_sections(self, doc_path: str, word=None) -> bool:
        """
        强制将上一节的页眉页脚复制到所有横版节，避免横版页眉页脚为空。
        """
        if not COM_AVAILABLE or not win32 or not pythoncom:
            print("⚠️ COM 不可用，无法复制横版页眉页脚")
            return False

        if not doc_path or not os.path.exists(doc_path):
            print("⚠️ 文档不存在，无法复制横版页眉页脚")
            return False

        doc_path = os.path.abspath(doc_path)
        own_word = False
        com_initialized = False
        doc = None

        header_types = [
            getattr(C, "wdHeaderFooterPrimary", 1),
            getattr(C, "wdHeaderFooterFirstPage", 2),
            getattr(C, "wdHeaderFooterEvenPages", 3),
        ]
        footer_types = [
            getattr(C, "wdHeaderFooterPrimary", 1),
            getattr(C, "wdHeaderFooterFirstPage", 2),
            getattr(C, "wdHeaderFooterEvenPages", 3),
        ]
        landscape_code = getattr(C, "wdOrientLandscape", 1)

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

            doc = word.Documents.Open(doc_path, ReadOnly=False)
            total_sections = doc.Sections.Count if doc else 0
            if total_sections <= 1:
                return True

            copied_sections = 0
            for idx in range(2, total_sections + 1):
                section = doc.Sections(idx)
                if section.PageSetup.Orientation != landscape_code:
                    continue

                prev_section = doc.Sections(idx - 1)
                for header_type in header_types:
                    try:
                        dest_header = section.Headers(header_type)
                        src_header = prev_section.Headers(header_type)
                        dest_header.LinkToPrevious = False
                        dest_header.Range.FormattedText = src_header.Range.FormattedText
                    except Exception:
                        continue

                for footer_type in footer_types:
                    try:
                        dest_footer = section.Footers(footer_type)
                        src_footer = prev_section.Footers(footer_type)
                        dest_footer.LinkToPrevious = False
                        dest_footer.Range.FormattedText = src_footer.Range.FormattedText
                    except Exception:
                        continue

                copied_sections += 1

            doc.Save()
            print(f"✅ 已为 {copied_sections} 个横版节复制页眉页脚")
            return True
        except Exception as exc:
            print(f"⚠️ 复制横版页眉页脚失败: {exc}")
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

    def _normalize_section_orientation_by_segments(
        self,
        doc_path: str,
        segments: List[Dict[str, Any]],
        word=None
    ) -> bool:
        """
        根据 segments 顺序强制校正最终文档每个节的方向。
        """
        if not COM_AVAILABLE or not win32 or not pythoncom or C is None:
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
                    getattr(C, "wdOrientLandscape", 1)
                    if orientation == "横版"
                    else getattr(C, "wdOrientPortrait", 0)
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

    def step1_4_split_by_orientation_segments(self, content_file: str) -> Tuple[str, List[dict]]:
        """
        Step 1.4: Split content into multiple files by continuous orientation segments.

        This step only saves intermediate files and does not alter the main pipeline.
        """
        print("-" * 50)
        print("🧭 步骤1.4: 按横竖段落拆分并保存临时文件")

        if not content_file or not os.path.exists(content_file):
            print("⚠️ 正文内容文件不存在，跳过步骤1.4")
            return content_file, []

        if not self.temp_dir:
            raise ValueError("临时目录未初始化")

        output_dir = os.path.join(self.temp_dir, "step1_4_orientation_segments")
        os.makedirs(output_dir, exist_ok=True)

        processor = DocumentOrientationProcessor()
        split_files = processor.split_document_by_orientation_segments(content_file, output_dir)

        try:
            self.intermediate_files["orientation_segments"] = json.dumps(
                split_files, ensure_ascii=False
            )
        except Exception:
            pass

        if self.save_intermediate_files and split_files:
            for i, item in enumerate(split_files, 1):
                try:
                    self._save_intermediate_file(
                        item["file_path"],
                        "step1_4",
                        f"segment_{i}_{item.get('orientation', 'unknown')}",
                    )
                except Exception:
                    continue

        return content_file, split_files

    def step2_pandoc_convert(
        self,
        content_file: str,
        template_file: str,
        include_toc: bool = True,
    ) -> Optional[str]:
        """
        步骤2: Pandoc转换 - 使用模板文件格式化正文内容

        Args:
            content_file: 正文内容文件路径
            template_file: 模板文件路径
            include_toc: 是否生成目录（默认为True）

        Returns:
            str: Pandoc处理后的文件路径
        """
        print("-" * 50)
        print("🔄 步骤2: Pandoc转换")

        try:
            # 初始化Pandoc转换器
            if self.pandoc_converter is None:
                # 查找pandoc可执行文件
                pandoc_path = self._find_pandoc_executable()
                if not pandoc_path:
                    print("⚠️ Pandoc不可用，跳过Pandoc转换步骤")

                    # 生成一个标记后的文件，表示跳过了Pandoc转换
                    base_name = os.path.splitext(os.path.basename(content_file))[0]
                    if not self.temp_dir:
                        raise ValueError("临时目录未初始化")
                    skipped_output = os.path.join(self.temp_dir, f"{base_name}_跳过Pandoc转换.docx")
                    shutil.copy2(content_file, skipped_output)

                    # 保存中间文件
                    self.intermediate_files['pandoc_converted'] = skipped_output

                    # 保存调试文件
                    if self.save_intermediate_files:
                        print(f"📁 正在保存step2中间文件到: {self.debug_output_dir}")
                        self._save_intermediate_file(skipped_output, "step2_pandoc", "跳过转换")

                    return skipped_output

                try:
                    self.pandoc_converter = PandocConverter(pandoc_path)
                except Exception as init_error:
                    print(f"⚠️ Pandoc初始化失败: {init_error}")
                    print("跳过Pandoc转换步骤")

                    # 生成一个标记后的文件，表示跳过了Pandoc转换
                    base_name = os.path.splitext(os.path.basename(content_file))[0]
                    if not self.temp_dir:
                        raise ValueError("临时目录未初始化")
                    init_failed_output = os.path.join(self.temp_dir, f"{base_name}_Pandoc初始化失败.docx")
                    shutil.copy2(content_file, init_failed_output)

                    # 保存中间文件
                    self.intermediate_files['pandoc_converted'] = init_failed_output

                    # 保存调试文件
                    if self.save_intermediate_files:
                        print(f"📁 正在保存step2中间文件到: {self.debug_output_dir}")
                        self._save_intermediate_file(init_failed_output, "step2_pandoc", "初始化失败")

                    return init_failed_output

            # 生成Pandoc输出文件路径
            base_name = os.path.splitext(os.path.basename(content_file))[0]
            if not self.temp_dir:
                raise ValueError("临时目录未初始化")
            pandoc_output = os.path.join(self.temp_dir, f"{base_name}_pandoc转换.docx")

            if self.save_intermediate_files:
                print(f"📄 正在使用模板转换: {os.path.basename(template_file)}")
                print(f"📤 输出文件: {os.path.basename(pandoc_output)}")

            # 使用模板进行转换，保持表格结构
            additional_args = [
                "--preserve-tabs",  # 保持制表符
                "--wrap=none",  # 不自动换行
                "--reference-links",  # 使用引用链接
                "--columns=80",  # 设置合适的列宽
                "--standalone",  # 独立文档模式
            ]
            if include_toc:
                additional_args.append("--table-of-contents")  # 仅在需要时生成目录

            success = self.pandoc_converter.convert_with_template(
                input_file=content_file,
                output_file=pandoc_output,
                template_file=template_file,
                additional_args=additional_args,
            )

            if success and os.path.exists(pandoc_output):
                print("✅ Pandoc转换成功!")
                # 保存中间文件
                self.intermediate_files['pandoc_converted'] = pandoc_output

                # 保存调试文件
                if self.save_intermediate_files:
                    print(f"   转换后文件: {os.path.basename(pandoc_output)}")
                    print(f"📁 正在保存step2中间文件到: {self.debug_output_dir}")
                    self._save_intermediate_file(pandoc_output, "step2_pandoc", "转换成功")

                return pandoc_output
            else:
                print("❌ Pandoc转换失败，使用原始文件继续")

                # 在转换失败时，复制原文件作为备用
                base_name = os.path.splitext(os.path.basename(content_file))[0]
                fallback_output = os.path.join(self.temp_dir, f"{base_name}_Pandoc失败备用.docx")
                shutil.copy2(content_file, fallback_output)

                # 保存中间文件
                self.intermediate_files['pandoc_converted'] = fallback_output

                # 保存调试文件
                if self.save_intermediate_files:
                    print(f"📁 正在保存step2中间文件到: {self.debug_output_dir}")
                    self._save_intermediate_file(fallback_output, "step2_pandoc", "失败备用")

                return fallback_output

        except Exception as e:
            print(f"❌ Pandoc转换过程中发生错误: {str(e)}")
            print("使用原始文件继续后续处理")

            # 在发生异常时，复制原文件作为备用
            base_name = os.path.splitext(os.path.basename(content_file))[0]
            if not self.temp_dir:
                raise ValueError("临时目录未初始化")
            error_output = os.path.join(self.temp_dir, f"{base_name}_Pandoc异常备用.docx")
            shutil.copy2(content_file, error_output)

            # 保存中间文件
            self.intermediate_files['pandoc_converted'] = error_output

            # 保存调试文件
            if self.save_intermediate_files:
                print(f"📁 正在保存step2中间文件到: {self.debug_output_dir}")
                self._save_intermediate_file(error_output, "step2_pandoc", "异常备用")

            return error_output
            success = self.pandoc_converter.convert_with_template(
                input_file=content_file,
                output_file=pandoc_output,
                template_file=template_file,
                additional_args=additional_args,
            )

            if success and os.path.exists(pandoc_output):
                print("✅ Pandoc转换成功!")
                # 保存中间文件
                self.intermediate_files['pandoc_converted'] = pandoc_output

                # 保存调试文件
                if self.save_intermediate_files:
                    print(f"   转换后文件: {os.path.basename(pandoc_output)}")
                    print(f"📁 正在保存step2中间文件到: {self.debug_output_dir}")
                    self._save_intermediate_file(pandoc_output, "step2_pandoc", "转换成功")

                return pandoc_output
            else:
                print("❌ Pandoc转换失败，使用原始文件继续")

                # 在转换失败时，复制原文件作为备用
                base_name = os.path.splitext(os.path.basename(content_file))[0]
                fallback_output = os.path.join(self.temp_dir, f"{base_name}_Pandoc失败备用.docx")
                shutil.copy2(content_file, fallback_output)

                # 保存中间文件
                self.intermediate_files['pandoc_converted'] = fallback_output

                # 保存调试文件
                if self.save_intermediate_files:
                    print(f"📁 正在保存step2中间文件到: {self.debug_output_dir}")
                    self._save_intermediate_file(fallback_output, "step2_pandoc", "失败备用")

                return fallback_output

        except Exception as e:
            print(f"❌ Pandoc转换过程中发生错误: {str(e)}")
            print("使用原始文件继续后续处理")

            # 在发生异常时，复制原文件作为备用
            base_name = os.path.splitext(os.path.basename(content_file))[0]
            if not self.temp_dir:
                raise ValueError("临时目录未初始化")
            error_output = os.path.join(self.temp_dir, f"{base_name}_Pandoc异常备用.docx")
            shutil.copy2(content_file, error_output)

            # 保存中间文件
            self.intermediate_files['pandoc_converted'] = error_output

            # 保存调试文件
            if self.save_intermediate_files:
                print(f"📁 正在保存step2中间文件到: {self.debug_output_dir}")
                self._save_intermediate_file(error_output, "step2_pandoc", "异常备用")

            return error_output

    def step3_format_tables(self, content_file: str, template_file: str, original_content_file: Optional[str] = None) -> \
    Optional[str]:
        """
        步骤3: 表格格式化 - 使用原始内容文件中的表格替换处理后的文件中的表格

        Args:
            content_file: 正文内容文件路径（被替换表格内容）
            template_file: 模板文件路径
            original_content_file: 原始正文内容文件路径（提供表格内容），如果提供则使用表格替换功能

        Returns:
            str: 表格格式化后的文件路径
        """
        print("-" * 50)
        print("📊 步骤3: 表格格式化")

        try:
            # 确保临时目录存在
            if not self.temp_dir:
                raise ValueError("临时目录未初始化")

            # 生成表格格式化输出文件路径
            base_name = os.path.splitext(os.path.basename(content_file))[0]

            # 如果提供了原始内容文件，则使用表格替换功能
            if original_content_file and os.path.exists(original_content_file):
                formatted_output = os.path.join(self.temp_dir, f"{base_name}_表格替换.docx")
                print("🔧 使用表格替换功能处理表格...")

                # 使用docx_table_replace.py中的方法进行表格替换
                try:
                    from utils.docx_table_replace import replace_tables_by_index

                    # 使用原始正文内容文件中的表格替换Pandoc转换后的文件中的表格
                    success = replace_tables_by_index(
                        original_path=original_content_file,  # 原始正文内容文件（提供表格）
                        edited_path=content_file,  # Pandoc转换后的文件（被替换表格）
                        output_path=formatted_output  # 输出文件
                    )

                    if success and os.path.exists(formatted_output):
                        print("✅ 表格替换成功!")

                        # 保存中间文件
                        self.intermediate_files['table_replaced'] = formatted_output

                        # 保存调试文件
                        if self.save_intermediate_files:
                            print(f"   表格替换后文件: {os.path.basename(formatted_output)}")
                            print(f"📁 正在保存step3中间文件到: {self.debug_output_dir}")
                            self._save_intermediate_file(formatted_output, "step3_table", "替换完成")

                        return formatted_output
                    else:
                        print("❌ 表格替换失败，使用原始文件继续")
                        return content_file

                except Exception as replace_error:
                    print(f"❌ 表格替换过程中发生错误: {str(replace_error)}")
                    print("使用原始文件继续后续处理")
                    return content_file
            else:
                print("⚠️ 未提供原始内容文件，跳过表格替换步骤")
                return content_file

        except Exception as e:
            print(f"❌ 表格格式化过程中发生错误: {str(e)}")
            return content_file

    def step4_merge_documents(self, cover_toc_file: str, processed_content_file: str, output_file: str) -> bool:
        """
        步骤4: 文档合并 - 将封面添加到正文开始

        Args:
            cover_toc_file: 封面文件路径
            processed_content_file: 处理后的正文文件路径
            output_file: 最终输出文件路径

        Returns:
            bool: 合并是否成功
        """
        print("-" * 50)
        print("📚 步骤4: 文档合并")

        try:
            if self.save_intermediate_files:
                print(f"📄 封面: {os.path.basename(cover_toc_file)}")
                print(f"📄 正文内容: {os.path.basename(processed_content_file)}")
                print(f"📤 最终输出: {os.path.basename(output_file)}")

            # 确保输出目录存在
            output_dir = os.path.dirname(output_file)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
                print(f"📁 创建输出目录: {output_dir}")

            # 使用docx_merge.py中的方法进行文档合并
            try:
                copy_all_to_beginning(
                    file_a=cover_toc_file,
                    file_b=processed_content_file,
                    output_file=output_file
                )

                if os.path.exists(output_file):
                    print("✅ 文档合并成功!")

                    # 验证输出文件的有效性
                    if self._validate_output_file(output_file):
                        print("✅ 输出文件验证通过")
                        # 保存最终文件
                        self.intermediate_files['final_document'] = output_file

                        # 保存调试文件
                        if self.save_intermediate_files:
                            print(f"   最终文档: {os.path.basename(output_file)}")
                            print(f"📁 正在保存step4最终文件到: {self.debug_output_dir}")
                            self._save_intermediate_file(output_file, "step4_final", "最终文档")
                        return True
                    else:
                        print("❌ 输出文件验证失败")
                        return False

                else:
                    print("❌ 文档合并失败，输出文件不存在")
                    return False

            except Exception as merge_error:
                print(f"❌ 使用docx_merge方法合并文档时发生错误: {str(merge_error)}")
                return False

        except Exception as e:
            print(f"❌ 文档合并过程中发生错误: {str(e)}")
            return False

    def step4_1_insert_toc_after_cover(self, output_file: str, toc_title: str = "目录") -> bool:
        """
        步骤4.1: 在合并后的文档中（封面后）插入完整目录

        Args:
            output_file: 合并后的文档路径
            toc_title: 目录标题

        Returns:
            bool: 是否成功
        """
        print("-" * 50)
        print("📑 步骤4.1: 插入完整目录 (COM)")

        if not COM_AVAILABLE:
            print("❌ COM组件不可用，无法插入目录")
            return False

        try:
            if self.save_intermediate_files:
                print(f"📄 目标文档: {os.path.basename(output_file)}")

            # 确保文件存在
            if not os.path.exists(output_file):
                print(f"❌ 文件不存在: {output_file}")
                return False

            import pythoncom
            import win32com.client as win32
            from win32com.client import constants as C
            import re

            pythoncom.CoInitialize()
            word = win32.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False

            doc = None
            try:
                doc = word.Documents.Open(output_file, ReadOnly=False)

                # 1. 清理已有的目录（如果有残留）
                if doc.TablesOfContents.Count > 0:
                    print(f"   发现 {doc.TablesOfContents.Count} 个旧目录，正在清理...")
                    for i in range(doc.TablesOfContents.Count, 0, -1):
                        doc.TablesOfContents(i).Delete()

                # 1.1 清理残留的目录标题/目录段落（避免出现两个“目录”）
                toc_titles = {
                    toc_title,
                    "目录",
                    "目      录",
                    "目 录",
                    "Table of Contents",
                    "Contents",
                    "TABLE OF CONTENTS",
                }

                toc_like_regex = re.compile(r"(\.{2,}|…+|\t)\s*\d+\s*$")

                def _should_delete_toc_para(para) -> bool:
                    try:
                        text = para.Range.Text.strip()
                    except Exception:
                        text = ""
                    if not text:
                        return False
                    style_name = ""
                    try:
                        style_name = str(para.Style.NameLocal)
                    except Exception:
                        try:
                            style_name = str(para.Style)
                        except Exception:
                            style_name = ""

                    # 1. 样式检查
                    if style_name.startswith("TOC") or "目录" in style_name:
                        return True
                    if "Hyperlink" in style_name or "超链接" in style_name:
                        return True

                    # 2. 标题文本检查
                    if text in toc_titles:
                        return True

                    # 3. 残留目录项特征检查 (如: "1. 概述 ...... 1")
                    if toc_like_regex.search(text):
                        return True
                    # 4. 超链接且疑似目录项
                    try:
                        if getattr(para.Range, "Hyperlinks", None) and para.Range.Hyperlinks.Count > 0:
                            if toc_like_regex.search(text) or text[:2].isdigit():
                                return True
                    except Exception:
                        pass
                    if (".." in text or "…" in text) and any(c.isdigit() for c in text[-5:]):
                        return True

                    return False

                try:
                    if doc.Sections.Count >= 2:
                        para_collection = doc.Sections(2).Range.Paragraphs
                    else:
                        para_collection = doc.Paragraphs
                    max_check = min(para_collection.Count, 160)
                    for i in range(max_check, 0, -1):
                        para = para_collection(i)
                        if _should_delete_toc_para(para):
                            para.Range.Delete()
                except Exception:
                    pass

                # 2. 确定插入位置
                insert_range = None

                # 2.0 新逻辑：直接跳过前3页，定位到第4页开头插入
                print("   尝试定位到第4页开头（跳过封面前3页）...")
                try:
                    # wdGoToPage=1, wdGoToAbsolute=1, Count=4
                    # 注意：如果文档不足4页，GoTo通常会停在文档末尾或报错
                    page4_range = doc.GoTo(What=1, Which=1, Count=4)

                    # 检查是否真的到了第4页（防止页数不足时的误判）
                    # 获取当前range所在的页码
                    # wdActiveEndPageNumber = 3
                    current_page = page4_range.Information(3)

                    if current_page >= 4:
                        insert_range = doc.Range(page4_range.Start, page4_range.Start)
                        print(f"   成功定位到第 {current_page} 页，准备插入目录")
                    else:
                        print(f"   文档只有 {current_page} 页，无法跳过前3页")
                        insert_range = None
                except Exception as e:
                    print(f"   定位第4页失败: {e}")
                    insert_range = None

                # 2.1 兜底逻辑：如果无法定位到第4页，则找第2节开始
                if insert_range is None:
                    if doc.Sections.Count >= 2:
                        insert_range = doc.Sections(2).Range
                        insert_range.Collapse(1) # wdCollapseStart = 1
                    else:
                        print("⚠️ 文档少于2节，尝试按页/标题定位目录插入点")
                        # 2.1 优先定位到第2页起始（更接近封面之后）
                        try:
                            page_count = doc.ComputeStatistics(2) # wdStatisticPages = 2
                        except Exception:
                            page_count = 0
                        if page_count and page_count >= 2:
                            try:
                                # wdGoToPage=1, wdGoToAbsolute=1
                                page2_range = doc.GoTo(What=1, Which=1, Count=2)
                                insert_range = doc.Range(page2_range.Start, page2_range.Start)
                            except Exception:
                                insert_range = None

                    # 2.2 如果仍未定位，尝试在第一个正文标题前插入
                    if insert_range is None:
                        for i in range(1, doc.Paragraphs.Count + 1):
                            para = doc.Paragraphs(i)
                            try:
                                text = para.Range.Text.strip()
                            except Exception:
                                text = ""
                            if not text:
                                continue
                            style_name = ""
                            try:
                                style_name = str(para.Style.NameLocal)
                            except Exception:
                                try:
                                    style_name = str(para.Style)
                                except Exception:
                                    style_name = ""
                            if style_name.startswith("Heading") or style_name.startswith("标题"):
                                insert_range = para.Range
                                insert_range.Collapse(1) # wdCollapseStart = 1
                                break

                    # 2.3 最终兜底：文档开头（仅在无法定位时）
                    if insert_range is None:
                        print("⚠️ 无法定位封面后位置，兜底在文档开头插入目录")
                        insert_range = doc.Range(0, 0)

                # 3. 插入目录标题
                title_start = insert_range.Start
                insert_range.InsertAfter(toc_title + "\r")
                try:
                    # 标题段落强制居中，避免默认样式导致左对齐
                    title_para = doc.Range(title_start, title_start + len(toc_title) + 1).Paragraphs(1)
                    title_para.Range.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_CENTER
                    para.Range.Font.Bold = True
                except Exception as align_err:
                    print(f"⚠️ 目录标题居中设置失败（step4.1）: {align_err}")

                # 尝试设置标题样式（如果有的话，通常TOC标题可以用 "TOC Heading" 或 "标题 1"）
                # 这里先不强求样式，让它保持默认或继承，后续步骤可能会处理
                # 也可以尝试设置为 "Normal" 并加粗居中，或者 "Heading 1"

                # 移动范围到标题之后，准备插入TOC域
                insert_range.Collapse(0) # wdCollapseEnd = 0

                # 4. 插入目录域
                # UseHeadingStyles=True: 使用标题样式
                # UpperHeadingLevel=1, LowerHeadingLevel=3: 显示1-3级标题
                toc = doc.TablesOfContents.Add(
                    Range=insert_range,
                    UseHeadingStyles=True,
                    UpperHeadingLevel=1,
                    LowerHeadingLevel=3
                )

                # 5. 更新域
                doc.Fields.Update()

                doc.Save()
                print("✅ 完整目录插入成功")

                # 保存中间文件
                if self.save_intermediate_files:
                    if not self.temp_dir:
                        # 如果没有临时目录，就不保存副本了，避免报错
                        pass
                    else:
                        base_name = os.path.splitext(os.path.basename(output_file))[0]
                        debug_path = os.path.join(self.temp_dir, f"{base_name}_step4_1_toc_inserted.docx")
                        shutil.copy2(output_file, debug_path)
                        print(f"📁 正在保存step4.1中间文件到: {self.debug_output_dir}")
                        self._save_intermediate_file(debug_path, "step4_1_toc", "目录插入完成")

                return True

            except Exception as e:
                print(f"❌ 插入目录失败: {e}")
                import traceback
                traceback.print_exc()
                return False
            finally:
                if doc:
                    try:
                        doc.Close(SaveChanges=0)
                    except Exception:
                        pass
                if word:
                    try:
                        word.Quit()
                    except:
                        pass
                pythoncom.CoUninitialize()

        except Exception as e:
            print(f"❌ 步骤4.1发生未捕获异常: {e}")
            return False

    def step5_update_toc_title(self, docx_file: str, new_title: str = "目      录") -> bool:
        """
        步骤5: 更新目录标题 - 修改文档中的目录标题

        Args:
            docx_file: 需要修改的文档路径
            new_title: 新的目录标题

        Returns:
            bool: 更新是否成功
        """
        print("-" * 50)
        print("🏷️ 步骤5: 更新目录标题")

        try:
            if self.save_intermediate_files:
                print(f"📄 目标文档: {os.path.basename(docx_file)}")
                print(f"🔤 新标题: '{new_title}'")

            # 检查文件是否存在
            if not os.path.exists(docx_file):
                print(f"❌ 文件不存在: {docx_file}")
                return False

            updated = False
            # 使用docx_update_toc_title中的方法更新目录标题
            try:
                # 使用XML方式更新目录标题，保留原有格式
                update_toc_title_xml(docx_file, new_title)
                print("✅ 目录标题更新成功!")
                updated = True
            except Exception as xml_error:
                print(f"⚠️ XML方式更新目录标题失败: {str(xml_error)}")
                print("尝试使用COM方式更新目录标题...")

                try:
                    # 使用COM方式更新目录标题
                    from utils.docx_update_toc_title import update_toc_title
                    update_toc_title(docx_file, new_title)
                    print("✅ 目录标题更新成功!")
                    updated = True
                except Exception as com_error:
                    print(f"❌ COM方式更新目录标题也失败: {str(com_error)}")
                    updated = False

            if updated:
                self._center_toc_title_with_com(docx_file, new_title)
                return True
            return False

        except Exception as e:
            print(f"❌ 更新目录标题过程中发生错误: {str(e)}")
            return False

    def step6_format_pictures(self, docx_file: str, formula_info_list: list | None = None) -> bool:
        """
        步骤6: 图片格式化 - 图片居中，单倍行距

        Args:
            docx_file: 需要处理的文档路径
            formula_info_list: 公式图片信息

        Returns:
            bool: 处理是否成功
        """
        print("-" * 50)
        print("🖼️ 步骤6: 图片格式化")

        try:
            if self.save_intermediate_files:
                print(f"📄 目标文档: {os.path.basename(docx_file)}")

            # 检查文件是否存在
            if not os.path.exists(docx_file):
                print(f"❌ 文件不存在: {docx_file}")
                return False

            # 生成处理后的文件路径
            base_name = os.path.splitext(os.path.basename(docx_file))[0]
            if not self.temp_dir:
                raise ValueError("临时目录未初始化")
            formatted_output = os.path.join(self.temp_dir, f"{base_name}_图片格式化.docx")

            # 使用docx_picture.py中的高级处理方式
            try:
                from utils.docx_picture import format_pictures_with_advanced_settings

                # 调用高级图片格式化函数
                success = format_pictures_with_advanced_settings(
                    doc_path=docx_file,
                    save_path=formatted_output,
                    formula_info_list=formula_info_list,
                    skip_cover=True
                )

                if success and os.path.exists(formatted_output):
                    print("✅ 图片格式化成功!")

                    # 保存中间文件
                    self.intermediate_files['picture_formatted'] = formatted_output

                    # 保存调试文件
                    if self.save_intermediate_files:
                        print(f"   图片格式化后文件: {os.path.basename(formatted_output)}")
                        print(f"📁 正在保存step6中间文件到: {self.debug_output_dir}")
                        self._save_intermediate_file(formatted_output, "step6_picture", "格式化完成")

                    # 将处理后的文件复制回原文件路径，以便后续步骤使用
                    shutil.copy2(formatted_output, docx_file)
                    return True
                else:
                    print("❌ 图片格式化失败")
                    return False

            except Exception as format_error:
                print(f"❌ 图片格式化过程中发生错误: {str(format_error)}")
                import traceback
                traceback.print_exc()
                return False

        except Exception as e:
            print(f"❌ 图片格式化过程中发生错误: {str(e)}")
            return False

    def step7_format_library_number(self, docx_file: str) -> bool:
        """
        步骤7: 库号信息格式化 - 将库号右靠齐

        Args:
            docx_file: 需要处理的文档路径

        Returns:
            bool: 处理是否成功
        """
        print("-" * 50)
        print("🔢 步骤7: 库号信息格式化")

        try:
            if self.save_intermediate_files:
                print(f"📄 目标文档: {os.path.basename(docx_file)}")

            # 检查文件是否存在
            if not os.path.exists(docx_file):
                print(f"❌ 文件不存在: {docx_file}")
                return False

            # 生成处理后的文件路径
            base_name = os.path.splitext(os.path.basename(docx_file))[0]
            if not self.temp_dir:
                raise ValueError("临时目录未初始化")
            formatted_output = os.path.join(self.temp_dir, f"{base_name}_库号格式化.docx")

            # 使用docx_supplement.py中的高级处理方式
            try:
                from utils.docx_supplement import format_library_number_advanced

                # 调用高级库号格式化函数
                success = format_library_number_advanced(
                    doc_path=docx_file,
                    save_path=formatted_output
                )

                if success and os.path.exists(formatted_output):
                    print("✅ 库号信息格式化成功!")

                    # 保存中间文件
                    self.intermediate_files['library_number_formatted'] = formatted_output

                    # 保存调试文件
                    if self.save_intermediate_files:
                        print(f"   库号格式化后文件: {os.path.basename(formatted_output)}")
                        print(f"📁 正在保存step7中间文件到: {self.debug_output_dir}")
                        self._save_intermediate_file(formatted_output, "step7_library_number", "格式化完成")

                    # 将处理后的文件复制回原文件路径，以便后续步骤使用
                    shutil.copy2(formatted_output, docx_file)
                    return True
                else:
                    print("❌ 库号信息格式化失败")
                    return False

            except Exception as format_error:
                print(f"❌ 库号信息格式化过程中发生错误: {str(format_error)}")
                import traceback
                traceback.print_exc()
                return False

        except Exception as e:
            print(f"❌ 库号信息格式化过程中发生错误: {str(e)}")
            return False

    def step8_insert_section_break(self, docx_file: str) -> bool:
        """
        步骤8: 在目录后插入分节符

        Args:
            docx_file: 需要处理的文档路径

        Returns:
            bool: 处理是否成功
        """
        print("-" * 50)
        print("📑 步骤8: 在目录后插入分节符")

        try:
            if self.save_intermediate_files:
                print(f"📄 目标文档: {os.path.basename(docx_file)}")

            # 检查文件是否存在
            if not os.path.exists(docx_file):
                print(f"❌ 文件不存在: {docx_file}")
                return False

            # 添加延迟，确保之前的COM操作完全释放资源
            print("⏳ 等待COM资源释放...")
            import time
            time.sleep(3)

            # 生成处理后的文件路径
            base_name = os.path.splitext(os.path.basename(docx_file))[0]
            if not self.temp_dir:
                raise ValueError("临时目录未初始化")
            formatted_output = os.path.join(self.temp_dir, f"{base_name}_插入分节符.docx")

            # 使用docx_supplement.py中的方法
            try:
                from utils.docx_supplement import insert_section_break_after_toc

                # 调用插入分节符函数
                success = insert_section_break_after_toc(
                    doc_path=docx_file,
                    save_path=formatted_output
                )

                if success and os.path.exists(formatted_output):
                    print("✅ 在目录后插入分节符成功!")

                    # 保存中间文件
                    self.intermediate_files['section_break_inserted'] = formatted_output

                    # 保存调试文件
                    if self.save_intermediate_files:
                        print(f"   插入分节符后文件: {os.path.basename(formatted_output)}")
                        print(f"📁 正在保存step8中间文件到: {self.debug_output_dir}")
                        self._save_intermediate_file(formatted_output, "step8_section_break", "插入完成")

                    # 将处理后的文件复制回原文件路径，以便后续步骤使用
                    shutil.copy2(formatted_output, docx_file)
                    return True
                else:
                    print("❌ 在目录后插入分节符失败")
                    # 尝试使用备选方法
                    print("🔄 尝试使用备选方法...")
                    return self._insert_section_break_fallback(docx_file, formatted_output)

            except Exception as format_error:
                print(f"❌ 在目录后插入分节符过程中发生错误: {str(format_error)}")
                # 尝试使用备选方法
                print("🔄 尝试使用备选方法...")
                return self._insert_section_break_fallback(docx_file, formatted_output)

        except Exception as e:
            print(f"❌ 在目录后插入分节符过程中发生错误: {str(e)}")
            return False

    def step9_process_sections(self, docx_file: str) -> bool:
        """
        步骤9: 处理文档节的页码设置
        - 取消第三节与第二节的链接
        - 处理第二节的页码（删除PAGE域）
        - 处理第三节的页码（重置为1）

        Args:
            docx_file: 需要处理的文档路径

        Returns:
            bool: 处理是否成功
        """
        print("-" * 50)
        print("📑 步骤9: 处理文档节的页码设置")

        try:
            if self.save_intermediate_files:
                print(f"📄 目标文档: {os.path.basename(docx_file)}")

            # 检查文件是否存在
            if not os.path.exists(docx_file):
                print(f"❌ 文件不存在: {docx_file}")
                return False

            # 添加延迟，确保之前的COM操作完全释放资源
            print("⏳ 等待COM资源释放...")
            import time
            time.sleep(3)

            # 生成临时文件路径
            base_name = os.path.splitext(os.path.basename(docx_file))[0]
            if not self.temp_dir:
                raise ValueError("临时目录未初始化")

            temp_file1 = os.path.join(self.temp_dir, f"{base_name}_取消节链接.docx")
            temp_file2 = os.path.join(self.temp_dir, f"{base_name}_处理第二节页码.docx")
            formatted_output = os.path.join(self.temp_dir, f"{base_name}_处理第三节页码.docx")

            # 步骤9.1: 使用docx_supplement.py中的方法取消第三节与第二节的链接
            print("\n步骤9.1: 取消第三节与第二节的链接...")
            try:
                from utils.docx_supplement import cancel_section_link_com

                success = cancel_section_link_com(
                    doc_path=docx_file,
                    save_path=temp_file1,
                    section_number=3  # 第三节
                )

                if not success or not os.path.exists(temp_file1):
                    print("❌ 步骤9.1失败，无法继续执行后续步骤")
                    return False
                else:
                    print("✅ 步骤9.1完成")

                    # 保存调试文件
                    if self.save_intermediate_files:
                        print(f"   取消节链接后文件: {os.path.basename(temp_file1)}")
                        print(f"📁 正在保存step9.1中间文件到: {self.debug_output_dir}")
                        self._save_intermediate_file(temp_file1, "step9_section_link", "取消链接完成")
            except Exception as e:
                print(f"❌ 步骤9.1失败: {e}")
                return False

            # 步骤9.2: 使用docx_supplement.py中的方法处理第二节的页码
            print("\n步骤9.2: 处理第二节的页码...")
            try:
                from utils.docx_supplement import process_section2_docx

                process_section2_docx(temp_file1, temp_file2, section_index=2)
                print("✅ 步骤9.2完成")

                # 保存调试文件
                if self.save_intermediate_files:
                    print(f"   处理第二节页码后文件: {os.path.basename(temp_file2)}")
                    print(f"📁 正在保存step9.2中间文件到: {self.debug_output_dir}")
                    self._save_intermediate_file(temp_file2, "step9_section2_page", "处理完成")
            except Exception as e:
                print(f"❌ 步骤9.2失败: {e}")
                return False

            # 步骤9.3: 使用docx_supplement.py中的方法处理第三节的页码
            print("\n步骤9.3: 处理第三节的页码...")
            try:
                from utils.docx_supplement import process_section3_docx

                process_section3_docx(temp_file2, formatted_output)
                print("✅ 步骤9.3完成")

                # 保存中间文件
                self.intermediate_files['section_page_processed'] = formatted_output

                # 保存调试文件
                if self.save_intermediate_files:
                    print(f"   处理第三节页码后文件: {os.path.basename(formatted_output)}")
                    print(f"📁 正在保存step9.3中间文件到: {self.debug_output_dir}")
                    self._save_intermediate_file(formatted_output, "step9_section3_page", "处理完成")

                # 将处理后的文件复制回原文件路径，以便后续步骤使用
                shutil.copy2(formatted_output, docx_file)
                return True
            except Exception as e:
                print(f"❌ 步骤9.3失败: {e}")
                return False

        except Exception as e:
            print(f"❌ 处理文档节的页码设置过程中发生错误: {str(e)}")
            return False

    def step11_format_library_number_advanced(self, docx_file: str) -> bool:
        """
        步骤11: 高级库号信息格式化 - 将"库号："设置为三号黑体加粗，数字字母设置为三号Times New Roman加粗

        Args:
            docx_file: 需要处理的文档路径

        Returns:
            bool: 处理是否成功
        """
        print("-" * 50)
        print("🔢 步骤11: 高级库号信息格式化")

        try:
            if self.save_intermediate_files:
                print(f"📄 目标文档: {os.path.basename(docx_file)}")

            # 检查文件是否存在
            if not os.path.exists(docx_file):
                print(f"❌ 文件不存在: {docx_file}")
                return False

            # 生成处理后的文件路径
            base_name = os.path.splitext(os.path.basename(docx_file))[0]
            if not self.temp_dir:
                raise ValueError("临时目录未初始化")
            formatted_output = os.path.join(self.temp_dir, f"{base_name}_库号高级格式化.docx")

            # 使用cover_number_format.py中的处理方式
            try:
                from utils.cover_number_format import process_document

                # 调用库号格式化函数
                process_document(docx_file, formatted_output)

                if os.path.exists(formatted_output):
                    print("✅ 库号信息高级格式化成功!")

                    # 保存中间文件
                    self.intermediate_files['library_number_advanced_formatted'] = formatted_output

                    # 保存调试文件
                    if self.save_intermediate_files:
                        print(f"   库号高级格式化后文件: {os.path.basename(formatted_output)}")
                        print(f"📁 正在保存step11中间文件到: {self.debug_output_dir}")
                        self._save_intermediate_file(formatted_output, "step11_library_number", "高级格式化完成")

                    # 将处理后的文件复制回原文件路径，以便后续步骤使用
                    shutil.copy2(formatted_output, docx_file)
                    return True
                else:
                    print("❌ 库号信息高级格式化失败")
                    return False

            except Exception as format_error:
                print(f"❌ 库号信息高级格式化过程中发生错误: {str(format_error)}")
                import traceback
                traceback.print_exc()
                return False

        except Exception as e:
            print(f"❌ 库号信息高级格式化过程中发生错误: {str(e)}")
            return False

    def step11_5_insert_by_map(self, docx_file: str) -> bool:
        """
        步骤11.5: 基于 step1.9 的 image_map.json 将提取的图片回插到文档中。

        - 输入：步骤11后的 docx_file
        - 使用：self.intermediate_files['step1_9_run_dir']（包含 image_map.json 和图片）
        - 输出：生成到 self.temp_dir，并在成功时覆盖回 docx_file 供步骤12继续使用
        - 调试暂存：仅在 save_intermediate_files=True 时同步保存到 self.debug_output_dir
        """
        print("-" * 50)
        print("🖼️ 步骤11.5: 基于映射回插图片（insert_by_map）")

        if os.name != "nt":
            print("⚠️ 当前系统非Windows，跳过 step11.5")
            return False

        if not docx_file or not os.path.exists(docx_file):
            print(f"⚠️ step11.5: 目标文档不存在，跳过: {docx_file}")
            return False

        run_dir = self.intermediate_files.get("step1_9_run_dir")
        if not run_dir or not os.path.exists(run_dir):
            print("⚠️ step11.5: 未找到 step1.9 输出目录（step1_9_run_dir），跳过")
            return False

        map_json = os.path.join(run_dir, "image_map.json")
        if not os.path.exists(map_json):
            print(f"⚠️ step11.5: 找不到 image_map.json，跳过: {map_json}")
            return False

        try:
            from utils.insert_by_map import insert_images_fix_layout
        except Exception as import_err:
            print(f"⚠️ step11.5: insert_by_map 不可用，跳过。原因: {import_err}")
            return False

        try:
            if not self.temp_dir:
                raise ValueError("临时目录未初始化")

            base_name = os.path.splitext(os.path.basename(docx_file))[0]
            output_path = os.path.join(self.temp_dir, f"{base_name}_step11_5_insert_by_map.docx")

            result_path = insert_images_fix_layout(
                target_doc_path=docx_file,
                image_dir=run_dir,
                output_doc_path=output_path,
                word_visible=False,
            )

            if not result_path or not os.path.exists(result_path):
                print("⚠️ step11.5: 插图回填失败，继续使用原文档")
                return False

            self.intermediate_files["step11_5_insert_by_map"] = result_path

            if self.save_intermediate_files:
                print(f"📁 正在保存step11.5中间文件到: {self.debug_output_dir}")
                self._save_intermediate_file(result_path, "step11_5_insert_by_map", "插图回填")

            shutil.copy2(result_path, docx_file)
            print("✅ step11.5: 插图回填完成")
            return True

        except Exception as err:
            print(f"⚠️ step11.5: 运行失败，继续使用原文档。详细信息：{err}")
            return False

    def step12_post_process_headers_footers(self, docx_path: str) -> bool:
        """
        步骤12: 后处理页眉页脚 - 在所有转换步骤完成后处理页眉页脚

        Args:
            docx_path: 文档路径

        Returns:
            bool: 处理是否成功
        """
        print("-" * 50)
        print("📑 步骤13: 后处理页眉页脚")

        try:
            if self.save_intermediate_files:
                print(f"📄 目标文档: {os.path.basename(docx_path)}")

            # 检查文件是否存在
            if not os.path.exists(docx_path):
                print(f"❌ 文件不存在: {docx_path}")
                return False

            # 生成处理后的文件路径
            base_name = os.path.splitext(os.path.basename(docx_path))[0]
            if not self.temp_dir:
                raise ValueError("临时目录未初始化")
            formatted_output = os.path.join(self.temp_dir, f"{base_name}_step13_headers_footers.docx")

            # 确定页脚文本内容
            if self.extracted_footer_content and self.extracted_footer_content.strip():
                footer_text = self.extracted_footer_content.strip()
                print(f"📄 使用提取的页脚内容: {footer_text}")
            else:
                footer_text = "PAGE \\* MERGEFORMAT / NUMPAGES \\* MERGEFORMAT"
                print(f"📄 使用默认页脚内容: {footer_text}")

            from docx import Document
            from docx.shared import Pt, Cm

            # 执行页眉页脚处理逻辑
            doc = Document(docx_path)

            # 遍历文档的所有sections
            for i, section in enumerate(doc.sections):
                # 获取当前节的页眉
                header = section.header

                # 检查页眉是否包含文本
                header_contains_text = any(paragraph.text.strip() for paragraph in header.paragraphs)

                if header_contains_text:
                    # 如果当前节的页眉包含文本，则为该节及后续节添加页脚
                    footer = section.footer

                    # 遍历页脚段落并查找是否已经有页码
                    existing_page_numbers = []
                    for paragraph in footer.paragraphs:
                        for run in paragraph.runs:
                            if "PAGE" in run.text or "NUMPAGES" in run.text:
                                existing_page_numbers.append(run.text)

                    # 删除多余的制表符等符号
                    if footer.paragraphs:
                        para = footer.paragraphs[0]
                        tabs_deleted = 0  # 已删除的制表符数量
                        runs = para.runs  # 段落的文本片段列表

                        # 遍历所有run，从第一个开始处理
                        for run in runs:
                            run_text = run.text  # 当前run的文本

                            if tabs_deleted >= 3:
                                break  # 已删除三个，停止处理

                            # 查找当前run中开头的制表符
                            if not run_text:
                                continue  # 空run跳过

                            # 统计当前run开头的制表符数量
                            start_tabs = 0
                            while start_tabs < len(run_text) and (run_text[start_tabs] == '\t'):
                                start_tabs += 1

                            if start_tabs == 0:
                                continue  # 该run开头没有制表符，跳过

                            # 计算需要从当前run删除的制表符数量
                            need_delete = min(start_tabs, 3 - tabs_deleted)

                            # 删除开头的need_delete个制表符
                            run.text = run_text[need_delete:]  # 截断文本，保留剩余内容
                            tabs_deleted += need_delete

                        # 在footer.paragraphs[0]中添加footer_text
                        para.add_run(footer_text)

                        # 设置字号五号字
                        for run in para.runs:
                            run.font.size = Pt(10.5)

            # 保存修改后的文档
            doc.save(formatted_output)
            print("✅ 页眉页脚处理成功!")

            # 保存中间文件
            self.intermediate_files['headers_footers_processed'] = formatted_output

            # 保存调试文件
            if self.save_intermediate_files:
                print(f"   页眉页脚处理后文件: {os.path.basename(formatted_output)}")
                print(f"📁 正在保存step12表格中间文件到: {self.debug_output_dir}")
                self._save_intermediate_file(formatted_output, "step13_headers_footers", "处理完成")

            # 将处理后的文件复制回原文件路径，以便后续步骤使用
            shutil.copy2(formatted_output, docx_path)
            return True

        except Exception as e:
            print(f"❌ 页眉页脚处理过程中发生错误: {str(e)}")
            return False

    def step13_word_table_replace(self, edited_docx: str, original_docx: str) -> bool:
        """Step12: Table replacement using the Word COM helper."""
        try:
            from utils.word_table_replace import replace_tables
        except ImportError as import_error:
            print(f"[warn] Unable to import word_table_replace: {import_error}")
            return False

        print("-" * 50)
        print("步骤12: 表格替换")

        if not original_docx:
            print("[warn] Missing step0 output file, skip table replacement")
            return False

        try:
            if not os.path.exists(edited_docx):
                print(f"[warn] Edited document not found: {edited_docx}")
                return False
            if not os.path.exists(original_docx):
                print(f"[warn] Original document not found: {original_docx}")
                return False

            if not self.temp_dir:
                raise ValueError("临时目录未初始化")

            cached_sections = None
            try:
                from docx import Document
                pre_doc = Document(edited_docx)
                cached_sections = []
                for section in pre_doc.sections:
                    section_snapshot = {}
                    ps = getattr(section, "page_setup", None)
                    if ps is not None:
                        section_snapshot = {
                            "footer_distance": getattr(ps, "footer_distance", None),
                            "header_distance": getattr(ps, "header_distance", None),
                            "top_margin": getattr(ps, "top_margin", None),
                            "bottom_margin": getattr(ps, "bottom_margin", None),
                            "left_margin": getattr(ps, "left_margin", None),
                            "right_margin": getattr(ps, "right_margin", None),
                            "gutter": getattr(ps, "gutter", None),
                        }
                    else:
                        section_snapshot = {
                            "footer_distance": getattr(section, "footer_distance", None),
                            "header_distance": getattr(section, "header_distance", None),
                            "top_margin": getattr(section, "top_margin", None),
                            "bottom_margin": getattr(section, "bottom_margin", None),
                            "left_margin": getattr(section, "left_margin", None),
                            "right_margin": getattr(section, "right_margin", None),
                            "gutter": getattr(section, "gutter", None),
                        }
                    cached_sections.append(section_snapshot)
            except Exception as snapshot_err:
                print(f"[warn] Unable to snapshot section settings: {snapshot_err}")
                cached_sections = None

            base_name = os.path.splitext(os.path.basename(edited_docx))[0]
            replaced_output = os.path.join(self.temp_dir, f"{base_name}_step12_table_replace.docx")

            if self.save_intermediate_files:
                print(f"源文件(步骤0): {os.path.basename(original_docx)}")
                print(f"当前待替换文件: {os.path.basename(edited_docx)}")
                print(f"表格替换输出: {os.path.basename(replaced_output)}")

            _kill_existing_winword_processes()
            replace_tables(src_path=original_docx, dst_path=edited_docx, out_path=replaced_output)

            if os.path.exists(replaced_output):
                if cached_sections is not None:
                    try:
                        from docx import Document
                        _wait_for_file_release_for_step(replaced_output)
                        post_doc = Document(replaced_output)
                        for section, saved in zip(post_doc.sections, cached_sections):
                            ps = getattr(section, "page_setup", None)
                            if ps is not None:
                                if saved.get("footer_distance") is not None:
                                    ps.footer_distance = saved["footer_distance"]
                                if saved.get("header_distance") is not None:
                                    ps.header_distance = saved["header_distance"]
                                if saved.get("top_margin") is not None:
                                    ps.top_margin = saved["top_margin"]
                                if saved.get("bottom_margin") is not None:
                                    ps.bottom_margin = saved["bottom_margin"]
                                if saved.get("left_margin") is not None:
                                    ps.left_margin = saved["left_margin"]
                                if saved.get("right_margin") is not None:
                                    ps.right_margin = saved["right_margin"]
                                if saved.get("gutter") is not None:
                                    ps.gutter = saved["gutter"]
                            else:
                                for key, value in saved.items():
                                    if value is None:
                                        continue
                                    if hasattr(section, key):
                                        setattr(section, key, value)
                        post_doc.save(replaced_output)
                    except Exception as restore_err:
                        print(f"[warn] Unable to restore section settings: {restore_err}")

                _wait_for_file_release_for_step(edited_docx)
                _wait_for_file_release_for_step(replaced_output)
                shutil.copy2(replaced_output, edited_docx)
                self.intermediate_files['table_replaced_before_headers'] = replaced_output

                if self.save_intermediate_files:
                    print(f"   表格替换后文件: {os.path.basename(replaced_output)}")
                    print(f"📁 正在保存step12表格中间文件到: {self.debug_output_dir}")
                    self._save_intermediate_file(replaced_output, "step12_table", "表格替换完成")

                print("表格替换完成")
                return True

            print("表格替换失败，未生成输出文件")
            return False

        except Exception as exc:
            print(f"表格替换过程中发生错误: {exc}")
            import traceback
            traceback.print_exc()
            return False




    def _insert_section_break_fallback(self, docx_file: str, formatted_output: str) -> bool:
        """
        备选方法：在目录后插入分节符的降级处理

        Args:
            docx_file: 需要处理的文档路径
            formatted_output: 输出文件路径

        Returns:
            bool: 处理是否成功
        """
        try:
            print("🔧 尝试使用XML方法插入分节符...")

            # 使用XML方法作为备选
            from utils.docx_section_break import insert_section_break_after_toc_xml

            success = insert_section_break_after_toc_xml(
                doc_path=docx_file,
                save_path=formatted_output
            )

            if success and os.path.exists(formatted_output):
                print("✅ 使用XML方法插入分节符成功!")
                return True
            else:
                print("❌ 使用XML方法插入分节符失败")
                return False

        except Exception as xml_error:
            print(f"❌ 使用XML方法插入分节符时发生错误: {str(xml_error)}")
            print("⚠️ 无法在目录后插入分节符，继续使用原有格式")
            return False

    def _validate_output_file(self, file_path: str) -> bool:
        """
        验证输出文件的有效性

        Args:
            file_path: 文件路径

        Returns:
            bool: 文件是否有效
        """
        try:
            if not os.path.exists(file_path):
                return False

            # 检查文件大小
            file_size = os.path.getsize(file_path)
            if file_size < 1000:  # 小于1KB可能有问题
                print(f"⚠️ 文件大小异常: {file_size} bytes")
                return False

            # 检查是否为有效的ZIP文件（DOCX本质上ZIP文件）
            import zipfile
            try:
                with zipfile.ZipFile(file_path, 'r') as zip_file:
                    # 检查必要的文件
                    required_files = ['[Content_Types].xml', '_rels/.rels', 'word/document.xml']
                    file_list = zip_file.namelist()

                    for req_file in required_files:
                        if req_file not in file_list:
                            print(f"⚠️ 缺少必要文件: {req_file}")
                            return False

                    return True
            except zipfile.BadZipFile:
                print("⚠️ 文件不是有效的ZIP格式")
                return False

        except Exception as e:
            print(f"⚠️ 文件验证过程中发生错误: {str(e)}")
            return False

    def step10_remove_highlights(self, source_file: str) -> str:
        """
        步骤10: 删除文档中所有的突出显示（高亮、底纹、颜色）

        Args:
            source_file: 源文档路径

        Returns:
            str: 去除突出显示后的文档路径
        """
        print("-" * 50)
        print("📑 步骤10: 删除文档中所有的突出显示")

        try:
            # 确保临时目录存在
            if not self.temp_dir:
                raise ValueError("临时目录未初始化")

            # 检查lxml是否可用
            if not LXML_AVAILABLE:
                print("⚠️ lxml库不可用，跳过突出显示删除步骤")
                return source_file

            # 生成输出文件路径
            base_name = os.path.splitext(os.path.basename(source_file))[0]
            no_highlight_file = os.path.join(self.temp_dir, f"{base_name}_无突出显示.docx")

            print(f"📄 输入文档: {os.path.basename(source_file)}")
            print(f"📤 输出文档: {os.path.basename(no_highlight_file)}")

            # 使用lxml删除Word文件中的所有高亮
            try:
                # 定义命名空间和常量
                W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                NSMAP = {"w": W_NS}
                XML_PARSER = etree.XMLParser(ns_clean=True, recover=True, remove_blank_text=False) if (
                            LXML_AVAILABLE and etree is not None) else None
                REMOVE_COLOR_NODE = True  # 彻底移除颜色节点

                def process_xml_bytes(data: bytes, remove_color_node=False) -> bytes:
                    """删除 w:highlight、w:shd，并处理 w:color"""
                    if not LXML_AVAILABLE or etree is None:
                        return data

                    try:
                        root = etree.fromstring(data, parser=XML_PARSER) if etree is not None else None
                        if root is None:
                            return data
                    except Exception:
                        return data

                    changed = False
                    # 删除 highlight
                    for node in list(root.iter(qn('w:highlight'))):
                        parent = node.getparent()
                        if parent is not None:
                            parent.remove(node)
                            changed = True

                    # 删除底纹
                    for node in list(root.iter(qn('w:shd'))):
                        parent = node.getparent()
                        if parent is not None:
                            parent.remove(node)
                            changed = True

                    # 处理颜色
                    for color in list(root.iter(qn('w:color'))):
                        val = color.get("val")
                        if val is not None and val.lower() != "auto":
                            if remove_color_node:
                                parent = color.getparent()
                                if parent is not None:
                                    parent.remove(color)
                                    changed = True
                            else:
                                color.set("val", "auto")
                                changed = True

                    if not changed:
                        return data
                    return etree.tostring(root, encoding="utf-8", xml_declaration=True) if (
                                LXML_AVAILABLE and etree is not None) else data

                # 处理DOCX文件
                src = Path(source_file)
                if not src.exists():
                    raise FileNotFoundError(f"文件不存在: {src}")

                dest = Path(no_highlight_file)

                with zipfile.ZipFile(src, 'r') as zin:
                    with zipfile.ZipFile(dest, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
                        for name in zin.namelist():
                            data = zin.read(name)
                            if name.startswith("word/") and name.endswith(".xml"):
                                try:
                                    new_data = process_xml_bytes(data, remove_color_node=REMOVE_COLOR_NODE)
                                    zout.writestr(name, new_data)
                                except Exception as e:
                                    print(f"⚠ 处理 {name} 出错，保留原文件。错误：{e}")
                                    zout.writestr(name, data)
                            else:
                                zout.writestr(name, data)

                print("✅ 突出显示删除成功!")

                # 保存中间文件到指定目录便于查看调试
                if self.save_intermediate_files:
                    print(f"   无突出显示文档: {os.path.basename(no_highlight_file)}")
                    print(f"📁 正在保存step10中间文件到: {self.debug_output_dir}")
                    self._save_intermediate_file(no_highlight_file, "step10_highlights", "无突出显示")

                # 保存中间文件路径
                self.intermediate_files['no_highlights'] = no_highlight_file

                return no_highlight_file

            except Exception as remove_error:
                print(f"❌ 突出显示删除过程中发生错误: {str(remove_error)}")
                print("   继续使用原始文档")
                return source_file

        except Exception as e:
            print(f"❌ 突出显示删除过程中发生错误: {str(e)}")
            return source_file

    def _find_pandoc_executable(self) -> Optional[str]:
        """
        查找Pandoc可执行文件

        Returns:
            str: Pandoc可执行文件路径
        """
        import subprocess

        # 可能的Pandoc位置 - 优先utils目录
        current_dir = os.path.dirname(os.path.abspath(__file__))
        parent_dir = os.path.dirname(current_dir)
        utils_pandoc = os.path.join(parent_dir, 'utils', 'pandoc.exe')

        possible_paths = [
            # 优先使用utils目录中的pandoc.exe
            utils_pandoc,
            # 系统PATH中的pandoc
            "pandoc",
            "pandoc.exe",
            # 当前目录下的pandoc.exe
            os.path.join(os.path.dirname(__file__), "pandoc.exe"),
            os.path.join(os.path.dirname(os.path.dirname(__file__)), "pandoc.exe"),
            # 常见安装位置
            r"C:\Program Files\Pandoc\pandoc.exe",
            r"C:\Program Files (x86)\Pandoc\pandoc.exe",
            # Conda环境
            os.path.join(os.environ.get('CONDA_PREFIX', ''), 'Scripts', 'pandoc.exe'),
            os.path.join(os.environ.get('CONDA_PREFIX', ''), 'bin', 'pandoc')
        ]

        for path in possible_paths:
            if not path:  # 跳过空路径
                continue

            try:
                # 尝试执行pandoc --version
                result = subprocess.run(
                    [path, "--version"],
                    capture_output=True,
                    text=True,
                    timeout=10,
                    creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
                )
                if result.returncode == 0:
                    return path
            except (subprocess.TimeoutExpired, FileNotFoundError, PermissionError, Exception):
                continue

        print("⚠️ 未找到Pandoc可执行文件")
        print("请通过以下方式安装Pandoc:")
        print("1. 下载安装: https://pandoc.org/installing.html")
        print("2. 使用conda: conda install pandoc")
        print("3. 使用choco: choco install pandoc")
        print(f"4. utils目录路径: {utils_pandoc}")
        return None

    def step14_format_landscape_headers_footers(self, output_file: str):
        """
        步骤14：格式化横板页眉页脚
        1️⃣ 找到倒数第一、第二小节 sectPr
        2️⃣ 获取 headerReference 和 footerReference 的 r:id
        3️⃣ 将倒数第二节页眉页脚内容替换为倒数第一节的内容
        4️⃣ 输出日志
        5️⃣ 可修改页眉 top 距离和页脚页码居中
        """
        import zipfile
        import os
        from lxml import etree

        W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        NSMAP = {"w": W_NS, "r": R_NS}

        def log(msg):
            print(f"[HeadersFooters] {msg}")

        try:
            # ---------- 1. 打开 docx ----------
            log(f"打开文件: {output_file}")
            with zipfile.ZipFile(output_file, "r") as z:
                file_list = z.namelist()
                doc_xml = z.read("word/document.xml")
                rels_files = {f: z.read(f) for f in file_list if f.startswith("word/_rels/")}

            # ---------- 2. 解析 document.xml ----------
            doc_tree = etree.fromstring(doc_xml)
            body = doc_tree.find("w:body", NSMAP)

            # ---------- 3. 找到倒数第一、第二小节 sectPr ----------
            sectPr_nodes = body.xpath("w:sectPr", namespaces=NSMAP)
            if len(sectPr_nodes) < 2:
                log("⚠️ 文档小节不足两节，无法处理横板页眉页脚")
                return False

            last_sectPr = sectPr_nodes[-1]  # 倒数第一节
            second_last_sectPr = sectPr_nodes[-2]  # 倒数第二节

            # ---------- 4. 获取 header/footer r:id ----------
            def get_refs(sect):
                headers = {el.get(f"{{{R_NS}}}id"): el.get("w:type") for el in sect.findall("w:headerReference", NSMAP)}
                footers = {el.get(f"{{{R_NS}}}id"): el.get("w:type") for el in sect.findall("w:footerReference", NSMAP)}
                return headers, footers

            last_headers, last_footers = get_refs(last_sectPr)
            second_headers, second_footers = get_refs(second_last_sectPr)

            log(f"倒数第一节 header IDs: {last_headers}, footer IDs: {last_footers}")
            log(f"倒数第二节 header IDs: {second_headers}, footer IDs: {second_footers}")

            # ---------- 5. 替换倒数第二节页眉页脚内容 ----------
            # 遍历 header/footer r:id 替换内容
            for rid, htype in second_headers.items():
                if rid in last_headers:
                    log(f"需要将倒数第二节 header {rid} 替换为倒数第一节内容")
                    # 这里可以在 rels 文件中找到对应 headerX.xml，复制内容
                    # 示例：rels_target = rels_files["word/_rels/document.xml.rels"] 中解析 rId->target
                    # 实际替换逻辑可根据 rels 解析实现

            for rid, ftype in second_footers.items():
                if rid in last_footers:
                    log(f"需要将倒数第二节 footer {rid} 替换为倒数第一节内容")

            # ---------- 6. 可修改位置 ----------
            # 1️⃣ 修改页眉 top 距离：修改 last_sectPr 或 second_last_sectPr 中 <w:pgMar w:header="1361" />
            # 2️⃣ 修改页脚页码居中：修改 footer 中 <v:shape style="mso-position-horizontal:center;" />

            log("横板页眉页脚格式化完成")
            return True

        except Exception as e:
            log(f"⚠️ 页眉页脚处理失败: {e}")
            return False

    def preprocess_remove_auto_toc(self, source_file: str) -> str:
        """
        预处理：移除自动目录（TOC），保留分节符（w:sectPr）
        - TOC SDT 块完全删除，文字不保留
        - 如果 TOC 中有分节符，则保留
        - 如果 TOC 中无分节符，可手动插入占位段落和分节符
        - 新增：检测 TOC 前是否有“目录”（允许空白字符），没有则自动补充
        """
        import os
        import re
        import zipfile
        import tempfile
        import shutil

        print("-" * 50)
        print("🧹 步骤0-前置：移除自动目录（TOC），保留分节符")

        try:
            if not self.temp_dir:
                raise ValueError("临时目录未初始化")

            if not os.path.exists(source_file):
                print(f"❌ 文件不存在: {source_file}")
                return source_file

            base_name = os.path.splitext(os.path.basename(source_file))[0]
            output_path = os.path.join(
                self.temp_dir, f"{base_name}_去目录.docx"
            )

            shutil.copy2(source_file, output_path)

            # 解压 docx
            tmp_dir = tempfile.mkdtemp(prefix="preprocess_toc_")
            with zipfile.ZipFile(output_path, 'r') as zf:
                zf.extractall(tmp_dir)

            document_xml = os.path.join(tmp_dir, "word", "document.xml")
            if not os.path.exists(document_xml):
                print("❌ document.xml 不存在，跳过目录预处理")
                shutil.rmtree(tmp_dir)
                return source_file

            with open(document_xml, "r", encoding="utf-8") as f:
                content = f.read()

            # 匹配整个 TOC SDT 块
            sdt_pattern = re.compile(
                r'<w:sdt\b.*?>'
                r'.*?'
                r'<w:docPartGallery[^>]*w:val="Table of Contents"[^>]*/>'
                r'.*?'
                r'</w:sdt>',
                re.DOTALL
            )

            if not sdt_pattern.search(content):
                print("ℹ️ 未检测到自动目录（Table of Contents），跳过")
                shutil.rmtree(tmp_dir)
                return source_file

            def has_catalog_title(before_xml: str) -> bool:
                """
                判断 TOC 前是否存在“目录”（允许空白）
                """
                # 提取最近一个段落
                ps = re.findall(r'<w:p\b.*?</w:p>', before_xml, re.DOTALL)
                if not ps:
                    return False

                last_p = ps[-1]

                # 提取文本
                texts = re.findall(r'<w:t[^>]*>(.*?)</w:t>', last_p)
                text = ''.join(texts)

                # 去掉所有空白字符
                text = re.sub(r'\s+', '', text)

                return bool(re.search(r'目\s*录', text))

            def replace_keep_sect(match):
                sdt_xml = match.group(0)

                # ===== 新增：判断 TOC 前是否已有“目录” =====
                start_idx = match.start()
                before_xml = content[:start_idx]

                need_add_title = not has_catalog_title(before_xml)

                # ===== 原有逻辑：提取分节符 =====
                sect_p = re.search(
                    r'<w:p\b[^>]*>.*?<w:sectPr\b.*?</w:sectPr>.*?</w:p>',
                    sdt_xml,
                    re.DOTALL
                )

                if sect_p:
                    replacement = sect_p.group(0)
                else:
                    replacement = (
                        '<w:p>'
                        '<w:r><w:t>1 占位</w:t></w:r>'
                        '<w:sectPr/>'
                        '</w:p>'
                    )

                # ===== 新增：必要时插入“目录” =====
                if need_add_title:
                    title_p = (
                        '<w:p>'
                        '<w:r><w:t>目录</w:t></w:r>'
                        '</w:p>'
                    )
                    return title_p + replacement
                else:
                    return replacement

            new_content, count = sdt_pattern.subn(replace_keep_sect, content)

            if count == 0:
                print("ℹ️ 自动目录匹配失败，未发生替换")
                shutil.rmtree(tmp_dir)
                return source_file

            # 写回 document.xml
            with open(document_xml, "w", encoding="utf-8") as f:
                f.write(new_content)

            # 重新打包 docx
            tmp_docx = output_path + ".tmp"
            with zipfile.ZipFile(tmp_docx, 'w', zipfile.ZIP_DEFLATED) as zf:
                for root, _, files in os.walk(tmp_dir):
                    for file in files:
                        abs_path = os.path.join(root, file)
                        rel_path = os.path.relpath(abs_path, tmp_dir)
                        zf.write(abs_path, rel_path)

            shutil.move(tmp_docx, output_path)
            shutil.rmtree(tmp_dir)

            print(f"✅ 自动目录已移除 {count} 处（文字 + SDT），分节符已保留或手动插入")

            if self.save_intermediate_files:
                self._save_intermediate_file(
                    output_path, "step0_preprocess", "去除自动目录，保留分节符"
                )

            return output_path

        except Exception as e:
            print(f"❌ 自动目录预处理失败: {e}")
            import traceback
            traceback.print_exc()
            return source_file

    def convert_document(
            self,
            source_file: str,
            output_file: str,
            template_file: Optional[str] = None,
            header_text: str = "格式化文档",
            header_left_text: str = "",
            header_right_text: str = "",
            toc_title: str = "目      录",
            save_intermediate: bool = True,
            intermediate_dir: Optional[str] = None,
            document_type: int = 1
    ) -> tuple[bool, str]:  # 返回 bool + str:
        """
        完整的文档格式化转换流程

        Args:
            source_file: 源文档路径
            output_file: 输出文档路径
            template_file: 模板文档路径（可选，默认使用template/reference.docx）
            header_text: 页眉文本
            header_left_text: 横板页眉左侧文本
            header_right_text: 横板页眉右侧文本
            toc_title: 目录标题（可选，默认为"目 录"）
            save_intermediate: 是否保存中间文件（默认为False）
            intermediate_dir: 中间文件保存目录（仅在save_intermediate为True时有效）
            document_type: 文档类型 (1, 2, 3, 4)

        Returns:
            bool: 转换是否成功
        """
        err_msg = ""  # 初始化错误日志
        original_stdout = None
        log_file = None

        # 设置是否保存中间文件
        self.save_intermediate_files = True

        document_type_map = {
            1: "初步设计",
            2: "可行性研究",
            3: "初步设计（代可研）",
            4: "可行性研究（代初设）"
        }

        # 设置文档类型
        self.document_type = document_type
        print(f"   文档类型: {document_type}")

        # ===== 创建日志文件 =====
        log_dir = os.path.join(os.getcwd(), "log")
        os.makedirs(log_dir, exist_ok=True)

        log_filename = os.path.basename(output_file) + ".log"
        log_path = os.path.join(log_dir, log_filename)

        import sys

        class Tee:
            def __init__(self, *files):
                self.files = files

            def write(self, data):
                for f in self.files:
                    f.write(data)
                    f.flush()

            def flush(self):
                for f in self.files:
                    f.flush()

        log_file = open(log_path, "w", encoding="utf-8")

        # 保存原stdout
        original_stdout = sys.stdout

        # 同时输出到控制台 + 文件
        sys.stdout = Tee(sys.stdout, log_file)

        print(f"📝 日志文件: {log_path}")

        # ===== 保存源文件到 src_files =====
        src_dir = os.path.join(os.getcwd(), "src_files")
        os.makedirs(src_dir, exist_ok=True)

        try:

            output_filename = os.path.basename(output_file)
            file_name, ext = os.path.splitext(output_filename)

            src_save_path = os.path.join(src_dir, f"src_{file_name}{ext}")

            shutil.copy2(source_file, src_save_path)

            print(f"📥 已保存源文件到 src_files: {src_save_path}")
        except Exception as e:
            print(f"⚠️ 保存源文件失败: {e}")

        # 如果指定了中间文件目录，则使用该目录
        if save_intermediate and intermediate_dir:
            self.debug_output_dir = intermediate_dir

        # 如果未指定模板文件，则使用默认模板
        if template_file is None:
            # 获取项目根目录
            current_dir = os.path.dirname(os.path.abspath(__file__))
            parent_dir = os.path.dirname(current_dir)
            template_file = os.path.join(parent_dir, 'template', 'reference_content.docx')
            if self.save_intermediate_files:
                print(f"信息: 未指定模板文件，使用默认模板: {template_file}")

        start_time = time.time()

        # ===== 新增：步骤计时器 =====
        step_timer = {}

        def log_step(step_name, start):
            duration = time.time() - start
            print(f"⏱️ [{step_name}] 耗时: {duration:.2f} 秒")
            step_timer[step_name] = duration

        print("🚀 开始文档格式化转换")
        print("=" * 80)
        print(f"📁 源文档: {source_file}")
        print(f"📄 模板文档: {template_file}")
        print(f"📤 输出文档: {output_file}")
        print(f"📋 页眉文本: '{header_text}'")
        print(f"📋 横板页眉左侧文本: '{header_left_text}'")
        print(f"📋 横板页眉右侧文本: '{header_right_text}'")
        print(f"📋 目录标题: '{toc_title}'")
        if save_intermediate:
            print(f"💾 保存中间文件: 是")
            print(f"📂 中间文件目录: {self.debug_output_dir}")
        else:
            print(f"💾 保存中间文件: 否")
        print(f"⏰ 开始时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print("=" * 80)
        print()
        print_dochistoryid(source_file)

        # from utils.doc_parser_projName import DocParser
        # parser = DocParser(source_file)
        # proj_name, file_type_name, err_code = \
        #     parser.get_first_nonempty_paragraph_text(document_type)
        #
        # if err_code != 0:
        #     raise ValueError(f"原文档类型与所选模板类型不符，封面前两页所写类型要严格对应")
        #
        # print(f"项目名称 & 文档类型: {proj_name} & {file_type_name}")
        print("转换进度 10%，正在预处理文档", flush=True)
        # 验证输入文件
        if not self.validate_input_files(source_file, template_file):
            err_msg = "❌ 输入文件验证失败"
            return False, err_msg

        # 创建临时目录
        self.temp_dir = tempfile.mkdtemp(prefix="doc_converter_")
        if self.save_intermediate_files:
            print(f"📁 临时目录: {self.temp_dir}")

        try:

            # ===== 取消自动排序 =====
            from utils.convert_numbers_to_text_with_com import convert_auto_number_to_text
            convert_auto_number_to_text(source_file)

            # from utils.remove_auto_numbering import remove_auto_numbering_inplace
            # try:
            #     has_auto_number = remove_auto_numbering_inplace(source_file)
            #     if has_auto_number:
            #         print("⚠️ 检测到“自动排序”标题，已转换为文本编号")
            #     else:
            #         print("✔ 未检测到自动排序")
            # except Exception as e:
            #     print(f"⚠️ 自动排序处理异常: {e}")

            step_start = time.time()

            # 自动修复文档 e.g. NULL 关系
            from utils.fix_docx import fix_docx_null_relationships
            fix_docx_null_relationships(source_file)

            # 2026年1月23日 15点57分
            # 删除封面中除“设计人员”外的表格
            from utils import remove_tables_before_designers
            rtbd_success = remove_tables_before_designers\
                .DocxTableProcessor(source_file)\
                .remove_tables_before_designers()
            log_step("remove_tables_before_designers", step_start)
            if rtbd_success:
                print("remove_tables_before_designers success")
            else:
                print("remove_tables_before_designers failed")

            print_dochistoryid(source_file)

            from utils.remove_toc_sdt_test import remove_toc_sdt
            remove_toc_sdt(source_file)

            # -------------------------------------------------
            # 预处理：删除自动目录（保留分节符）
            # 2026年1月30日
            # -------------------------------------------------
            print("🧹 预处理源文档：移除自动目录")
            step_start = time.time()
            source_file = self.preprocess_remove_auto_toc(source_file)
            log_step("remove_auto_toc", step_start)
            print_dochistoryid(source_file)

            # 修复标题样式缺失问题
            step_start = time.time()
            from utils.fix_docx import fix_docx_heading_styles
            fix_docx_heading_styles(source_file)
            log_step("fix_heading_styles", step_start)

            # 预处理：清除源文档页脚所有内容
            step_start = time.time()
            from utils.del_footer import remove_all_footers_lxml
            remove_all_footers_lxml(source_file)
            log_step("remove_footers", step_start)

            # 预处理：遍历所有 sectPr 标签，判断每个小节的长和宽，
            # 如果长 < 宽，那就将这个小节 sectPr 的orient设置为横向，否则将其设置为竖向
            step_start = time.time()
            from utils.fix_section_orientation import fix_docx_section_orientation
            fix_docx_section_orientation(source_file)
            log_step("fix_orientation", step_start)

            # 10.22新增
            step_start = time.time()
            preprocessed_file = self.step0_preprocess_headings_new(source_file)
            log_step("step0_preprocess_headings", step_start)
            source_file = preprocessed_file  # 使用预处理后的文件进行后续操作
            self.step0_processed_source = preprocessed_file

            print_dochistoryid(source_file)

            # 处理封面表格
            from utils.handle_first_table import shorten_table_rows
            height_of_first_table = shorten_table_rows(source_file)

            # 步骤1: 文档拆分
            step_start = time.time()
            cover_toc_file, content_file, project_name, document_type_info = self.step1_split_document(source_file)
            log_step("step1_split_document", step_start)
            if not cover_toc_file or not content_file:
                err_msg = "文档拆分失败"
                print(f"❌ {err_msg}")
                return False, err_msg

            # 记录content_file中所有表格的字体大小
            # 并将记录保存到 table_font_sizes 中
            from utils.table_font_recorder import record_table_font_sizes

            table_font_sizes = record_table_font_sizes(content_file)

            from utils.table_font_setter import set_table_font_small_five
            set_table_font_small_five(cover_toc_file)

            # 将该步骤放在文档拆分之后，是为了拿到项目名称和文档类型
            # 步骤0: 页眉页脚替换
            updated_template_file = self.step0_replace_header_footer(source_file, template_file, project_name,
                                                                     document_type_map.get(document_type))
            # 使用更新后的模板文件进行后续处理
            template_file = updated_template_file

            from utils.remove_header_image import process_docx
            process_docx(content_file, template_file)

            original_content_for_tables = content_file

            # 步骤1.4: 按横竖段落拆分并保存临时文件 (已合并到1.5中，此处移除以避免重复)
            # try:
            #     _, _ = self.step1_4_split_by_orientation_segments(content_file)
            # except Exception as step14_error:
            #     print(f"⚠️ 步骤1.4 处理失败，继续后续流程: {step14_error}")
            print("转换进度 20%，正在按段落拆分文档", flush=True)
            # 步骤1.5: 按横竖段落拆分 + 分段处理 + 按序合并（横版跳过Pandoc/表格）
            step_start = time.time()
            formula_info_list: List[Dict[str, Any]] = []
            pythoncom.CoInitialize()
            word = win32.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            word.AutomationSecurity = 3
            content_file, formula_info_all, err_msg, picture_info_all = self.step1_5_split_by_orientation_segments_pipeline(
                content_file=content_file,
                template_file=template_file,
                word=word
            )
            log_step("step1_5_pipeline", step_start)
            if err_msg != "":
                return False, err_msg

            from utils.rm_pic_ind import remove_picture_paragraph_indent
            remove_picture_paragraph_indent(content_file)

            landscape_sections = []
            if word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass
                print("[聯合] Word 应用已退出。")
                del word
            # pythoncom.CoUninitialize()

            # 步骤2/3 已在 step1.5 中按段落处理完成
            table_formatted_file = content_file
            print("转换进度 40%，正在合并文档", flush=True)
            # 步骤4: 文档合并
            step_start = time.time()
            success = self.step4_merge_documents(cover_toc_file, table_formatted_file, output_file)
            log_step("step4_merge_documents", step_start)

            if success:
                # 步骤4.1: 插入完整目录
                step_start = time.time()
                toc_insert_success = self.step4_1_insert_toc_after_cover(output_file, toc_title)
                log_step("step4_1_insert_toc", step_start)
                if not toc_insert_success:
                    print("⚠️ 插入完整目录失败，继续执行")
                    return False, "插入完整目录失败"

                # 步骤4.5: 在文档末尾追加横向
                step_start = time.time()
                if landscape_sections:
                    valid_sections: List[str] = []
                    for sec_path in landscape_sections:
                        if not sec_path:
                            continue
                        abs_path = os.path.abspath(sec_path)
                        if not os.path.exists(abs_path):
                            continue
                        try:
                            if os.path.getsize(abs_path) <= 0:
                                continue
                        except Exception:
                            continue
                        valid_sections.append(abs_path)
                    if valid_sections:
                        print("-" * 50)
                        print("步骤4.5: 在合并文档末尾追加横向")
                        if not self.temp_dir:
                            raise ValueError("临时目录未初始化")
                        base_name = os.path.splitext(os.path.basename(output_file))[0]
                        appended_outputs: List[str] = []
                        current_base = output_file
                        for idx, section_path in enumerate(valid_sections, 1):
                            step45_output = os.path.join(
                                self.temp_dir,
                                f"{base_name}_step45_append_landscape_{idx:02d}.docx",
                            )
                            append_success = self._append_document_with_formatting(
                                base_file=current_base,
                                file_to_append=section_path,
                                output_file=step45_output,
                            )
                            if append_success and os.path.exists(step45_output):
                                shutil.copy2(step45_output, output_file)
                                current_base = step45_output
                                appended_outputs.append(step45_output)
                                if self.save_intermediate_files:
                                    self._save_intermediate_file(
                                        step45_output,
                                        "step4_5_append",
                                        f"landscape_section_{idx}",
                                    )
                            else:
                                print(f"[warn] step4.5: failed to append landscape section: {os.path.basename(section_path)}")
                                return False, "追加横向段落失败"
                        if appended_outputs:
                            self.intermediate_files['landscape_appended'] = ';'.join(appended_outputs)
                        else:
                            self.intermediate_files.pop('landscape_appended', None)
                            print("[warn] step4.5: no landscape sections appended")
                            return False, "追加横向段落失败"
                    else:
                        print("[warn] step4.5: landscape sections empty, skip append")
                        self.intermediate_files.pop('landscape_appended', None)
                        return False, "追加横向段落失败"
                else:
                    print("[info] step4.5: no landscape sections, skip append")
                    self.intermediate_files.pop('landscape_appended', None)
                log_step("step4.5", step_start)
                print("转换进度 60%，正在更新目录标题", flush=True)

                # 步骤5: 更新目录标题
                step_start = time.time()
                toc_update_success = self.step5_update_toc_title(output_file, toc_title)
                log_step("step5_update_toc_title", step_start)
                # if toc_update_success:
                #     print("✅ 目录标题更新完成!")
                # else:
                #     print("⚠️ 目录标题更新失败，继续使用原有标题")
                if not toc_update_success:
                    print("⚠️ 目录标题更新失败，继续使用原有标题")
                    return False, "目录标题更新失败"
                print("转换进度 70%，正在格式化图片", flush=True)

                print(formula_info_list)

                # 步骤6: 图片格式化
                step_start = time.time()
                picture_format_success = self.step6_format_pictures(output_file, formula_info_list)
                log_step("step6_format_pictures", step_start)
                if not picture_format_success:
                    print("⚠️ 图片格式化失败，继续使用原有格式")
                    return False, "图片格式化失败"

                # 步骤7: 库号信息格式化
                step_start = time.time()
                library_number_format_success = self.step7_format_library_number(output_file)
                log_step("step7_library_number", step_start)
                if not library_number_format_success:
                    print("⚠️ 库号信息格式化失败，继续使用原有格式")
                    return False, "库号信息格式化失败"

                # 步骤8: 在目录后插入分节符
                step_start = time.time()
                section_break_insert_success = self.step8_insert_section_break(output_file)
                log_step("step8_insert_section_break", step_start)
                if not section_break_insert_success:
                    print("⚠️ 在目录后插入分节符失败，继续使用原有格式")
                    return False, "在目录后插入分节符失败"

                # 步骤9: 处理文档节的页码设置
                step_start = time.time()
                section_page_process_success = self.step9_process_sections(output_file)
                if not section_page_process_success:
                    print("⚠️ 处理文档节的页码设置失败，继续使用原有格式")
                    return False, "处理文档节的页码设置失败"

                # ============================
                # 🔽 新增：保存步骤9结果文件
                # ============================
                if self.save_intermediate_files:
                    try:
                        step9_output = os.path.join(
                            self.temp_dir,
                            f"{os.path.splitext(os.path.basename(output_file))[0]}_step9_after_sections.docx"
                        )
                        shutil.copy2(output_file, step9_output)

                        self._save_intermediate_file(
                            step9_output,
                            "step9",
                            "after_section_page_process"
                        )

                        print(f"💾 已保存步骤9结果文件: {step9_output}")

                    except Exception as e:
                        print(f"⚠️ 保存步骤9中间文件失败: {e}")
                log_step("step9_process_sections", step_start)
                print("转换进度 80%，正在更新目录域", flush=True)
                # 步骤9.5: 更新目录域 (刷新页码和标题)
                step_start = time.time()
                print("-" * 50)
                print("🔄 步骤9.5: 更新目录域")
                if COM_AVAILABLE and win32 and pythoncom:
                    word = None
                    doc = None
                    com_initialized = False
                    try:
                        pythoncom.CoInitialize()
                        com_initialized = True
                        word = win32.DispatchEx("Word.Application")
                        word.Visible = False
                        word.DisplayAlerts = 0
                        try:
                            word.AutomationSecurity = 3
                        except Exception:
                            pass
                        doc = word.Documents.Open(output_file, ReadOnly=False)
                        for toc in doc.TablesOfContents:
                            try:
                                toc.Update()
                            except Exception:
                                pass
                        try:
                            doc.Fields.Update()
                        except Exception:
                            pass
                        doc.Save()
                        print("✅ 目录域更新成功!")
                    except Exception as e:
                        print(f"⚠️ 目录域更新失败: {e}")
                        return False, "目录域更新失败"
                    finally:
                        if doc is not None:
                            try:
                                doc.Close(SaveChanges=0)
                            except Exception:
                                pass
                        if word is not None:
                            try:
                                word.Quit()
                            except Exception:
                                pass
                        if com_initialized:
                            try:
                                pythoncom.CoUninitialize()
                            except Exception:
                                pass
                else:
                    print("⚠️ COM 不可用，跳过目录域更新")
                    return False, "目录域更新失败"
                log_step("step9.5", step_start)
                # 步骤10: 删除文档中所有的突出显示
                step_start = time.time()
                no_highlights_file = self.step10_remove_highlights(output_file)
                # 如果成功去除突出显示，将结果复制回输出文件
                if no_highlights_file != output_file and os.path.exists(no_highlights_file):
                    shutil.copy2(no_highlights_file, output_file)
                    print("✅ 突出显示删除成功!")
                else:
                    print("⚠️ 删除突出显示失败，继续使用原有格式")
                    return False, "删除突出显示失败"
                log_step("step10_remove_highlights", step_start)

                # 步骤11: 高级库号信息格式化
                step_start = time.time()
                library_number_advanced_format_success = self.step11_format_library_number_advanced(output_file)
                if not library_number_advanced_format_success:
                    print("⚠️ 库号信息高级格式化失败，继续使用原有格式")
                    return False, "库号信息高级格式化失败"
                log_step("step11_advanced_library", step_start)

                # 步骤11.5: 基于 step1.9 映射回插图片（将输出覆盖回 output_file，供步骤12使用）
                step_start = time.time()
                step11_5_success = self.step11_5_insert_by_map(output_file)
                if not step11_5_success:
                    print("⚠️ step11.5 插图回填失败/跳过，继续使用当前文档")
                    return False, "插图回填失败"
                print("转换进度 90%，正在处理文档细节", flush=True)
                log_step("step11_5_insert_by_map", step_start)

                step_start = time.time()
                # 步骤12 预处理: 删除与首张表重复的表格
                self.step12_remove_duplicate_tables(output_file)
                # 步骤12 预处理: 删除重复段落
                self.step12_remove_duplicate_phrase(output_file)
                log_step("step12", step_start)

                # 步骤13: 横板页眉左右对齐
                # landscape_header_success = self.step13_format_landscape_headers(
                #     output_file,
                #     header_left_text,
                #     header_right_text
                # )
                # if not landscape_header_success:
                #     print("⚠️ 横板页眉处理失败/跳过，继续使用当前文档")

                # 步骤13：表格替换
                step_start = time.time()
                if self.step0_processed_source:
                    table_replace_success = self.step13_word_table_replace(
                        edited_docx=output_file,
                        original_docx=self.step0_processed_source
                    )
                    if not table_replace_success:
                        print('⚠️ 表格替换失败，继续使用当前文档')
                        return False, "表格替换失败"
                else:
                    print('⚠️ 跳过表格替换: 未获取步骤0的输出文件')
                    return False, "表格替换失败"

                # 设置output_file中所有表格的字体大小
                from utils.table_font_recorder import restore_table_font_sizes

                # ====== 👇 在这里插入 ======
                new_item = {
                    "font_name": "宋体",
                    "font_size": 14.0
                }

                table_font_sizes = {
                    k + 1 if isinstance(k, int) else k: v
                    for k, v in table_font_sizes.items()
                }

                table_font_sizes[0] = new_item

                # 保证顺序正确（很重要）
                table_font_sizes = dict(sorted(table_font_sizes.items(), key=lambda x: x[0]))

                print("++++++++++++++++++++++")
                print(table_font_sizes)
                print("++++++++++++++++++++++")

                # =========================
                restore_table_font_sizes(output_file, table_font_sizes)

                from utils.table_font_setter import set_table_font_four
                set_table_font_four(output_file)

                from utils.handle_first_table import restore_table_rows
                restore_table_rows(output_file, height_of_first_table)
                log_step("step13_table_replace", step_start)

                # headers_footers_process_success = self.step12_post_process_headers_footers(output_file)
                # if not headers_footers_process_success:
                #     print("⚠️ 页眉页脚处理失败，继续使用原有格式")

                # FIX：2026年1月22日 20点05分 —— 22点18分
                # 步骤14：格式化横板页眉页脚
                step_start = time.time()
                from utils import format_landscape_headers_footers
                step14_success = (format_landscape_headers_footers.DocxProcessor()
                                  .step14_format_landscape_headers_footers(output_file))
                if step14_success:
                    print("step14 SUCCESS")
                else:
                    print("step14 FAILED")
                    return False, "页眉页脚处理失败"
                log_step("step14_headers_footers", step_start)

                # 调用清理样式函数
                from utils.remove_heading_keep_options import DocxStyleCleaner

                # 在 convert_document 方法末尾
                cleaner = DocxStyleCleaner()  # 创建实例
                success = cleaner.remove_keep_options_from_heading_styles(output_file)
                if not success:
                    print("清理标题样式失败，但继续执行")  # 可选处理

                print("===========================================")
                print("\n📊 各步骤耗时统计（降序）:")
                for k, v in sorted(step_timer.items(), key=lambda x: x[1], reverse=True):
                    print(f"{k}: {v:.2f}s")
                # print(picture_info_all)
                # from utils.picture_format_keeper import restore_picture_formats
                # restore_picture_formats(output_file, picture_info_all)
                print("===========================================")

                end_time = time.time()
                duration = end_time - start_time

                print("\n" + "=" * 80)
                print("✅ 文档转换成功!")
                print(f"⏱️ 总耗时: {duration:.2f} 秒")
                print(f"📤 最终文档: {output_file}")

                # ===== 保存目标文件到 obj_files =====
                obj_dir = os.path.join(os.getcwd(), "obj_files")
                os.makedirs(obj_dir, exist_ok=True)

                try:
                    shutil.copy2(
                        output_file,
                        os.path.join(obj_dir, os.path.basename(output_file))
                    )
                    print(f"📦 已保存目标文件到 obj_files")
                except Exception as e:
                    print(f"⚠️ 保存目标文件失败: {e}")

                # 显示中间文件信息
                if save_intermediate:
                    print("\n📋 中间文件保存在临时目录:")
                    for key, path in self.intermediate_files.items():
                        if os.path.exists(path):
                            print(f"   {key}: {os.path.basename(path)}")

                    print(f"\n📁 所有中间文件已同步保存到: {self.debug_output_dir}")
                    print("🔍 您可以在该目录中查看每个步骤的处理结果，便于调试和优化")
                else:
                    print("\n📋 中间文件未保存（根据设置）")

                print("\n💡 提示:")
                print("   - 在Word中打开文档，右键目录选择'更新域'来刷新页码")
                print("   - 检查文档格式是否符合要求")
                print("=" * 80)
                print("转换结束！")
                print("转换进度 100%，点击重新转换", flush=True)
                return True, "文档转换成功"
            else:
                print("❌ 转换失败: 文档合并失败")
                return False, "文档合并失败"

        except Exception as e:
            print(f"\n❌ 转换过程中发生错误: {str(e)}")
            return False, "转换过程中发生错误"

        finally:
            # 不立即清理临时文件，保留中间结果供调试
            if self.save_intermediate_files:
                print(f"\n📁 临时文件保留在: {self.temp_dir}")
                print("   您可以手动删除该目录，或重启程序时自动清理")

            # ===== 恢复stdout =====
            try:
                if original_stdout:
                    sys.stdout = original_stdout
                if log_file:
                    log_file.close()
            except Exception:
                pass


def quick_convert_document(
        source_file: str,
        output_file: str,
        template_file: Optional[str] = None,
        header_text: str = "格式化文档",
        toc_title: str = "目 录",
        save_intermediate: bool = False,
        intermediate_dir: Optional[str] = None,
        document_type: int = 1
) -> bool:
    """
    便捷函数: 快速进行文档格式化转换

    Args:
        source_file: 源文档路径
        output_file: 输出文档路径
        template_file: 模板文档路径（可选，默认使用template/reference.docx）
        header_text: 页眉文本
        toc_title: 目录标题（可选，默认为"目 录"）
        save_intermediate: 是否保存中间文件（默认为False）
        intermediate_dir: 中间文件保存目录（仅在save_intermediate为True时有效）
        document_type: 文档类型 (1, 2, 3, 4)

    Returns:
        bool: 转换是否成功
    """
    with DocumentConverter(document_type=document_type) as converter:
        return converter.convert_document(
            source_file=source_file,
            output_file=output_file,
            template_file=template_file,
            header_text=header_text,
            toc_title=toc_title,
            save_intermediate=save_intermediate,
            intermediate_dir=intermediate_dir,
            document_type=document_type
        )


if __name__ == "__main__":
    source_file = r"C:\Users\yanha\Desktop\数字总师\文档\可行性报告（test）.docx"

    # 创建result目录
    result_dir = os.path.join(current_dir, "result")
    os.makedirs(result_dir, exist_ok=True)

    # 输出文件路径
    output_file = os.path.join(result_dir, "formatted_document.docx")

    # 使用默认模板（template/reference.docx）

    # 执行转换
    with DocumentConverter(document_type=1) as converter:  # 添加文档类型参数
        success = converter.convert_document(
            source_file=source_file,
            output_file=output_file,
            header_text="数字总师可行性报告",
            toc_title="目      录",
            save_intermediate=False,
            document_type=1  # 添加文档类型参数
        )

