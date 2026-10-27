import os
import json
import xml.etree.ElementTree as ET
from xml.dom import minidom
from typing import Dict, Any, List, Tuple
from docx import Document
from docx.oxml.ns import qn
import zipfile
import tempfile
import shutil
import re
import win32com.client
import pythoncom
import time



from service.dify_service import DifyService
# 导入Document类（忽略静态分析工具的警告）
# pyright: reportAttributeAccessIssue=false
# try:
#     from docx import Document
# except ImportError:
#     Document = None

"""
WordStyler - 文档模板内容替换工具

该模块提供了多种方式来替换DOCX文档模板中的内容，同时保持原有格式不变。
支持两种主要的实现方式：
1. 基于XML节点处理的方式
2. 基于python-docx库API的方式

主要功能：
- 从配置文件加载替换映射关系
- 从文档中提取内容并创建替换字典
- 支持跨多个文本节点的内容替换
- 保持原有文档格式
- 支持临时文件和持久化文件保存选项

函数说明：
- load_replace_config: 加载替换配置文件
- save_replace_config: 保存替换配置文件（保留此函数以供其他模块使用）
- replace_text_in_runs: 在段落的runs中替换文本，保持格式
- get_paragraph_text: 获取段落的完整文本
- create_replacement_dict_from_content: 根据提取的内容和配置的key创建替换字典
- extract_text_from_xml: 从docx文件的XML中提取所有文本内容
* replace_content_in_template_docx: 使用python-docx方式替换模板中的内容
* replace_content_in_template_xml_to_docx: 使用XML方式处理并生成DOCX文件
- extract_content_from_docx: 从docx文件中提取内容
- match_content_to_config: 将提取的内容匹配到配置键
- replace_content_in_cover: 快捷函数，从给定文件中识别配置key对应的value并替换模板内容
"""


COVER_DATE_PATTERN = re.compile(
    r"(\d{4}年\d{1,2}月\d{1,2}[日号]|\d{4}-\d{1,2}-\d{1,2}|\d{4}/\d{1,2}/\d{1,2}|\d{4}年\d{1,2}月)"
)
COVER_DATE_YMD_PATTERN = re.compile(
    r"(\d{4}年\d{1,2}月\d{1,2}[日号]|\d{4}-\d{1,2}-\d{1,2}|\d{4}/\d{1,2}/\d{1,2})"
)
COVER_DATE_YM_PATTERN = re.compile(r"\d{4}年\d{1,2}月")


def _normalize_cover_date_text(text: str) -> str:
    """清理封面日期中的空白字符（含全角空格）。"""
    if text is None:
        return ""
    return re.sub(r"[\s\u3000]+", "", str(text)).strip()


def _extract_normalized_cover_date(text: str) -> str:
    """从文本中提取日期并去空格，未匹配时返回空字符串。"""
    normalized = _normalize_cover_date_text(text)
    if not normalized:
        return ""
    match = COVER_DATE_PATTERN.search(normalized)
    if match:
        return match.group(1)
    return ""


def convert_new_config_to_old_format(config: Dict[str, Any]) -> Dict[str, str]:
    """
    将新格式的配置转换为旧格式
    
    Args:
        config: 新格式的配置
        
    Returns:
        Dict[str, str]: 旧格式的替换映射
    """
    replace_dict = {}
    
    # 添加项目名 - 使用固定的键
    if "项目名" in config:
        replace_dict["项目名"] = config.get("项目名", "")
    
    # 添加时间 - 使用固定的键
    if "时间" in config:
        replace_dict["时间"] = config.get("时间", "")
    
    # 添加库号 - 使用固定的键
    if "库号" in config:
        replace_dict["库号"] = config.get("库号", "")
    
    # 添加人员信息 - 按顺序添加，使用占位符格式
    if "人员" in config and isinstance(config["人员"], list):
        # 处理所有人员，不再限制为4个
        for i, person in enumerate(config["人员"]):
            if isinstance(person, dict) and "姓名" in person and "职务" in person:
                # 构造人员信息字符串（使用占位符格式）
                person_info = f"{person['职务']}：{person['姓名']}"
                # 使用固定的键来避免冲突
                replace_dict[f"人员_{i}"] = person_info
    
    return replace_dict

def load_replace_config(config_path: str) -> Dict[str, str]:
    """
    加载替换配置文件
    
    Args:
        config_path: 配置文件路径
        
    Returns:
        Dict[str, str]: key-value替换映射
    """
    with open(config_path, 'r', encoding='utf-8') as f:
        config = json.load(f)
    
    # 检查配置文件格式
    if "替换内容" in config:
        # 旧格式
        return config.get("替换内容", {})
    elif "项目名" in config:
        # 新格式，需要转换为旧格式用于替换
        return convert_new_config_to_old_format(config)
    else:
        return {}


def save_replace_config(config_path: str, replace_dict: Dict[str, str]) -> None:
    """
    保存替换配置文件 (保留此函数以供其他模块使用，但按用户要求不在此模块中调用)
    
    Args:
        config_path: 配置文件路径
        replace_dict: key-value替换映射
    """
    config = {"替换内容": replace_dict}
    with open(config_path, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=4)


def replace_text_in_runs(runs, old_text, new_text):
    """
    在runs中替换文本，保持格式
    
    Args:
        runs: 段落的runs列表
        old_text: 要替换的旧文本
        new_text: 新文本
    """
    if not old_text or not new_text:
        return
    
    # 收集所有run的文本
    full_text = ''.join(run.text for run in runs)
    
    # 检查是否包含要替换的文本
    if old_text in full_text:
        # 找到包含旧文本的第一个run的位置
        start_pos = full_text.find(old_text)
        end_pos = start_pos + len(old_text)
        
        # 如果新旧文本长度相同，保持原有结构
        if len(old_text) == len(new_text):
            # 逐字符替换，保持run结构
            current_pos = 0
            new_text_index = 0
            
            for run in runs:
                run_len = len(run.text)
                # 检查run是否与要替换的文本有交集
                if current_pos < end_pos and current_pos + run_len > start_pos:
                    # 计算在当前run中需要替换的部分
                    local_start = max(0, start_pos - current_pos)
                    local_end = min(run_len, end_pos - current_pos)
                    
                    # 替换文本
                    if local_start < local_end:
                        prefix = run.text[:local_start]
                        suffix = run.text[local_end:]
                        replacement = new_text[new_text_index:new_text_index + (local_end - local_start)]
                        new_text_index += len(replacement)
                        run.text = prefix + replacement + suffix
                current_pos += run_len
        else:
            # 长度不同的情况，清空所有相关的run，并在第一个run中放入新文本
            current_pos = 0
            first_run = None
            prefix_text = ""
            suffix_text = ""
            
            for run in runs:
                run_len = len(run.text)
                # 检查run是否与要替换的文本有交集
                if current_pos < end_pos and current_pos + run_len > start_pos:
                    if first_run is None:
                        first_run = run
                        # 保存前缀文本
                        if current_pos < start_pos:
                            prefix_text = full_text[current_pos:start_pos]
                    # 清空文本但保留格式
                    run.text = ""
                elif current_pos + run_len > end_pos and first_run is not None:
                    # 保存后缀文本
                    suffix_start = max(0, end_pos - current_pos)
                    if suffix_start < run_len:
                        suffix_text = run.text[suffix_start:]
                    run.text = ""
                current_pos += run_len
            
            # 在第一个run中设置新文本
            if first_run is not None:
                # 计算前缀文本
                prefix_len = start_pos
                for run in runs:
                    if run == first_run:
                        break
                    prefix_len -= len(run.text)
                
                if prefix_len > 0:
                    prefix_text = first_run.text[:prefix_len]
                
                first_run.text = prefix_text + new_text + suffix_text


def get_paragraph_text(paragraph):
    """
    获取段落的完整文本
    
    Args:
        paragraph: python-docx段落对象
        
    Returns:
        str: 段落的完整文本
    """
    return ''.join(run.text for run in paragraph.runs)


def iter_paragraphs_from_tables(tables):
    for table in tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    yield paragraph


def iter_all_paragraphs(doc):
    for paragraph in doc.paragraphs:
        yield paragraph
    for paragraph in iter_paragraphs_from_tables(doc.tables):
        yield paragraph
    for section in doc.sections:
        for header in (section.header, section.first_page_header, section.even_page_header):
            for paragraph in header.paragraphs:
                yield paragraph
            for paragraph in iter_paragraphs_from_tables(header.tables):
                yield paragraph
        for footer in (section.footer, section.first_page_footer, section.even_page_footer):
            for paragraph in footer.paragraphs:
                yield paragraph
            for paragraph in iter_paragraphs_from_tables(footer.tables):
                yield paragraph


def apply_character_spacing_to_cover(doc, spacing_pt=7):
    """
    检查封面内容，如果包含"初步设计"或"可行性研究"，则设置其字间距。
    """
    target_texts = ["初步设计", "可行性研究"]
    from docx.oxml import OxmlElement
    from docx.enum.text import WD_LINE_SPACING
    
    for paragraph in iter_all_paragraphs(doc):
        full_text = get_paragraph_text(paragraph).strip()
        
        # 检查是否包含目标文本
        is_target = False
        for target in target_texts:
            if target in full_text:
                is_target = True
                break
        
        if not is_target:
            continue
            
        # 找到了，对包含目标文本的 run 设置间距
        # 同时修正行距，避免大字号被“行距过小”裁切
        try:
            paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            paragraph.paragraph_format.line_spacing = None
            paragraph.paragraph_format.space_before = 0
            paragraph.paragraph_format.space_after = 0
        except Exception:
            pass

        for run in paragraph.runs:
            # 如果这一段很短（就是标题），或者 run 包含关键字
            if len(full_text) < 20 or any(t in run.text for t in target_texts):
                 r_pr = run._element.get_or_add_rPr()
                 # 7 pt = 140 twips (1 point = 20 twips)
                 spacing_val = int(spacing_pt * 20)
                 
                 spacing = r_pr.find(qn('w:spacing'))
                 if spacing is None:
                     spacing = OxmlElement('w:spacing')
                     r_pr.append(spacing)
                 
                 spacing.set(qn('w:val'), str(spacing_val))


def force_cover_title_spacing_in_docx(file_path, spacing_pt=7):
    """
    打开指定路径的 Word 文件，调整封面标题字间距后保存。
    """
    if not os.path.exists(file_path):
        return
    try:
        from docx import Document
        doc = Document(file_path)
        apply_character_spacing_to_cover(doc, spacing_pt=spacing_pt)
        doc.save(file_path)
    except Exception as e:
        print(f"force_cover_title_spacing_in_docx error: {e}")


def clear_run_spacing(doc, keywords=None):
    if keywords is None:
        keywords = ["可行性研究", "初步设计", "代可行", "代初步"]
    for paragraph in iter_all_paragraphs(doc):
        paragraph_text = get_paragraph_text(paragraph)
        if not paragraph_text or not any(keyword in paragraph_text for keyword in keywords):
            continue
        for run in paragraph.runs:
            r_pr = run._element.find(qn('w:rPr'))
            if r_pr is None:
                continue
            spacing = r_pr.find(qn('w:spacing'))
            if spacing is not None:
                r_pr.remove(spacing)


def extract_cover_paragraphs(
        file_path: str,
        marker: str = "各专业参加设计人员名单",
        skip_empty: bool = True,
        filter_design_terms: bool = True  # False：不启用过滤。True（默认）：会过滤掉包含以下任意关键词的段落（忽略大小写和空格）
) -> Tuple[List[str], bool]:
    """
    从 Word 文件开头提取普通段落文本，直到遇到 marker（忽略空格和大小写）。
    可选地过滤掉包含"可行性研究""初步设计""（代初步）""（代可行）"等关键词的段落。

    参数:
        file_path: str, .docx 文件路径
        marker: str, 停止标记
        skip_empty: bool, 是否跳过空段落
        filter_design_terms: bool, 是否过滤掉设计类段落,False：不启用过滤。True（默认）：会过滤掉包含以下任意关键词的段落（忽略大小写和空格）

    返回:
        paragraphs: list[str] — 提取到的段落文本
        found_marker: bool — 是否找到标记
    """

    if not os.path.exists(file_path):
        raise FileNotFoundError(f"文件不存在: {file_path}")

    # 先检查是否已导入Document
    if Document is None:
        raise ImportError("未找到python-docx库，请安装：pip install python-docx")

    doc = Document(file_path)
    paragraphs = []
    found_marker = False

    # 预处理 marker：忽略空格与大小写
    marker_pattern = re.sub(r"\s+", "", marker).lower()

    # 定义过滤关键词
    design_keywords = [
        "可行性研究",
        "初步设计",
        "代初步",
        "代可行",
    ]

    # 构建正则：允许空格、括号、多种形式
    design_pattern = re.compile(
        r"[\(（]?\s*(?:代初步|代可行)\s*[\)）]?|可行性研究|初步设计",
        flags=re.IGNORECASE
    )

    for para in doc.paragraphs:
        text = para.text or ""
        if skip_empty and not text.strip():
            continue

        # 去掉所有空格，统一中英文冒号
        text_clean = re.sub(r"\s+", "", text)
        text_clean_lower = text_clean.lower()

        # 检测停止标志
        if marker_pattern in text_clean_lower.replace("：", "").replace(":", ""):
            found_marker = True
            break

        # 将全角括号替换为半角
        normalized = text.replace("（", "(").replace("）", ")")

        # 检查过滤关键词
        if filter_design_terms and design_pattern.search(normalized):
            continue

        paragraphs.append(text.strip())

    return paragraphs, found_marker


def extract_content_from_docx(docx_path: str) -> Dict[str, str]:
    """
    从docx文件中提取内容（使用更健壮的方法）
    
    Args:
        docx_path: docx文件路径
        
    Returns:
        Dict[str, str]: 提取的内容字典
    """
    # 先检查是否已导入Document
    if Document is None:
        raise ImportError("未找到python-docx库，请安装：pip install python-docx")
    
    # 使用新的提取方法
    paragraphs, _ = extract_cover_paragraphs(docx_path)
    
    # 根据用户说明，索引为 0, 2, 3, 5, 6, 7 的内容分别对应配置文件中要替换的 6 项内容
    content_dict = {}
    
    # 确保有足够的段落内容
    if len(paragraphs) > 7:
        content_dict["content_0"] = paragraphs[0]  # 项目名称
        content_dict["content_1"] = paragraphs[2]  # 日期
        content_dict["content_2"] = paragraphs[3]  # 库号
        content_dict["content_3"] = paragraphs[5]  # 主管领导
        content_dict["content_4"] = paragraphs[6]  # 副总工程师
        content_dict["content_5"] = paragraphs[7]  # 总设计师
    else:
        # 如果段落数量不足，抛出异常或使用空值
        # 根据项目规范，应该确保能提取到所需内容
        raise ValueError(f"文档内容不足，仅提取到{len(paragraphs)}段，需要至少8段内容")
    
    return content_dict


def match_content_to_config_new_format(content_dict: Dict[str, str], config: Dict[str, Any]) -> Dict[str, str]:
    """
    将提取的内容匹配到新格式的配置
    
    Args:
        content_dict: 从文档中提取的内容
        config: 新格式的配置
        
    Returns:
        Dict[str, str]: 匹配后的替换字典
    """
    replace_dict = {}
    
    # 获取内容值列表
    content_values = list(content_dict.values())
    
    # 初始化匹配变量
    project_name = ""
    date_text = ""
    library_number = ""
    personnel_list = []
    
    # 遍历提取的内容，根据特征进行匹配
    for content_item in content_values:
        # 跳过目录等无关内容
        if content_item in ["目录", "第一章 摘要      1", "第二章  背景概述 3", "§2.1  政策背景   3", 
                           "§2.2  经济背景   4", "§2.3  产业背景   6", "第三章  电池市场分析     9"]:
            continue
            
        # 匹配项目名称（以"项目"结尾）
        if not project_name and content_item.strip().endswith("项目"):
            # 进一步验证不是日期或其他信息
            if not any(date_keyword in content_item for date_keyword in ['年', '月', '日']) or \
               content_item.count('年') < 2:  # 避免将日期误识别为项目名
                project_name = content_item
                
        # 匹配日期（支持年月日和年月格式），先去空格后判断
        if not date_text:
            normalized_content = _normalize_cover_date_text(content_item)
            has_digit = any(char.isdigit() for char in normalized_content)
            if has_digit:
                matched_date = _extract_normalized_cover_date(normalized_content)
                if matched_date:
                    # 优先选择更具体的日期（包含年份的）
                    if '202' in matched_date or '201' in matched_date:  # 2020年代或2010年代
                        date_text = matched_date
                    elif not date_text:  # 如果还没有找到日期，则使用第一个匹配的
                        date_text = matched_date
                        
        # 匹配库号（保持原有逻辑但增强鲁棒性）
        if not library_number:
            # 更宽松的库号匹配规则
            # 检查是否包含库号关键字或者看起来像库号的字符串（数字和字母组合）
            if ('库号' in content_item and (':' in content_item or '：' in content_item)) or \
               (any(keyword in content_item for keyword in ['库号', '编号', 'NO.', 'No']) and 
                any(char.isdigit() for char in content_item)) or \
               (len(content_item) > 10 and any(char.isdigit() for char in content_item) and 
                any(char.isalpha() for char in content_item)):  # 看起来像库号的字符串（长度>10，包含数字和字母）
                library_number = content_item
                
        # 匹配人员信息（收集所有可能的人员信息）
        if ('主管领导' in content_item or '副总工程师' in content_item or 
            '总设计师' in content_item or '工程师' in content_item or 
            '领导' in content_item or '设计师' in content_item):
            personnel_list.append(content_item)
    
    # 填充替换字典
    replace_dict["项目名"] = project_name
    replace_dict["时间"] = _normalize_cover_date_text(date_text)
    replace_dict["库号"] = library_number
    
    # 处理人员信息，支持任意数量的人员
    for i, person_info in enumerate(personnel_list):
        replace_dict[f"人员_{i}"] = person_info
    
    return replace_dict

def match_content_to_config(content_dict: Dict[str, str], config_keys: list) -> Dict[str, str]:
    """
    将提取的内容匹配到配置键
    使用更智能的匹配逻辑，基于内容特征进行匹配
    
    Args:
        content_dict: 从文档中提取的内容
        config_keys: 配置文件中的键列表
        
    Returns:
        Dict[str, str]: 匹配后的替换字典
    """
    # 初始化替换字典
    replace_dict = {key: "" for key in config_keys}
    
    # 获取内容值列表
    content_values = list(content_dict.values())
    
    # 初始化匹配变量
    project_name = ""
    date_text = ""
    library_number = ""
    leader_info = ""
    vp_engineer_info = ""
    chief_designer_info = ""
    
    # 遍历提取的内容，根据特征进行匹配
    for content_item in content_values:
        # 跳过目录等无关内容
        if content_item in ["目录", "第一章 摘要      1", "第二章  背景概述 3", "§2.1  政策背景   3", 
                           "§2.2  经济背景   4", "§2.3  产业背景   6", "第三章  电池市场分析     9"]:
            continue
            
        # 匹配项目名称（以"项目"结尾）
        if not project_name and content_item.strip().endswith("项目"):
            # 进一步验证不是日期或其他信息
            if not any(date_keyword in content_item for date_keyword in ['年', '月', '日']) or \
               content_item.count('年') < 2:  # 避免将日期误识别为项目名
                project_name = content_item
                continue
            
        # 匹配日期（支持年月日和年月格式），先去空格后判断
        if not date_text:
            normalized_content = _normalize_cover_date_text(content_item)
            has_digit = any(char.isdigit() for char in normalized_content)
            if has_digit:
                matched_date = _extract_normalized_cover_date(normalized_content)
                if matched_date:
                    # 优先选择更具体的日期（包含年份的）
                    if '202' in matched_date or '201' in matched_date:  # 2020年代或2010年代
                        date_text = matched_date
                    elif not date_text:  # 如果还没有找到日期，则使用第一个匹配的
                        date_text = matched_date
                    continue
                    
        # 匹配库号（保持原有逻辑但增强鲁棒性）
        if not library_number:
            # 更宽松的库号匹配规则
            # 检查是否包含库号关键字或者看起来像库号的字符串（数字和字母组合）
            if ('库号' in content_item and (':' in content_item or '：' in content_item)) or \
               (any(keyword in content_item for keyword in ['库号', '编号', 'NO.', 'No']) and 
                any(char.isdigit() for char in content_item)) or \
               (len(content_item) > 10 and any(char.isdigit() for char in content_item) and 
                any(char.isalpha() for char in content_item)):  # 看起来像库号的字符串（长度>10，包含数字和字母）
                library_number = content_item
                continue
                
        # 匹配人员信息（增强鲁棒性）
        if not leader_info and ('主管领导' in content_item or ('主管' in content_item and '领导' in content_item)):
            leader_info = content_item
            continue
        elif not vp_engineer_info and ('副总工程师' in content_item or ('副总' in content_item and '工程师' in content_item)):
            vp_engineer_info = content_item
            continue
        elif not chief_designer_info and ('总设计师' in content_item or ('总设' in content_item and '计师' in content_item)):
            chief_designer_info = content_item
            continue
    
    # 根据配置键的特征进行匹配
    for key in config_keys:
        # 项目名称匹配（以"项目"结尾的配置键）
        if key.strip().endswith("项目") and project_name:
            replace_dict[key] = project_name
        # 日期匹配（包含年月相关的配置键）
        elif any(date_keyword in key for date_keyword in ['年', '月', '日']) and date_text:
            replace_dict[key] = _normalize_cover_date_text(date_text)
        # 库号匹配
        elif any(lib_keyword in key for lib_keyword in ['库号', '编号']):
            if library_number:
                # 对于库号，我们需要确保它有正确的前缀以匹配模板格式
                # 如果库号值不包含"库号："前缀，则添加前缀以匹配模板格式
                if not (library_number.startswith("库号") or library_number.startswith("编号")):
                    # 检查配置键中的前缀，使用相同的冒号类型
                    if "库号" in key:
                        # 使用配置键中的冒号类型
                        if "库号：" in key:
                            library_number = f"库号：{library_number}"
                        else:
                            library_number = f"库号:{library_number}"
                    elif "编号" in key:
                        # 使用配置键中的冒号类型
                        if "编号：" in key:
                            library_number = f"编号：{library_number}"
                        else:
                            library_number = f"编号:{library_number}"
                # 确保库号格式正确，移除多余的空格和换行符
                cleaned_library_number = library_number.strip().replace('\n', '').replace('\r', '')
                replace_dict[key] = cleaned_library_number
            # 如果没有直接匹配到库号，但配置键包含库号前缀，尝试从content_dict中查找
            elif "content_2" in content_dict:
                # 假设content_2是库号（根据extract_content_from_docx函数的逻辑）
                library_value = content_dict["content_2"]
                if library_value:
                    # 如果库号值不包含"库号："前缀，则添加前缀以匹配模板格式
                    if not (library_value.startswith("库号") or library_value.startswith("编号")):
                        # 检查配置键中的前缀，使用相同的冒号类型
                        if "库号" in key:
                            # 使用配置键中的冒号类型
                            if "库号：" in key:
                                library_value = f"库号：{library_value}"
                            else:
                                library_value = f"库号:{library_value}"
                        elif "编号" in key:
                            # 使用配置键中的冒号类型
                            if "编号：" in key:
                                library_value = f"编号：{library_value}"
                            else:
                                library_value = f"编号:{library_value}"
                    replace_dict[key] = library_value
                else:
                    # 如果content_2为空，但我们需要确保库号字段不为空
                    # 这种情况下，我们使用一个默认值或者保持原配置键的值
                    replace_dict[key] = key  # 保持原配置键的值
        # 人员信息匹配 - 按照顺序匹配
        elif key.startswith("人员_"):
            # 从content_dict中查找对应的人员信息
            # 人员信息在content_dict中的键是content_3, content_4, content_5, ...
            person_index = int(key.split("_")[1])
            content_key = f"content_{3 + person_index}"
            if content_key in content_dict:
                replace_dict[key] = content_dict[content_key]
        # 其他人员信息匹配（保持原有的职务匹配逻辑作为备选）
        elif '主管领导' in key and leader_info:
            replace_dict[key] = leader_info
        elif '副总工程师' in key and vp_engineer_info:
            replace_dict[key] = vp_engineer_info
        elif '总设计师' in key and chief_designer_info:
            replace_dict[key] = chief_designer_info
    
    return replace_dict


def replace_content_in_template_xml_to_docx(template_docx_path: str, replace_dict: Dict[str, str], output_docx_path: str, save_file: bool = False) -> str:
    """
    使用XML方式替换模板中的内容并生成DOCX文件
    与replace_content_in_template_docx方法类似，直接处理DOCX文件
    
    Args:
        template_docx_path: 模板docx文件路径
        replace_dict: key-value替换映射
        output_docx_path: 输出的DOCX文件路径
        save_file: 是否保存文件到指定路径，默认为False（不保存，使用临时文件）
        
    Returns:
        str: 实际的文件路径（临时路径或指定路径）
    """
    # 先检查是否已导入Document
    if Document is None:
        raise ImportError("未找到python-docx库，请安装：pip install python-docx")
    
    import tempfile
    import shutil
    
    # 确定实际的输出路径
    if save_file:
        # 保存到指定路径
        actual_output_path = output_docx_path
        output_dir = os.path.dirname(actual_output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # 如果输出文件已存在，先尝试删除
        if os.path.exists(actual_output_path):
            try:
                os.remove(actual_output_path)
            except Exception as e:
                # 如果删除失败，可能是文件被占用，等待一段时间再试
                import time
                time.sleep(0.1)
                try:
                    os.remove(actual_output_path)
                except Exception:
                    raise Exception(f"无法覆盖输出文件 {actual_output_path}，文件可能正在被使用")
    else:
        # 使用临时文件
        temp_dir = tempfile.mkdtemp()
        actual_output_path = os.path.join(temp_dir, os.path.basename(output_docx_path))
    
    # 使用python-docx方式处理，确保与replace_content_in_template_docx方法一致
    # 打开模板文档
    doc = Document(template_docx_path)
    
    # 替换段落中的文本，保留格式
    for paragraph in doc.paragraphs:
        # 获取段落的完整文本
        full_text = get_paragraph_text(paragraph)
        
        # 检查是否需要替换
        for key, value in replace_dict.items():
            if key in full_text and value:  # 只有当value不为空时才替换
                # 使用我们自定义的函数替换文本
                replace_text_in_runs(paragraph.runs, key, value)
    
    # 替换表格中的文本，保留格式
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    # 获取段落的完整文本
                    full_text = get_paragraph_text(paragraph)
                    
                    # 检查是否需要替换
                    for key, value in replace_dict.items():
                        if key in full_text and value:  # 只有当value不为空时才替换
                            # 使用我们自定义的函数替换文本
                            replace_text_in_runs(paragraph.runs, key, value)
    
    # 保存文档
    clear_run_spacing(doc)
    doc.save(actual_output_path)
    
    # 如果不保存文件，返回临时路径（程序结束后会自动清理）
    # 如果保存文件，返回指定路径
    return actual_output_path


def replace_content_in_template_docx(template_docx_path: str, replace_dict: Dict[str, str], output_docx_path: str, save_file: bool = False) -> str:
    """
    使用python-docx方式替换模板中的内容，保留原有格式
    
    Args:
        template_docx_path: 模板docx文件路径
        replace_dict: key-value替换映射
        output_docx_path: 输出docx文件路径
        save_file: 是否保存文件到指定路径，默认为False（不保存，使用临时文件）
        
    Returns:
        str: 实际的文件路径（临时路径或指定路径）
    """
    # 先检查是否已导入Document
    if Document is None:
        raise ImportError("未找到python-docx库，请安装：pip install python-docx")
    
    import tempfile
    import shutil
    
    # 确定实际的输出路径
    if save_file:
        # 保存到指定路径
        actual_output_path = output_docx_path
        output_dir = os.path.dirname(actual_output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # 如果输出文件已存在，先尝试删除
        if os.path.exists(actual_output_path):
            try:
                os.remove(actual_output_path)
            except Exception as e:
                # 如果删除失败，可能是文件被占用，等待一段时间再试
                import time
                time.sleep(0.1)
                try:
                    os.remove(actual_output_path)
                except Exception:
                    raise Exception(f"无法覆盖输出文件 {actual_output_path}，文件可能正在被使用")
    else:
        # 使用临时文件
        temp_dir = tempfile.mkdtemp()
        actual_output_path = os.path.join(temp_dir, os.path.basename(output_docx_path))
    
    # 打开模板文档
    doc = Document(template_docx_path)
    
    # 替换段落中的文本，保留格式
    for paragraph in doc.paragraphs:
        # 获取段落的完整文本
        full_text = get_paragraph_text(paragraph)
        
        # 检查是否需要替换
        for key, value in replace_dict.items():
            if key in full_text and value:  # 只有当value不为空时才替换
                # 使用我们自定义的函数替换文本
                replace_text_in_runs(paragraph.runs, key, value)
    
    # 替换表格中的文本，保留格式
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    # 获取段落的完整文本
                    full_text = get_paragraph_text(paragraph)
                    
                    # 检查是否需要替换
                    for key, value in replace_dict.items():
                        if key in full_text and value:  # 只有当value不为空时才替换
                            # 使用我们自定义的函数替换文本
                            replace_text_in_runs(paragraph.runs, key, value)
    
    # 保存文档
    clear_run_spacing(doc)
    doc.save(actual_output_path)
    
    # 返回实际的文件路径
    return actual_output_path


def replace_content_in_cover(config_path: str, source_docx_path: str, template_docx_path: str, output_docx_path: str, save_file: bool = False) -> str:
    """
    快捷函数，从给定文件中识别配置key对应的value并替换模板内容
    
    Args:
        config_path: 配置文件路径
        source_docx_path: 源DOCX文件路径（从中提取内容）
        template_docx_path: 模板DOCX文件路径
        output_docx_path: 输出DOCX文件路径
        save_file: 是否保存文件到指定路径，默认为False（不保存，使用临时文件）
        
    Returns:
        str: 实际的文件路径（临时路径或指定路径）
    """
    # 1. 加载配置文件
    replace_dict = load_replace_config(config_path)
    print(f"加载配置文件: {config_path}")
    
    # 2. 从源DOCX文件中提取内容
    if os.path.exists(source_docx_path):
        print(f"从源文件提取内容: {source_docx_path}")
        content_dict = extract_content_from_docx(source_docx_path)
    else:
        print(f"源文件不存在，使用模拟数据: {source_docx_path}")
        # 如果源文件不存在，使用模拟数据
        content_dict = {
            "content_0": "高性能、安全、环保的磷酸亚铁锂电池材料项目",
            "content_1": "2024年6月15日",
            "content_2": "库号：20032357ARPD01AK009",
            "content_3": "公司主管领导：王 小 华",
            "content_4": "副总工程师  ：赵 小 四",
            "content_5": "总设计师 ：刘 小 五"
        }
    
    # 3. 将提取的内容匹配到配置键
    config_keys = list(replace_dict.keys())
    updated_replace_dict = match_content_to_config(content_dict, config_keys)
    
    # 4. 打印匹配结果
    print("\n=== 填入value后的key-value结构 ===")
    for key, value in updated_replace_dict.items():
        print(f"'{key}': '{value}'")
    
    # 5. 使用匹配后的字典替换模板内容
    print(f"\n开始替换模板内容: {template_docx_path}")
    actual_path = replace_content_in_template_docx(
        template_docx_path, 
        updated_replace_dict, 
        output_docx_path, 
        save_file
    )
    
    # print(f"替换完成，输出路径: {actual_path}")
    return actual_path


def replace_content_in_cover_auto(source_docx_path: str, output_docx_path: str, document_type: int, save_file: bool = False) -> str:
    """
    自动选择模板和配置文件的快捷函数，从给定文件中识别配置key对应的value并替换模板内容
    
    Args:
        source_docx_path: 源DOCX文件路径（从中提取内容）
        output_docx_path: 输出DOCX文件路径
        document_type: 文档类型 (1, 2, 3, 4)
        save_file: 是否保存文件到指定路径，默认为False（不保存，使用临时文件）
        
    Returns:
        str: 实际的文件路径（临时路径或指定路径）
    """
    # 获取项目根目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    parent_dir = os.path.dirname(current_dir)
    template_dir = os.path.join(parent_dir, 'template')
    
    # 1. 分析源文档中的项目字数和日期类型
    project_length, date_type = analyze_cover_content(source_docx_path)
    print(f"分析源文档结果: 项目字数={project_length}, 日期类型={date_type}")
    
    # 2. 根据项目字数确定字数参数
    if project_length <= 16:
        length_param = 1
    elif project_length <= 32:
        length_param = 2
    else:
        length_param = 3
    
    # 3. 确定日期参数
    date_param = date_type  # 1=年月日型, 2=年月型
    
    # 4. 确定文档类型参数
    doc_type_param = document_type  # 1, 2, 3, 4
    
    # 5. 构建配置文件路径
    config_filename = f"cover_replace_config_{date_param}_{length_param}.json"
    config_path = os.path.join(template_dir, config_filename)
    
    # 6. 构建模板文件路径
    template_filename = f"reference_cover{doc_type_param}_{date_param}_{length_param}.docx"
    template_path = os.path.join(template_dir, template_filename)
    
    # 7. 检查文件是否存在
    if not os.path.exists(config_path):
        print(f"警告: 配置文件不存在: {config_path}")
        # 尝试使用默认配置文件
        config_path = os.path.join(template_dir, 'cover_replace_config.json')
        if not os.path.exists(config_path):
            raise FileNotFoundError(f"配置文件不存在: {config_path}")
    
    if not os.path.exists(template_path):
        print(f"警告: 模板文件不存在: {template_path}")
        # 尝试使用默认模板文件
        template_path = os.path.join(template_dir, 'reference_cover1.docx')
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"模板文件不存在: {template_path}")
    
    print(f"选择的配置文件: {config_path}")
    print(f"选择的模板文件: {template_path}")
    
    # 8. 加载配置文件
    replace_dict = load_replace_config(config_path)
    print(f"加载配置文件: {config_path}")
    
    # 9. 从源DOCX文件中提取内容
    if os.path.exists(source_docx_path):
        print(f"从源文件提取内容: {source_docx_path}")
        content_dict = extract_content_from_docx(source_docx_path)
    else:
        print(f"源文件不存在，使用模拟数据: {source_docx_path}")
        # 如果源文件不存在，使用模拟数据
        content_dict = {
            "content_0": "高性能、安全、环保的磷酸亚铁锂电池材料项目",
            "content_1": "2024年6月15日",
            "content_2": "库号：20032357ARPD01AK009",
            "content_3": "公司主管领导：王 小 华",
            "content_4": "副总工程师  ：赵 小 四",
            "content_5": "总设计师 ：刘 小 五"
        }
    
    # 10. 将提取的内容匹配到配置键
    config_keys = list(replace_dict.keys())
    updated_replace_dict = match_content_to_config(content_dict, config_keys)
    
    # 11. 打印匹配结果
    print("\n=== 填入value后的key-value结构 ===")
    for key, value in updated_replace_dict.items():
        print(f"'{key}': '{value}'")
    
    # 12. 使用匹配后的字典替换模板内容
    print(f"\n开始替换模板内容: {template_path}")
    actual_path = replace_content_in_template_docx(
        template_path, 
        updated_replace_dict, 
        output_docx_path, 
        save_file
    )
    
    # print(f"替换完成，输出路径: {actual_path}")
    return actual_path


def replace_content_in_cover_by_dify(source_docx_path: str, output_docx_path: str, document_type: int, save_file: bool = False) -> Tuple[str, str, str]:
    """
    使用Dify服务自动获取项目信息并替换封面内容
    
    Args:
        source_docx_path: 源DOCX文件路径（用于上传到Dify服务）
        output_docx_path: 输出DOCX文件路径
        document_type: 文档类型 (1, 2, 3, 4)
        save_file: 是否保存文件到指定路径，默认为False（不保存，使用临时文件）
        
    Returns:
        str: 实际的文件路径（临时路径或指定路径）
    """

    def config2dify(config_dict, project_dict):
        """
        以 config 为 key，dify 为 value
        人员使用"职务：姓名"方式映射，多余的 dify 人员以空键 "" 加入

        输出类似：
        {
            "高性能、安全、环保的磷酸亚铁锂电池材料项目": "鞍钢股份炼焦总厂1#、2#焦炉烟气脱硝脱硝扩容改造项目",
            "2024年4月": "2025年6月",
            "10032357ARPD01AK008": "10180286CDYT01EP001",
            "公司主管领导：张 小 明": "公司主管领导：王寅虎",
            "副总工程师  ：李 小 二": "事业部经理：齐  猛",
            "总设计师 ：陈 小 红": "总设计师：王维浩",
            "": "清洁部经理：张小二"
        }

        """
        project_date = _normalize_cover_date_text(project_dict.get("时间", ""))
        replacer = {
            config_dict["项目名"]: project_dict["项目名"],
            config_dict["时间"]: project_date,
            config_dict["库号"]: project_dict["库号"],
        }

        # 取出 config 中的人员信息（按序号）
        config_person_keys = sorted(
            [k for k in config_dict.keys() if k.startswith("人员_")],
            key=lambda x: int(x.split("_")[1])
        )
        config_person_texts = [config_dict[k] for k in config_person_keys]
        print(f"人员信息：{config_person_texts}")
        
        # 打印调试信息
        print("=== 调试信息 ===")
        for i, text in enumerate(config_person_texts):
            print(f"  人员_{i}: '{text}' (长度: {len(text)})")
        
        # 转换为 (职务, 姓名) 二元组
        def parse_person(text):
            # 不再去除空格，保留原始格式
            parts = text.split("：", 1)
            if len(parts) == 2:
                # 保留职务和姓名中的空格，但需要标准化冒号周围的空格
                role = parts[0].rstrip()  # 去除职务右侧的空格
                name = parts[1]
                return role, name
            else:
                return text, ""
        config_people = [parse_person(t) for t in config_person_texts]
        dify_people = project_dict["人员"]

        # 一一对应
        for i, (cfg_role, cfg_name) in enumerate(config_people):
            # 使用原始文本作为键，确保与文档中的内容完全匹配
            original_text = config_person_texts[i]
            if i < len(dify_people):
                dify_role = dify_people[i]["职务"].strip()
                dify_name = dify_people[i]["姓名"].strip()
                replacer[original_text] = f"{dify_role}：{dify_name}"
            else:
                replacer[original_text] = ""

        # 若 Dify 人员更多，添加空键 ""
        if len(dify_people) > len(config_people):
            for j in range(len(config_people), len(dify_people)):
                p = dify_people[j]
                replacer[""] = f"{p['职务']}：{p['姓名']}"

        return replacer

    def replacements2docx(template_docx_path: str, output_docx_path: str, replacements: dict, personnel_list: list, save_file: bool = False) -> str:
        """
        使用python-docx方式替换模板中的内容
        
        Args:
            template_docx_path: 模板docx文件路径
            output_docx_path: 输出docx文件路径
            replacements: 替换映射字典
            personnel_list: 人员列表
            save_file: 是否保存文件到指定路径，默认为False（不保存，使用临时文件）
            
        Returns:
            str: 实际的文件路径（临时路径或指定路径）
        """
        # 确定实际的输出路径
        if save_file:
            # 保存到指定路径
            actual_output_path = output_docx_path
            output_dir = os.path.dirname(actual_output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # 如果输出文件已存在，先尝试删除
            if os.path.exists(actual_output_path):
                try:
                    os.remove(actual_output_path)
                except Exception as e:
                    # 如果删除失败，可能是文件被占用，等待一段时间再试
                    import time
                    time.sleep(0.1)
                    try:
                        os.remove(actual_output_path)
                    except Exception:
                        raise Exception(f"无法覆盖输出文件 {actual_output_path}，文件可能正在被使用")
        else:
            # 使用临时文件
            temp_dir = tempfile.mkdtemp()
            actual_output_path = os.path.join(temp_dir, os.path.basename(output_docx_path))
        
        # 打开模板文档
        doc = Document(template_docx_path)
        
        # 过滤掉空键的替换项，只保留有效的替换项
        valid_replacements = {k: v for k, v in replacements.items() if k and v}
        
        # 收集需要添加的额外人员信息
        extra_personnel = [v for k, v in replacements.items() if k == "" and v]
        
        # 替换段落中的文本，保留格式
        for paragraph in doc.paragraphs:
            # 获取段落的完整文本
            full_text = get_paragraph_text(paragraph)
            
            # 检查是否需要替换
            for key, value in valid_replacements.items():
                if key in full_text and value:  # 只有当value不为空时才替换
                    # 使用我们自定义的函数替换文本
                    replace_text_in_runs(paragraph.runs, key, value)
        
        # 如果有额外的人员信息需要添加
        if extra_personnel:
            # 查找人员信息段落
            person_paragraphs = []
            for i, paragraph in enumerate(doc.paragraphs):
                full_text = get_paragraph_text(paragraph)
                # 检查是否包含人员信息（以冒号结尾的职务信息）
                if any(role in full_text for role in ["主管领导", "副总工程师", "总设计师", "经理", "主任", "职务"]):
                    person_paragraphs.append((i, paragraph))
            
            # 如果找到了人员信息段落，在最后一个人员信息后添加额外的人员信息
            if person_paragraphs:
                # 获取最后一个人员信息段落的索引
                last_person_index = person_paragraphs[-1][0]
                
                # 收集要插入的段落内容
                paragraphs_to_insert = list(extra_personnel)
                
                # 从后往前插入，这样索引不会变化
                for i in range(len(paragraphs_to_insert)):
                    extra_person = paragraphs_to_insert[len(paragraphs_to_insert) - 1 - i]
                    # 在最后一个人员段落之后插入新段落
                    insert_index = last_person_index + 1 + (len(paragraphs_to_insert) - 1 - i)
                    
                    # 确保索引不超出范围
                    if insert_index <= len(doc.paragraphs):
                        # 插入新段落
                        if insert_index < len(doc.paragraphs):
                            new_paragraph = doc.paragraphs[insert_index].insert_paragraph_before(extra_person)
                        else:
                            new_paragraph = doc.add_paragraph(extra_person)
                        
                        # 复制格式
                        last_person_paragraph = person_paragraphs[-1][1]
                        if last_person_paragraph.runs:
                            if new_paragraph.runs:
                                new_run = new_paragraph.runs[0]
                                old_run = last_person_paragraph.runs[0]
                                # 复制字体格式
                                new_run.font.name = old_run.font.name
                                new_run.font.size = old_run.font.size
                                new_run.font.bold = old_run.font.bold
                                new_run.font.italic = old_run.font.italic
            else:
                # 如果没有找到人员信息段落，就在文档末尾添加
                for extra_person in extra_personnel:
                    doc.add_paragraph(extra_person)
        # 保存文档
        clear_run_spacing(doc)
        apply_character_spacing_to_cover(doc)
        doc.save(actual_output_path)
        
        # 返回实际的文件路径
        return actual_output_path
    
    # 获取项目根目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    parent_dir = os.path.dirname(current_dir)
    template_dir = os.path.join(parent_dir, 'template')
    
    # 初始化Dify服务（使用默认API密钥）
    # dify_service = DifyService(api_key="app-UjqMjco0NjWwN4GbWHuNjnKg")
    dify_service = DifyService(api_key="app-x6Yx0iAI3aUFMpGS9fpJQowq", base_url="http://172.23.109.63:8088/v1")
    
    try:
        # 从Dify获取项目信息
        project_info_json = dify_service.extract_cover(source_docx_path)
        try:
            project_info = _parse_first_json(project_info_json)
        except json.JSONDecodeError as e:
            snippet = str(project_info_json)
            print(f"Dify返回内容无法解析为JSON: {e}")
            print(f"Dify返回内容片段: {snippet[:300]}")
            raise
        
        print(f"从Dify获取到的项目信息: {project_info}")
        
        # 检查是否有错误
        if project_info.get("status") == "error":
            raise Exception(f"Dify服务返回错误: {project_info.get('message')}")
        
        # 提取关键信息
        project_name = project_info.get("项目名", "")
        document_type_info = project_info.get("文本类型", "")
        date_info = _normalize_cover_date_text(project_info.get("时间", ""))
        project_info["时间"] = date_info
        library_number = project_info.get("库号", "")
        personnel_list = project_info.get("人员", [])
        
        # 分析项目名称长度和日期类型
        project_length = len(project_name) if project_name else 0
        project_lines = [line.strip() for line in project_name.splitlines() if line.strip()]
        project_line_count = len(project_lines)
        
        # 确定日期参数 (简化处理，根据日期字符串中是否包含日来判断)
        date_param = 1 if date_info and ('日' in date_info or '-' in date_info or '/' in date_info) else 2
        
        # 确定字数参数（优先按实际行数选择模板）
        if project_line_count >= 3:
            length_param = 3
        elif project_line_count == 2:
            length_param = 2
        elif project_length <= 16:
            length_param = 1
        elif project_length <= 32:
            length_param = 2
        else:
            length_param = 3
            
        # 确定文档类型参数
        doc_type_param = document_type  # 1, 2, 3, 4
        
        # 构建配置文件路径
        config_filename = f"cover_replace_config_{date_param}_{length_param}.json"
        config_path = os.path.join(template_dir, config_filename)
        
        # 构建模板文件路径
        template_filename = f"reference_cover{doc_type_param}_{date_param}_{length_param}.docx"
        template_path = os.path.join(template_dir, template_filename)
        
        # 检查文件是否存在，如果不存在则使用默认文件
        if not os.path.exists(config_path):
            print(f"警告: 配置文件不存在: {config_path}")
            # 尝试使用默认配置文件
            config_path = os.path.join(template_dir, 'cover_replace_config.json')
            if not os.path.exists(config_path):
                raise FileNotFoundError(f"配置文件不存在: {config_path}")
        
        if not os.path.exists(template_path):
            print(f"警告: 模板文件不存在: {template_path}")
            # 尝试使用默认模板文件
            template_path = os.path.join(template_dir, 'reference_cover1.docx')
            if not os.path.exists(template_path):
                raise FileNotFoundError(f"模板文件不存在: {template_path}")
        
        print(f"选择的配置文件: {config_path}")
        print(f"选择的模板文件: {template_path}")
        
        # 加载配置文件
        config_dict = load_replace_config(config_path)
        print(f"加载配置文件：{config_dict}")
        
        # 使用config2dify生成替换映射
        replacements = config2dify(config_dict, project_info)
        print(f"替换映射为:{replacements}")
        
        # 打印人员信息详情用于调试
        print("=== 人员信息详情 ===")
        config_person_keys = sorted(
            [k for k in config_dict.keys() if k.startswith("人员_")],
            key=lambda x: int(x.split("_")[1])
        )
        config_person_texts = [config_dict[k] for k in config_person_keys]
        print(f"配置文件中的人员信息: {config_person_texts}")
        
        dify_people = project_info.get("人员", [])
        print(f"Dify返回的人员信息: {dify_people}")
        
        # 直接使用config2dify返回的替换字典，无需额外处理
        updated_replace_dict = replacements
        
        # 打印匹配结果
        print("构建替换字典:")
        print(updated_replace_dict)
        print("\n=== 填入value后的key-value结构 ===")
        for key, value in updated_replace_dict.items():
            print(f"'{key}': '{value}'")
        
        # 使用replacements2docx函数替换模板内容
        print(f"\n开始替换模板内容: {template_path}")
        
        # 使用新的replacements2docx函数替换内容
        actual_output_path = replacements2docx(template_path, output_docx_path, replacements, personnel_list, save_file)

        return actual_output_path, project_name, document_type_info
        
    except Exception as e:
        print(f"使用Dify处理封面时出错: {e}")
        raise


def create_replacement_dict_from_content(content_dict: Dict[str, str], config_keys: list) -> Dict[str, str]:
    """
    根据提取的内容和配置的key创建替换字典
    
    Args:
        content_dict: 从文档中提取的内容
        config_keys: 配置文件中的key列表
        
    Returns:
        Dict[str, str]: 完整的替换字典
    """
    replace_dict = {key: "" for key in config_keys}
    
    # 这里需要实现具体的匹配逻辑，目前是简化处理
    # 实际应用中可能需要更复杂的匹配规则
    content_values = list(content_dict.values())
    
    # 简单的一一对应匹配
    for i, key in enumerate(config_keys):
        if i < len(content_values):
            replace_dict[key] = content_values[i]
    
    return replace_dict


def _parse_first_json(text: str) -> Dict[str, Any]:
    """
    解析字符串中的第一个JSON对象/数组，避免出现Extra data错误。
    """
    if isinstance(text, dict):
        return text
    if text is None:
        raise json.JSONDecodeError("Empty input", "", 0)
    s = str(text).strip()
    decoder = json.JSONDecoder()
    try:
        obj, _ = decoder.raw_decode(s)
        return obj
    except json.JSONDecodeError:
        # 尝试截取首个JSON对象或数组
        start_obj = s.find('{')
        start_arr = s.find('[')
        if start_obj == -1 and start_arr == -1:
            raise
        if start_arr != -1 and (start_obj == -1 or start_arr < start_obj):
            start = start_arr
        else:
            start = start_obj
        if start != -1:
            obj, _ = decoder.raw_decode(s[start:])
            return obj
        raise


def extract_text_from_xml(docx_path: str) -> str:
    """
    从docx文件的XML中提取所有文本内容
    
    Args:
        docx_path: docx文件路径
        
    Returns:
        str: 提取的文本内容
    """
    # 创建临时目录
    with tempfile.TemporaryDirectory() as temp_dir:
        # 解压docx文件
        with zipfile.ZipFile(docx_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        # 读取document.xml
        document_xml_path = os.path.join(temp_dir, 'word', 'document.xml')
        if os.path.exists(document_xml_path):
            tree = ET.parse(document_xml_path)
            root = tree.getroot()
            
            # 定义命名空间
            namespaces = {
                'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
            }
            
            # 提取所有文本
            text_elements = root.findall('.//w:t', namespaces)
            texts = [elem.text for elem in text_elements if elem.text]
            return ' '.join(texts)
    
    return ""


def analyze_cover_content(docx_path: str) -> tuple:
    """
    分析封面文件中的项目名称字数和日期形式
    
    Args:
        docx_path: DOCX文件路径
        
    Returns:
        tuple: (项目名称字数, 日期形式) 
               日期形式: 1表示年月日型(如2020年2月12日), 2表示年月型(如2020年2月), 0表示未识别到日期
    """
    # 先检查是否已导入Document
    if Document is None:
        raise ImportError("未找到python-docx库，请安装：pip install python-docx")
    
    try:
        doc = Document(docx_path)
        project_name_length = 0
        date_type = 0  # 0:未识别, 1:年月日型, 2:年月型
        
        # 提取所有文本内容
        all_text = []
        
        # 提取段落内容
        for paragraph in doc.paragraphs:
            if paragraph.text.strip():
                all_text.append(paragraph.text.strip())
        
        # 提取表格内容
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if cell.text.strip():
                        all_text.append(cell.text.strip())
        
        # 初始化匹配变量
        project_name = ""
        date_text = ""
        
        # 分析文本内容（参考match_content_to_config函数的逻辑）
        for content_item in all_text:
            # 跳过目录等无关内容
            if content_item in ["目录", "第一章 摘要      1", "第二章  背景概述 3", "§2.1  政策背景   3", 
                               "§2.2  经济背景   4", "§2.3  产业背景   6", "第三章  电池市场分析     9"]:
                continue
                
            # 匹配项目名称（以"项目"结尾）
            if not project_name and content_item.strip().endswith("项目"):
                # 进一步验证不是日期或其他信息
                if not any(date_keyword in content_item for date_keyword in ['年', '月', '日']) or \
                   content_item.count('年') < 2:  # 避免将日期误识别为项目名
                    project_name = content_item
                    project_name_length = len(project_name.strip())
            
            # 匹配日期（支持年月日和年月格式），先去空格后判断
            if not date_text:
                normalized_content = _normalize_cover_date_text(content_item)
                has_digit = any(char.isdigit() for char in normalized_content)
                if has_digit:
                    matched_date = _extract_normalized_cover_date(normalized_content)
                    if matched_date:
                        # 优先选择更具体的日期（包含年份的）
                        if '202' in matched_date or '201' in matched_date:  # 2020年代或2010年代
                            date_text = matched_date
                        elif not date_text:  # 如果还没有找到日期，则使用第一个匹配的
                            date_text = matched_date
            
            # 确定日期类型
            if date_text and date_type == 0:
                normalized_date_text = _normalize_cover_date_text(date_text)
                # 检查是否匹配年月日格式
                if COVER_DATE_YMD_PATTERN.search(normalized_date_text):
                    date_type = 1
                # 检查是否匹配年月格式
                elif COVER_DATE_YM_PATTERN.search(normalized_date_text):
                    date_type = 2
        
        return (project_name_length, date_type)
    
    except Exception as e:
        print(f"分析封面内容时出错: {e}")
        return (0, 0)


# 更新主程序入口以移除测试函数调用
if __name__ == "__main__":
    # 配置文件路径
    config_path = os.path.join('template', 'cover_replace_config.json')
    
    # 模板路径
    template_docx_path = os.path.join('template', 'reference_cover1.docx')
    
    # 输出路径
    output_docx_path = os.path.join('result', 'output_cover.docx')
    
    # 确保输出目录存在
    os.makedirs(os.path.dirname(output_docx_path), exist_ok=True)
    
    # 从指定的DOCX文件中提取内容并替换模板内容
    print("=== 从指定DOCX文件中提取内容并替换模板内容 ===")
    source_docx_path = r"C:\Users\yanha\Desktop\数字总师\WordStyler\docx\test.docx"
    
    try:
        # 使用快捷函数处理
        actual_path = replace_content_in_cover(
            config_path=config_path,
            source_docx_path=source_docx_path,
            template_docx_path=template_docx_path,
            output_docx_path=output_docx_path,
            save_file=False  # 保存到指定路径
        )
        print(f"处理完成，文件路径: {actual_path}")
    except Exception as e:
        print(f"处理出错: {e}")
    
    print("\n=== 处理完成 ===")
