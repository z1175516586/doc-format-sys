"""
DOCX图片格式处理工具
提供修改正文中图片格式的功能，如居中、单倍行距等，同时不修改页眉上的logo
支持通过配置文件设置图片格式，方便用户自定义
"""

import json
import os
import sys
import importlib
from pathlib import Path

from winerror import NOERROR

# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.dirname(__file__)))


def load_picture_settings(config_path="config/picture_settings.json"):
    """
    加载图片格式配置
    
    Args:
        config_path (str): 配置文件路径
        
    Returns:
        dict: 图片格式配置字典
    """
    try:
        # 获取项目根目录
        project_root = Path(__file__).parent.parent
        config_file_path = project_root / config_path
        
        if not config_file_path.exists():
            print(f"⚠️ 配置文件 {config_file_path} 不存在，使用默认配置")
            return get_default_picture_settings()
        
        with open(config_file_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        
        return config.get("picture_format", get_default_picture_settings())
    
    except Exception as e:
        print(f"❌ 加载图片配置失败: {e}")
        return get_default_picture_settings()


def get_default_picture_settings():
    """
    获取默认图片格式配置
    
    Returns:
        dict: 默认图片格式配置
    """
    return {
        "alignment": "center",  # 对齐方式: left, center, right, justify
        "line_spacing": 1.0,    # 行距: 1.0为单倍行距
        "before_spacing": 0,    # 段前间距
        "after_spacing": 0,     # 段后间距
        "keep_with_next": False, # 与下段同页
        "keep_lines": False,    # 段中不分页
        "picture_width": None,  # 图片宽度 (单位: 英寸)
        "picture_height": None, # 图片高度 (单位: 英寸)
        "wrap_type": "inline",  # 环绕方式: inline, topAndBottom, square, tight等
        "skip_cover": False     # 跳过封面区域图片格式处理
    }


def format_pictures_in_document(doc_path, save_path=None, config_path="config/picture_settings.json"):
    """
    修改DOCX文档中正文图片的格式，不修改页眉中的logo
    
    Args:
        doc_path (str): 输入的DOCX文件路径
        save_path (str, optional): 保存路径，默认覆盖原文件
        config_path (str): 图片格式配置文件路径
        
    Returns:
        bool: 操作是否成功
    """
    try:
        # 动态导入docx模块
        docx = importlib.import_module('docx')
        Document = docx.Document
        WD_ALIGN_PARAGRAPH = importlib.import_module('docx.enum.text').WD_ALIGN_PARAGRAPH
        
        # 加载配置
        picture_settings = load_picture_settings(config_path)
        print(f"🔧 使用图片格式配置: {picture_settings}")
        
        # 打开文档
        doc = Document(doc_path)
        print(f"📄 成功加载文档: {doc_path}")
        
        # 获取对齐方式枚举值
        alignment_map = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY
        }
        alignment = alignment_map.get(picture_settings["alignment"], WD_ALIGN_PARAGRAPH.CENTER)
        
        # 获取行距设置
        line_spacing = picture_settings["line_spacing"]
        before_spacing = picture_settings["before_spacing"]
        after_spacing = picture_settings["after_spacing"]
        keep_with_next = picture_settings["keep_with_next"]
        keep_lines = picture_settings["keep_lines"]
        picture_width = picture_settings["picture_width"]
        picture_height = picture_settings["picture_height"]
        wrap_type = picture_settings["wrap_type"]
        
        # 处理正文中的图片（不处理页眉页脚中的图片）
        formatted_count = 0
        
        # 遍历文档中的所有段落
        for paragraph in doc.paragraphs:
            # 检查段落中是否包含图片
            if paragraph.runs:
                for run in paragraph.runs:
                    # 检查run中是否有图片
                    if run._element.xpath('.//w:drawing') or run._element.xpath('.//w:pict'):
                        # 设置段落格式
                        paragraph.alignment = alignment
                        
                        # 设置段落行距和间距
                        paragraph_format = paragraph.paragraph_format
                        paragraph_format.line_spacing = line_spacing
                        paragraph_format.space_before = before_spacing
                        paragraph_format.space_after = after_spacing
                        paragraph_format.keep_with_next = keep_with_next
                        paragraph_format.keep_together = keep_lines
                        
                        formatted_count += 1
                        print(f"✅ 已格式化段落中的图片，当前段落对齐方式: {picture_settings['alignment']}")
        
        # 保存文档
        output_path = save_path if save_path else doc_path
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
        
        doc.save(output_path)
        print(f"💾 文档已保存到: {output_path}")
        print(f"🎉 成功格式化了 {formatted_count} 个包含图片的段落")
        
        return True
        
    except Exception as e:
        print(f"❌ 处理文档时发生错误: {e}")
        import traceback
        traceback.print_exc()
        return False


def format_pictures_with_advanced_settings(doc_path, formula_info_list: list | None = None, save_path=None, config_path="config/picture_settings.json", skip_cover=False):
    """
    使用高级设置修改DOCX文档中正文图片的格式
    
    Args:
        doc_path (str): 输入的DOCX文件路径
        save_path (str, optional): 保存路径，默认覆盖原文件
        config_path (str): 图片格式配置文件路径
        skip_cover (bool): 是否跳过封面区域的图片段落
        
    Returns:
        bool: 操作是否成功
    """
    try:
        # 动态导入docx模块
        docx = importlib.import_module('docx')
        Document = docx.Document
        WD_ALIGN_PARAGRAPH = importlib.import_module('docx.enum.text').WD_ALIGN_PARAGRAPH
        
        # 加载配置
        picture_settings = load_picture_settings(config_path)
        print(f"🔧 使用图片格式配置: {picture_settings}")
        
        # 打开文档
        doc = Document(doc_path)
        print(f"📄 成功加载文档: {doc_path}")
        
        # 获取对齐方式
        alignment_map = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY
        }
        alignment = alignment_map.get(picture_settings["alignment"], WD_ALIGN_PARAGRAPH.CENTER)
        
        # 获取行距设置
        line_spacing = picture_settings["line_spacing"]
        before_spacing = picture_settings["before_spacing"]
        after_spacing = picture_settings["after_spacing"]
        keep_with_next = picture_settings["keep_with_next"]
        keep_lines = picture_settings["keep_lines"]
        picture_width = picture_settings["picture_width"]
        picture_height = picture_settings["picture_height"]
        wrap_type = picture_settings["wrap_type"]
        
        # 处理正文中的图片
        formatted_count = 0
        skip_cover = bool(skip_cover or picture_settings.get("skip_cover", False))
        if skip_cover:
            has_section_break = any(p._element.xpath('.//w:sectPr') for p in doc.paragraphs)
            if not has_section_break:
                skip_cover = False
        cover_ended = not skip_cover

        # 遍历文档中的所有段落
        for paragraph in doc.paragraphs:
            if skip_cover and not cover_ended:
                if paragraph._element.xpath('.//w:sectPr'):
                    cover_ended = True
                continue
            # 检查段落中是否包含图片
            if contains_picture(paragraph):
                # 设置段落格式
                set_paragraph_format(paragraph, alignment, line_spacing, before_spacing,
                                   after_spacing, keep_with_next, keep_lines)

                formatted_count += 1
                print(f"✅ 已格式化段落中的图片，当前段落对齐方式: {picture_settings['alignment']}")

        # ========= 新增：恢复公式图片原始尺寸 =========
        restore_formula_sizes_by_file(doc, formula_info_list)

        # 保存文档
        output_path = save_path if save_path else doc_path
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
        
        doc.save(output_path)
        print(f"💾 文档已保存到: {output_path}")
        print(f"🎉 成功格式化了 {formatted_count} 个包含图片的段落")
        
        return True
        
    except Exception as e:
        print(f"❌ 处理文档时发生错误: {e}")
        import traceback
        traceback.print_exc()
        return False


from docx import Document
from docx.shared import Cm

def px_to_cm(px, dpi=96):
    return px / dpi * 2.54


def restore_formula_sizes_by_file(doc, formula_info_list):
    """
    恢复公式图片原始尺寸

    Args:
        doc: docx.Document 对象
        formula_info_list: List[Dict]，每个字典包含 'file', 'width', 'height'
    """
    from docx.shared import Cm

    def px_to_cm(px, dpi=96):
        return px / dpi * 2.54
    print("⭐开始恢复公式图片原尺寸")
    success_count = 0
    for shape in doc.inline_shapes:
        try:
            shape_file = shape._inline.graphic.graphicData.pic.nvPicPr.cNvPr.get("descr", "")
        except Exception:
            shape_file = ""
        match = next((f for f in formula_info_list if f["file"].endswith(shape_file)), None)
        if match:
            shape.width = Cm(px_to_cm(match["width"]))
            shape.height = Cm(px_to_cm(match["height"]))
            success_count += 1
        print(f"√ 公式图片 {shape_file} 恢复成功")
    print("⭐公式图片已全部恢复原尺寸")
    return success_count


def contains_picture(paragraph):
    if "库号" in (paragraph.text or ""):
        return False
    """
    检查段落是否包含图片
    
    Args:
        paragraph: docx段落对象
        
    Returns:
        bool: 是否包含图片
    """
    # 检查段落中是否有图片
    for run in paragraph.runs:
        if run._element.xpath('.//w:drawing') or run._element.xpath('.//w:pict'):
            return True
    return False


def set_paragraph_format(paragraph, alignment, line_spacing, before_spacing, 
                        after_spacing, keep_with_next, keep_lines):
    """
    设置段落格式
    
    Args:
        paragraph: docx段落对象
        alignment: 对齐方式
        line_spacing: 行距
        before_spacing: 段前间距
        after_spacing: 段后间距
        keep_with_next: 与下段同页
        keep_lines: 段中不分页
    """
    # 设置对齐方式
    paragraph.alignment = alignment
    
    # 设置段落格式
    paragraph_format = paragraph.paragraph_format
    paragraph_format.line_spacing = line_spacing
    paragraph_format.space_before = before_spacing
    paragraph_format.space_after = after_spacing
    paragraph_format.keep_with_next = keep_with_next
    paragraph_format.keep_together = keep_lines


def main():
    """
    主函数 - 提供命令行接口
    """
    print("🚀 开始执行DOCX图片格式处理脚本...")
    print("=" * 50)
    
    # 示例用法
    # format_pictures_in_document("input.docx", "output.docx")
    
    print("💡 使用方法:")
    print("   format_pictures_in_document('input.docx', 'output.docx')")
    print("   format_pictures_with_advanced_settings('input.docx', 'output.docx')")
    print("=" * 50)
    print("✅ 脚本执行完毕。")


if __name__ == "__main__":
    main()
