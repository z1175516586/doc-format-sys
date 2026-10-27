import os
import zipfile
import tempfile
import shutil
import re
from pathlib import Path


DRAWING_PATTERN = re.compile(
    r"<w:drawing[\s\S]*?</w:drawing>"
    r"|<w:pict[\s\S]*?</w:pict>"
    r"|<v:imagedata[\s\S]*?/>",
    re.MULTILINE
)



def _extract_docx(docx_path: str, extract_dir: str):
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        zip_ref.extractall(extract_dir)


def _repack_docx(folder_path: str, output_path: str):
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
        for root, _, files in os.walk(folder_path):
            for file in files:
                abs_path = os.path.join(root, file)
                rel_path = os.path.relpath(abs_path, folder_path)
                zip_ref.write(abs_path, rel_path)


def _has_drawing_in_headers_footers(extracted_path: str) -> bool:
    word_dir = Path(extracted_path) / "word"
    if not word_dir.exists():
        return False

    for xml_file in word_dir.glob("header*.xml"):
        if DRAWING_PATTERN.search(xml_file.read_text(encoding="utf-8")):
            return True

    for xml_file in word_dir.glob("footer*.xml"):
        if DRAWING_PATTERN.search(xml_file.read_text(encoding="utf-8")):
            return True

    return False


def _remove_drawing_from_headers_footers(extracted_path: str):
    word_dir = Path(extracted_path) / "word"
    if not word_dir.exists():
        return
    # 页脚别删，只处理页眉
    for xml_file in list(word_dir.glob("header*.xml")):
        content = xml_file.read_text(encoding="utf-8")
        new_content = DRAWING_PATTERN.sub("", content)
        xml_file.write_text(new_content, encoding="utf-8")


def process_docx(content_file: str, template_file: str):
    """
    如果 content_file 页眉页脚出现图片 → 不处理
    如果没有 → 删除 template_file 页眉页脚所有图片
    """

    if not os.path.exists(content_file):
        raise FileNotFoundError(f"content_file 不存在: {content_file}")

    if not os.path.exists(template_file):
        raise FileNotFoundError(f"template_file 不存在: {template_file}")

    with tempfile.TemporaryDirectory() as temp_dir:

        # 1️⃣ 检查 content_file
        content_extract = os.path.join(temp_dir, "content")
        os.makedirs(content_extract)
        _extract_docx(content_file, content_extract)

        has_drawing = _has_drawing_in_headers_footers(content_extract)

        if has_drawing:
            print("✅ content_file 检测到页眉页脚包含图片 → 不处理 template_file")
            return

        print("⚠️ content_file 未检测到页眉页脚图片 → 开始清理 template_file")

        # 2️⃣ 处理 template_file
        template_extract = os.path.join(temp_dir, "template")
        os.makedirs(template_extract)
        _extract_docx(template_file, template_extract)

        _remove_drawing_from_headers_footers(template_extract)

        # 重新打包覆盖 template_file
        backup_path = template_file + ".bak"
        shutil.copy2(template_file, backup_path)
        print(f"📦 已备份原模板: {backup_path}")

        _repack_docx(template_extract, template_file)

        print("✅ template_file 页眉页脚图片已清除")


# ======================
# 主函数（可独立测试）
# ======================

if __name__ == "__main__":
    # import sys
    #
    # if len(sys.argv) != 3:
    #     print("用法: python remove_header_image.py content.docx template.docx")
    #     sys.exit(1)
    #
    # content_path = os.path.abspath(sys.argv[1])
    # template_path = os.path.abspath(sys.argv[2])

    content_path = r"D:\ASE_develop_file\doc-convertor\test\tmpqsax7fgc_预处理标题_正文内容.docx"
    template_path = r"D:\ASE_develop_file\doc-convertor\test\reference_content.docx"

    process_docx(content_path, template_path)
