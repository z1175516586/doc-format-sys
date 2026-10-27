from lxml import etree
import zipfile
import os
import shutil

def remove_all_footers_lxml(docx_path: str):
    print("===== 进入 remove_all_footers_lxml 函数 ======")

    if not os.path.isfile(docx_path):
        print(f"[错误] 文件不存在: {docx_path}")
        return

    temp_dir = docx_path + "_temp_extracted"
    os.makedirs(temp_dir, exist_ok=True)

    try:
        with zipfile.ZipFile(docx_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        print("[日志] DOCX 解压完成")

        word_dir = os.path.join(temp_dir, "word")
        footer_files = [f for f in os.listdir(word_dir) if f.startswith("footer") and f.endswith(".xml")]

        for footer_file in footer_files:
            footer_path = os.path.join(word_dir, footer_file)
            parser = etree.XMLParser(remove_blank_text=True)
            tree = etree.parse(footer_path, parser)
            root = tree.getroot()

            # 删除所有子元素，但保留根和命名空间
            for child in list(root):
                root.remove(child)

            tree.write(footer_path, encoding="UTF-8", xml_declaration=True, pretty_print=True)
            print(f"[日志] 已清空页脚内容: {footer_file}")

        with zipfile.ZipFile(docx_path, 'w', zipfile.ZIP_DEFLATED) as zip_write:
            for foldername, subfolders, filenames in os.walk(temp_dir):
                for filename in filenames:
                    file_path = os.path.join(foldername, filename)
                    arcname = os.path.relpath(file_path, temp_dir)
                    zip_write.write(file_path, arcname)
        print("[日志] DOCX 文件已原地修改完成")

    finally:
        shutil.rmtree(temp_dir)
        print(f"[日志] 临时文件夹已删除: {temp_dir}")

    print("===== 退出 remove_all_footers_lxml 函数 ======")

if __name__ == "__main__":
    remove_all_footers_lxml(r"D:\ASE_develop_file\app\tmp8kjekq03.docx")
