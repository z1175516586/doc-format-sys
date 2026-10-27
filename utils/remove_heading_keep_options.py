# remove_heading_keep_options.py
import os
import time
import zipfile
from lxml import etree

class DocxStyleCleaner:
    def __init__(self):

        pass
        # 固定输入文件（可根据需要修改）
        # self.input_file = r"E:\APP测试图片\word测试\test.docx"
        #
        # # 输出文件夹
        # timestamp = time.strftime("%Y%m%d_%H%M%S")
        # self.output_dir = r"E:\APP测试图片\word测试\test-ok.docx"
        # os.makedirs(self.output_dir, exist_ok=True)
        #
        # # 输出文件路径（加后缀避免覆盖原文件）
        # base_name = os.path.basename(self.input_file).replace('.docx', '')
        # self.output_file = os.path.join(
        #     self.output_dir,
        #     f"{base_name}_cleaned_{timestamp}.docx"
        # )

    def _remove_keep_options_from_heading_styles(self, docx_path: str) -> bool:
        """
        删除 docx 中所有标题样式里的 keepNext 和 keepLines 属性
        """
        import zipfile
        from lxml import etree

        W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        NSMAP = {"w": W_NS}

        try:
            # 1. 解压 docx
            with zipfile.ZipFile(docx_path, "r") as zin:
                files = {name: zin.read(name) for name in zin.namelist()}

            if "word/styles.xml" not in files:
                print("⚠️ 未找到 word/styles.xml，跳过样式清理")
                return False

            # 2. 解析 styles.xml
            styles_tree = etree.fromstring(files["word/styles.xml"])
            styles = styles_tree.xpath("//w:style", namespaces=NSMAP)

            # 标题样式 ID 集（中英文）
            heading_ids = {
                "Heading1", "Heading2", "Heading3", "Heading4", "Heading5",
                "Heading6", "Heading7", "Heading8", "Heading9",
                "标题1", "标题2", "标题3", "标题4", "标题5",
                "标题6", "标题7", "标题8", "标题9",
            }

            def is_heading_style(style_elem):
                style_id = style_elem.get(f"{{{W_NS}}}styleId", "")
                if style_id in heading_ids:
                    return True
                name_elem = style_elem.find("w:name", NSMAP)
                if name_elem is not None:
                    name_val = name_elem.get(f"{{{W_NS}}}val", "")
                    if name_val in heading_ids or name_val.startswith(("heading", "标题")):
                        return True
                return False

            modified = False
            for style in styles:
                if not is_heading_style(style):
                    continue
                pPr = style.find("w:pPr", NSMAP)
                if pPr is None:
                    continue
                keep_next = pPr.find("w:keepNext", NSMAP)
                keep_lines = pPr.find("w:keepLines", NSMAP)
                if keep_next is not None:
                    pPr.remove(keep_next)
                    modified = True
                    print(f"  - 已从样式 {style.get(f'{{{W_NS}}}styleId', 'unknown')} 中删除 keepNext")
                if keep_lines is not None:
                    pPr.remove(keep_lines)
                    modified = True
                    print(f"  - 已从样式 {style.get(f'{{{W_NS}}}styleId', 'unknown')} 中删除 keepLines")

            if modified:
                files["word/styles.xml"] = etree.tostring(
                    styles_tree,
                    xml_declaration=True,
                    encoding="UTF-8",
                    standalone="yes"
                )
                # 重新打包
                with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as zout:
                    for name, data in files.items():
                        zout.writestr(name, data)
                print("✅ 已清理标题样式中的“与下段同页/段中不分页”属性")
            else:
                print("ℹ️ 未发现需要清理的标题样式属性")
            return True
        except Exception as e:
            print(f"❌ 清理标题样式失败: {e}")
            return False

    def remove_keep_options_from_heading_styles(self, output_file: str) -> bool:
        """
        删除所有标题样式中的 keepNext（与下段同页）和 keepLines（段中不分页）
        """
        # XML 命名空间
        W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        NSMAP = {"w": W_NS}

        def log(msg):
            print(f"[StyleCleaner] {msg}")

        try:
            # 1. 解压 docx
            with zipfile.ZipFile(output_file, "r") as zin:
                files = {name: zin.read(name) for name in zin.namelist()}

            # 2. 检查 styles.xml 是否存在
            if "word/styles.xml" not in files:
                log("❌ 未找到 word/styles.xml，无法处理样式")
                return False

            # 3. 解析 styles.xml
            styles_tree = etree.fromstring(files["word/styles.xml"])
            # 获取所有样式元素
            styles = styles_tree.xpath("//w:style", namespaces=NSMAP)

            # 定义标题样式的 ID 列表（中英文常见样式名）
            heading_ids = [
                # 英文标题
                "Heading1", "Heading2", "Heading3", "Heading4", "Heading5",
                "Heading6", "Heading7", "Heading8", "Heading9",
                # 中文标题（Word 中文版）
                "标题1", "标题2", "标题3", "标题4", "标题5",
                "标题6", "标题7", "标题8", "标题9",
                # 有时用户自定义样式可能包含这些，但为了安全只处理标准标题
            ]

            # 也可以使用更宽松的匹配：样式名称中带“标题”或“Heading”
            # 这里采用精确匹配 + 后缀判断
            def is_heading_style(style_elem):
                style_id = style_elem.get(f"{{{W_NS}}}styleId", "")
                if style_id in heading_ids:
                    return True
                # 检查样式名称（w:name）
                name_elem = style_elem.find("w:name", NSMAP)
                if name_elem is not None:
                    name_val = name_elem.get(f"{{{W_NS}}}val", "")
                    if name_val in heading_ids or name_val.startswith("heading") or name_val.startswith("标题"):
                        return True
                return False

            modified = False
            for style in styles:
                if not is_heading_style(style):
                    continue

                log(f"处理标题样式: {style.get(f'{{{W_NS}}}styleId', 'unknown')}")

                # 获取段落属性 w:pPr
                pPr = style.find("w:pPr", NSMAP)
                if pPr is None:
                    continue  # 该样式没有段落属性，跳过

                # 删除 keepNext 和 keepLines
                keep_next = pPr.find("w:keepNext", NSMAP)
                keep_lines = pPr.find("w:keepLines", NSMAP)

                if keep_next is not None:
                    pPr.remove(keep_next)
                    log("  - 已删除 keepNext (与下段同页)")
                    modified = True
                if keep_lines is not None:
                    pPr.remove(keep_lines)
                    log("  - 已删除 keepLines (段中不分页)")
                    modified = True

            if not modified:
                log("⚠️ 未在任何标题样式中找到 keepNext 或 keepLines，无需修改")
            else:
                # 写回 styles.xml
                files["word/styles.xml"] = etree.tostring(
                    styles_tree,
                    xml_declaration=True,
                    encoding="UTF-8",
                    standalone="yes"
                )
                log("✅ 样式文件已更新")

            # 4. 重新打包 docx
            with zipfile.ZipFile(output_file, "w", zipfile.ZIP_DEFLATED) as zout:
                for name, data in files.items():
                    zout.writestr(name, data)

            log(f"✅ 处理完成，文件保存至: {output_file}")
            return True

        except Exception as e:
            log(f"❌ 处理失败: {e}")
            return False


if __name__ == "__main__":
    import shutil

    processor = DocxStyleCleaner()

    # 复制输入文件到输出路径（保留原始文件不变）
    shutil.copyfile(processor.input_file, processor.output_file)
    print(f"测试文件已复制到: {processor.output_file}")

    success = processor.remove_keep_options_from_heading_styles(processor.output_file)
    if success:
        print("标题样式的“与下段同页/段中不分页”已全部取消 ✅")
    else:
        print("处理过程中出现错误，请检查输入文件路径或样式定义 ⚠️")