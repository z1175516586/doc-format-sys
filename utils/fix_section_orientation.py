import zipfile
import os
import shutil
from lxml import etree


def fix_docx_section_orientation(docx_path: str):
    """
    函数梗概：
    ----------------
    遍历 docx 中所有 sectPr（节属性），根据页面尺寸自动修正横竖版：

    规则：
    - 如果 w > h → landscape
    - 如果 h >= w → portrait

    并强制修正：
    - orient 属性
    - w / h 数值（避免“看起来横版但实际竖版”的问题）

    ⚠️ 原地修改 docx，不做备份
    """

    print("===== 进入 fix_docx_section_orientation 函数 ======")

    if not os.path.isfile(docx_path):
        print(f"[错误] 文件不存在: {docx_path}")
        return

    temp_dir = docx_path + "_tmp_orient"
    os.makedirs(temp_dir, exist_ok=True)

    try:
        # 1. 解压
        with zipfile.ZipFile(docx_path, 'r') as z:
            z.extractall(temp_dir)

        document_xml = os.path.join(temp_dir, "word", "document.xml")

        if not os.path.exists(document_xml):
            print("[错误] 未找到 document.xml")
            return

        parser = etree.XMLParser(remove_blank_text=True)
        tree = etree.parse(document_xml, parser)
        root = tree.getroot()

        ns = root.nsmap
        if None in ns:
            ns['w'] = ns.pop(None)

        sectPr_list = root.xpath(".//w:sectPr", namespaces=ns)

        print(f"[日志] 找到 sectPr 数量: {len(sectPr_list)}")

        for i, sect in enumerate(sectPr_list, 1):
            pgSz = sect.find("w:pgSz", namespaces=ns)
            if pgSz is None:
                continue

            w = pgSz.get(f"{{{ns['w']}}}w")
            h = pgSz.get(f"{{{ns['w']}}}h")

            if not w or not h:
                continue

            w = int(w)
            h = int(h)

            # =============================
            # 核心逻辑：判断并修正
            # =============================
            if w > h:
                # 横版
                pgSz.set(f"{{{ns['w']}}}orient", "landscape")

                # 确保宽 > 高（关键！）
                if w < h:
                    pgSz.set(f"{{{ns['w']}}}w", str(h))
                    pgSz.set(f"{{{ns['w']}}}h", str(w))

                print(f"[日志] 第{i}个节: 横向")

            else:
                # 竖版
                pgSz.set(f"{{{ns['w']}}}orient", "portrait")

                # 确保高 > 宽
                if h < w:
                    pgSz.set(f"{{{ns['w']}}}w", str(h))
                    pgSz.set(f"{{{ns['w']}}}h", str(w))

                print(f"[日志] 第{i}个节: 纵向")

        # 写回
        tree.write(document_xml, encoding="UTF-8", xml_declaration=True, pretty_print=True)

        # 重新打包
        with zipfile.ZipFile(docx_path, 'w', zipfile.ZIP_DEFLATED) as z:
            for foldername, _, filenames in os.walk(temp_dir):
                for filename in filenames:
                    file_path = os.path.join(foldername, filename)
                    arcname = os.path.relpath(file_path, temp_dir)
                    z.write(file_path, arcname)

        print("[日志] 方向修正完成")

    except Exception as e:
        print(f"[错误] 处理异常: {e}")

    finally:
        shutil.rmtree(temp_dir)

    print("===== 退出 fix_docx_section_orientation 函数 ======")