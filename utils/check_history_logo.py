# 属性同时尝试 {namespace}name 和普通 name。

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

# Word 命名空间
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
SETTINGS_XML_PATH = "word/settings.xml"

def print_dochistoryid(file_path: str) -> dict:
    """
    ✅ 查找指定 docx 文件 settings.xml 中的 dochistoryid
    返回格式：
    {
        "exists": True/False,
        "history_id": str or None
    }
    并打印 history_id 是否存在及其值
    """
    suffix = Path(file_path).suffix.lower()
    if suffix != ".docx" or not zipfile.is_zipfile(file_path):
        result = {"exists": False, "history_id": None}
        print(f"📌 history_id exists: {result['exists']}, value: {result['history_id']}")
        return result

    history_id = None
    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            if SETTINGS_XML_PATH in z.namelist():
                xml_data = z.read(SETTINGS_XML_PATH)
                root = ET.fromstring(xml_data)
                # 找到 docVars 节点
                doc_vars = root.find(f"{{{W_NS}}}docVars")
                if doc_vars:
                    for var in doc_vars.findall(f"{{{W_NS}}}docVar"):
                        # ✅ 属性是带命名空间前缀的，需要加 W_NS
                        name = var.attrib.get(f"{{{W_NS}}}name") or var.attrib.get("name")
                        val = var.attrib.get(f"{{{W_NS}}}val") or var.attrib.get("val")
                        # 不区分大小写匹配
                        if name and name.lower() == "dochistoryid":
                            history_id = val
                            break
    except Exception as e:
        print(f"⚠️ 读取失败: {e}")

    result = {"exists": bool(history_id), "history_id": history_id}
    print(f"📌 history_id exists: {result['exists']}, value: {result['history_id']}")
    return result