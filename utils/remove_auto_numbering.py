import sys
import zipfile
import shutil
import tempfile
from pathlib import Path
import xml.etree.ElementTree as ET
import re

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}
ET.register_namespace("w", W_NS)


class NumberingConverter:
    def __init__(self):
        self.abstract_nums = {}
        self.num_map = {}
        self.counters = {}

    # -------------------------
    # 解析 numbering.xml
    # -------------------------
    def parse_numbering(self, numbering_xml):
        tree = ET.fromstring(numbering_xml)

        # abstractNum
        for abs_num in tree.findall("w:abstractNum", NS):
            abs_id = abs_num.attrib[f"{{{W_NS}}}abstractNumId"]
            lvl_map = {}

            for lvl in abs_num.findall("w:lvl", NS):
                ilvl = int(lvl.attrib[f"{{{W_NS}}}ilvl"])
                start = lvl.find("w:start", NS)
                numFmt = lvl.find("w:numFmt", NS)
                lvlText = lvl.find("w:lvlText", NS)

                start_val = int(start.attrib[f"{{{W_NS}}}val"]) if start is not None else 1

                lvl_map[ilvl] = {
                    "start": start_val if start_val > 0 else 1,
                    "numFmt": numFmt.attrib[f"{{{W_NS}}}val"] if numFmt is not None else "decimal",
                    "lvlText": lvlText.attrib[f"{{{W_NS}}}val"] if lvlText is not None else "%1",
                }

            self.abstract_nums[abs_id] = lvl_map

        # num → abstractNumId
        for num in tree.findall("w:num", NS):
            num_id = num.attrib[f"{{{W_NS}}}numId"]
            abs_id = num.find("w:abstractNumId", NS).attrib[f"{{{W_NS}}}val"]
            self.num_map[num_id] = abs_id

    # -------------------------
    # 生成编号文本
    # -------------------------
    def build_number_text(self, num_id, ilvl):
        if num_id not in self.num_map:
            return ""

        abs_id = self.num_map[num_id]
        if abs_id not in self.abstract_nums:
            return ""

        lvl_def = self.abstract_nums[abs_id].get(ilvl)
        if not lvl_def:
            return ""

        # 初始化计数器
        if num_id not in self.counters:
            self.counters[num_id] = [0] * 9

        counters = self.counters[num_id]

        # 重置低级计数
        for i in range(ilvl + 1, 9):
            counters[i] = 0

        # 保证父级编号存在
        for i in range(ilvl):
            if counters[i] == 0:
                parent_def = self.abstract_nums[abs_id].get(i)
                if parent_def:
                    counters[i] = parent_def["start"]

        # 递增当前级
        if counters[ilvl] == 0:
            counters[ilvl] = lvl_def["start"]
        else:
            counters[ilvl] += 1

        # bullet
        if lvl_def["numFmt"] == "bullet":
            return lvl_def["lvlText"] + " "

        # 替换 %1 %2
        text = lvl_def["lvlText"]
        for i in range(1, 10):
            placeholder = f"%{i}"
            if placeholder in text:
                val = counters[i - 1]
                text = text.replace(placeholder, str(val))

        return text + " "

    # -------------------------
    # 处理 document.xml
    # -------------------------
    def process_document(self, document_xml):
        detected = False
        xml_str = document_xml.decode("utf-8")

        # 匹配每个 <w:p> 段落
        p_pattern = re.compile(r"(<w:p\b[^>]*>)(.*?)(</w:p>)", re.DOTALL)
        numPr_pattern = re.compile(r"<w:numPr>.*?</w:numPr>", re.DOTALL)
        ilvl_pattern = re.compile(r'<w:ilvl w:val="(\d+)"')
        numId_pattern = re.compile(r'<w:numId w:val="(\d+)"')

        # 使用列表作为可修改的标记
        detected_flag = [False]

        def replace_num(match):
            start_tag, inner_xml, end_tag = match.groups()
            numPr_match = numPr_pattern.search(inner_xml)
            if not numPr_match:
                return match.group(0)  # 没有自动编号

            detected_flag[0] = True  # 标记已经处理过

            # 获取 ilvl 和 numId
            ilvl_m = ilvl_pattern.search(numPr_match.group(0))
            numId_m = numId_pattern.search(numPr_match.group(0))
            if not ilvl_m or not numId_m:
                return match.group(0)

            ilvl = int(ilvl_m.group(1))
            numId = numId_m.group(1)

            # 生成文本编号
            number_text = self.build_number_text(numId, ilvl)

            # 删除 <w:numPr>
            inner_xml = numPr_pattern.sub("", inner_xml, count=1)

            # 在第一个 <w:r> 前插入编号文本 <w:r><w:t xml:space="preserve">编号</w:t></w:r>
            r_insert = f'<w:r><w:t xml:space="preserve">{number_text}</w:t></w:r>'
            inner_xml = re.sub(r'(<w:r\b)', r_insert + r'\1', inner_xml, count=1)

            return start_tag + inner_xml + end_tag

        xml_str_new = p_pattern.sub(replace_num, xml_str)

        detected = detected_flag[0]

        return xml_str_new.encode("utf-8"), detected

# -------------------------
# 主函数
# -------------------------
def remove_auto_numbering_inplace(docx_path: str) -> bool:
    temp_dir = tempfile.mkdtemp()
    detected = False

    try:
        with zipfile.ZipFile(docx_path, "r") as zin:
            zin.extractall(temp_dir)

        word_dir = Path(temp_dir) / "word"
        numbering_path = word_dir / "numbering.xml"
        document_path = word_dir / "document.xml"

        converter = NumberingConverter()

        if numbering_path.exists():
            with open(numbering_path, "rb") as f:
                converter.parse_numbering(f.read())

        with open(document_path, "rb") as f:
            new_doc_xml, detected = converter.process_document(f.read())

        if not detected:
            return False

        with open(document_path, "wb") as f:
            f.write(new_doc_xml)

        # 原地覆盖
        new_docx_path = docx_path + ".tmp"

        with zipfile.ZipFile(new_docx_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for file in Path(temp_dir).rglob("*"):
                if file.is_file():
                    zout.write(file, file.relative_to(temp_dir))

        shutil.move(new_docx_path, docx_path)

        return True

    finally:
        shutil.rmtree(temp_dir)

def test_open_with_com(docx_path: str) -> bool:
    """
    使用 COM 打开并关闭 Word 文档，验证文件是否损坏
    """
    try:
        import win32com.client

        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # 不弹错误对话框

        doc = word.Documents.Open(docx_path)
        doc.Close(False)

        word.Quit()
        return True

    except Exception as e:
        print("COM 打开失败：", e)
        return False

if __name__ == "__main__":
    path = r"D:\ASE_develop_file\doc-convertor\result\鲅鱼圈1号高炉热风炉煤气换热器改造工程  可研2026.1.21---姜文雪.docx"

    ok = remove_auto_numbering_inplace(path)

    if ok:
        print("自动编号已转换为文本。")

        if test_open_with_com(path):
            print("Word COM 打开成功，文件未损坏。")
        else:
            print("Word COM 打开失败，文件可能损坏。")
    else:
        print("未检测到自动编号。")