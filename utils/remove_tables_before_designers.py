# remove_tables_before_designers.py
import os
from docx import Document
from docx.oxml import OxmlElement

class DocxTableProcessor:
    """
    删除第一个小节中“设计人员名单”之前的所有表格，
    但保留表格内容为普通段落。
    """

    def __init__(self, input_file: str):
        if not os.path.exists(input_file):
            raise FileNotFoundError(f"文件不存在: {input_file}")
        self.input_file = input_file
        self.doc = Document(input_file)

        # 输出路径，默认为覆盖原文件
        self.output_file = input_file

    def remove_tables_before_designers(self) -> bool:
        """
        核心函数：删除第一个小节“设计人员名单”前的表格，保留内容
        """
        def log(msg):
            print(f"[TableProcessor] {msg}")

        try:
            body_elm = self.doc.element.body
            to_remove = []

            # 标记是否遇到设计人员名单
            designers_found = False

            # 遍历 body 子元素
            for child in list(body_elm):
                texts = "".join(child.itertext())
                if "设计人员名单" in texts:
                    designers_found = True
                    break  # 停止处理，保留设计人员名单及之后内容

                if child.tag.endswith('tbl'):
                    to_remove.append(child)

            if not to_remove:
                log("没有需要处理的表格")
                self.doc.save(self.output_file)
                log(f"已保存文件: {self.output_file}")
                return True

            log(f"发现 {len(to_remove)} 个表格需要处理")

            for tbl in to_remove:
                parent = tbl.getparent()
                insert_index = parent.index(tbl)

                new_paragraphs_text = []
                for row in tbl.findall('w:tr', namespaces=tbl.nsmap):
                    row_cells = []
                    for cell in row.findall('w:tc', namespaces=tbl.nsmap):
                        cell_paragraphs = []
                        for p in cell.findall('w:p', namespaces=tbl.nsmap):
                            texts = [t.text.strip() for t in p.findall('.//w:t', namespaces=tbl.nsmap) if t.text]
                            if texts:
                                cell_paragraphs.append("".join(texts))
                        row_cells.append(" ".join(cell_paragraphs))
                    new_paragraphs_text.append("\t".join(row_cells))

                for idx, text in enumerate(new_paragraphs_text):
                    p = OxmlElement('w:p')
                    r = OxmlElement('w:r')
                    t = OxmlElement('w:t')
                    t.text = text
                    r.append(t)
                    p.append(r)
                    parent.insert(insert_index + idx, p)

                parent.remove(tbl)

            # 覆盖保存原文件
            self.doc.save(self.output_file)
            log(f"✅ 已覆盖原文件: {self.output_file}")
            return True

        except Exception as e:
            log(f"❌ 处理失败: {e}")
            return False
