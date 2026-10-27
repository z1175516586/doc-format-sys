# doc_parser_projName.py

from docx import Document

class DocParser:
    def __init__(self, file_path: str):
        """
        初始化 DocParser 实例
        :param file_path: Word 文件路径 (.docx)
        """
        self.file_path = file_path
        self.doc = Document(file_path)

    def get_first_nonempty_paragraph_text(self, file_type: int = 1) -> tuple[str, str, int]:
        """
        返回:
            result_text,
            last_normalized_para,
            err_code
        """

        stop_words_map = {
            1: "初步设计",
            2: "可行性研究",
            3: "初步设计（代可研）",
            4: "可行性研究（代初设）"
        }

        stop_word = stop_words_map.get(file_type)
        if not stop_word:
            return "", "", 1002  # 非法类型

        normalized_stop = "".join(stop_word.split())
        result_text = ""
        last_normalized_para = ""
        total_length = 0  # 新增：累计字符数

        for para in self.doc.paragraphs:
            para_text = para.text
            if not para_text.strip():
                continue

            normalized_para = "".join(para_text.split())
            last_normalized_para = normalized_para

            idx = normalized_para.find(normalized_stop)

            if idx != -1:
                result_text += normalized_para[:idx]
                return result_text, last_normalized_para, 0  # 成功

            result_text += normalized_para
            total_length += len(normalized_para)

            # 新增逻辑：超过100字符仍未找到
            if total_length >= 50:
                return result_text, last_normalized_para, 1001

        # 遍历结束仍未找到
        return result_text, last_normalized_para, 1001

if __name__ == "__main__":
    # 简单测试
    parser = DocParser(r"D:\ASE_develop_file\doc-convertor\result\converted_鞍钢股份炼焦总厂1#、2#焦炉烟气脱硫脱硝扩容改造项目20250618.docx")
    text, last_para = parser.get_first_nonempty_paragraph_text(1)
    print(f"第一个段落文本（去空格后）：{text}，最后处理段落：{last_para}")
