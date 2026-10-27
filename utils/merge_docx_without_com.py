import os
import shutil
from docx import Document
from copy import deepcopy


def is_empty_paragraph(p):
    """
    判断是否为空段落（没有文本、没有图片等）
    """
    texts = p.xpath('.//w:t')
    return len(texts) == 0


def merge_docx_without_com(base_file, append_file, output_file):
    """
    合并两个 docx 文件（不使用 COM），修复：
    1. 页面方向错乱
    2. 分页符问题（改为分节）
    3. 首行空行问题
    """

    import os
    import shutil
    from docx import Document
    from copy import deepcopy
    from docx.oxml import OxmlElement

    # 确保目录存在
    os.makedirs(os.path.dirname(output_file), exist_ok=True)

    # 复制 base 作为结果文件
    shutil.copy2(base_file, output_file)

    # 打开文档
    doc_base = Document(output_file)
    doc_append = Document(append_file)

    body_base = doc_base.element.body
    body_append = doc_append.element.body

    # ===============================
    # 🔥 关键1：处理 base 的最后一个 sectPr
    # ===============================
    base_sectPr_list = body_base.xpath('./w:sectPr')
    if not base_sectPr_list:
        raise Exception("❌ base 文档没有 sectPr，结构异常")

    base_sectPr = base_sectPr_list[0]

    # 从 body 移除（避免后面覆盖）
    body_base.remove(base_sectPr)

    # 创建分节段落（把原 sectPr 挂进去）
    p = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    pPr.append(base_sectPr)
    p.append(pPr)

    # 插入（只执行一次）
    body_base.append(p)

    # ===============================
    # 🔥 关键2：处理 append 文档
    # ===============================
    elements = list(body_append)

    # 👉 跳过开头空段落（避免多空行）
    start_index = 0
    while start_index < len(elements):
        el = elements[start_index]
        if el.tag.endswith('p'):
            texts = el.xpath('.//w:t')
            if len(texts) == 0:
                start_index += 1
                continue
        break

    # ===============================
    # 🔥 关键3：拷贝 append 内容（保留 sectPr）
    # ===============================
    for i in range(start_index, len(elements)):
        el = elements[i]
        new_el = deepcopy(el)
        body_base.append(new_el)

    # ===============================
    # 保存
    # ===============================
    doc_base.save(output_file)

    print(f"✅ 合并完成: {output_file}")

if __name__ == "__main__":
    base_file = r"D:\ASE_develop_file\doc-convertor\sql\test.docx"
    append_file = r"D:\ASE_develop_file\doc-convertor\sql\append.docx"
    output_file = r"D:\ASE_develop_file\doc-convertor\sql\merged.docx"

    merge_docx_without_com(base_file, append_file, output_file)