def remove_picture_paragraph_indent(doc_path: str):
    """
    查找 docx 中所有包含图片的段落，并删除（覆盖）其缩进

    :param doc_path: docx 文件绝对路径
    """
    from docx import Document
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    doc = Document(doc_path)

    def contains_picture(paragraph):
        # 判断段落中是否有图片（w:drawing 或 w:pict）
        return bool(paragraph._element.xpath('.//w:drawing') or
                    paragraph._element.xpath('.//w:pict'))

    count = 0

    for paragraph in doc.paragraphs:
        if not contains_picture(paragraph):
            continue

        p = paragraph._element
        pPr = p.find(qn('w:pPr'))

        if pPr is None:
            pPr = OxmlElement('w:pPr')
            p.insert(0, pPr)

        # ===== 核心：覆盖缩进（不要删除）=====
        ind = pPr.find(qn('w:ind'))
        if ind is None:
            ind = OxmlElement('w:ind')
            pPr.append(ind)

        ind.set(qn('w:left'), '0')
        ind.set(qn('w:leftChars'), '0')
        ind.set(qn('w:firstLine'), '0')
        ind.set(qn('w:firstLineChars'), '0')

        count += 1

    doc.save(doc_path)
    print(f"处理完成，共修改 {count} 个图片段落")

if __name__ == "__main__":
    target_file = r"D:\ASE_develop_file\原稿＋排版后\260401154520 - 副本.docx"
    remove_picture_paragraph_indent(target_file)
