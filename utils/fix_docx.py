# 修复 docx 文件中 .rels 里的 Target="../NULL" 非法关系
# 这种情况出现的原因大部分是复制粘贴导致的，导致文档中存在无效的引用
# 带有NULL引用的文件，可以被wps或word打开，但是python-docx 无法读取
import zipfile
import re
import os
import shutil
import tempfile
import xml.etree.ElementTree as ET



def fix_docx_null_relationships(docx_path: str) -> bool:
    """
    修复 docx 文件中 .rels 里的 Target="../NULL" 非法关系

    :param docx_path: docx 文件路径
    :return: 是否做了修改 (True=已修复, False=未发现问题)
    """

    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"文件不存在: {docx_path}")

    if not zipfile.is_zipfile(docx_path):
        print("⚠️ 文件不是合法的 docx(zip)结构，跳过 NULL 修复")
        return False

    print("🔍 开始扫描 NULL 关系...")

    modified = False

    with zipfile.ZipFile(docx_path, "r") as zin:
        files = {item.filename: zin.read(item.filename) for item in zin.infolist()}

    for name, content in files.items():
        if name.endswith(".rels"):
            try:
                text = content.decode("utf-8")
            except UnicodeDecodeError:
                continue

            if 'Target="../NULL"' in text:
                print(f"🔧 修复 NULL 引用: {name}")

                # 删除整个 NULL Relationship 节点
                new_text = re.sub(
                    r'<Relationship[^>]*Target="\.\./NULL"[^>]*/>',
                    '',
                    text
                )

                files[name] = new_text.encode("utf-8")
                modified = True

    if modified:
        print("🛠 检测到非法 NULL 关系，正在重写 docx...")
        with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for fname, data in files.items():
                zout.writestr(fname, data)
        print("✅ NULL 关系已清理完成")
    else:
        print("✅ 未发现 NULL 关系")

    return modified

# 修复“找不到 Heading 样式”问题
# 首先需要检查 styles.xml 是否存在 heading 样式
# 如果不存在，需要手动添加
# 还需要修改 document.xml 中的段落样式引用
# 设置 w:pStyle 为 heading 样式

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NSMAP = {"w": W_NS}

def fix_docx_heading_styles(docx_path: str) -> None:
    print("🔧 开始修复 Heading 样式...")

    with zipfile.ZipFile(docx_path, "r") as zin:
        files = {item.filename: zin.read(item.filename) for item in zin.infolist()}

    if "word/styles.xml" not in files:
        print("❌ 未找到 styles.xml")
        return

    content = files["word/styles.xml"].decode("utf-8")

    modified = False
    missing_levels = []   # ✅ 收集缺失级别

    level_to_styleid = {
        1: 100,
        2: 101,
        3: 102,
        4: 103
    }

    for level in (1, 2, 3, 4):
        if f'w:name w:val="heading {level}"' not in content.lower():
            print(f"➕ 插入缺失样式: heading {level}")
            missing_levels.append(level)

            style_block = f"""
      <w:style w:type="paragraph" w:styleId="{level_to_styleid[level]}">
        <w:name w:val="heading {level}"/>
        <w:basedOn w:val="a"/>
        <w:next w:val="a"/>
        <w:qFormat/>
        <w:pPr>
          <w:keepNext/>
          <w:keepLines/>
          <w:outlineLvl w:val="{level - 1}"/>
        </w:pPr>
      </w:style>
    """
            content = content.replace("</w:styles>", style_block + "\n</w:styles>")
            modified = True
        else:
            print(f"✅ 已存在样式: heading {level}")

    # ✅ 如果有缺失级别，同时修复 document.xml
    document_modified = False
    if missing_levels:
        print("🔄 同步修复 document.xml 中缺失的 heading 段落...")
        document_modified = fix_missing_headings_after_toc(files, missing_levels)

    # ✅ 只要 styles 或 document 有修改，都写回
    if modified or document_modified:

        print("🛠 正在写回 docx...")

        with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for fname, data in files.items():
                if fname == "word/styles.xml" and modified:
                    zout.writestr(fname, content.encode("utf-8"))
                else:
                    zout.writestr(fname, data)

        print("🎉 修复完成")

    else:
        print("无需修复")
# 需要由外层传入：missing_levels = [2,3,4] 这样的列表
# 对应 styleId 101,102,103

def _assign_headings_in_next_section(tail: str, missing_levels: list[int]) -> tuple[str, bool]:
    """
    在目录后的小节中，任意选择若干段落，
    将其设置为缺失的 heading 样式
    """

    style_id_map = {
        1: "100",
        2: "101",
        3: "102",
        4: "103"
    }

    modified = False
    level_index = 0

    # 匹配普通段落 <w:p>...</w:p>
    paragraph_pattern = re.compile(r'(<w:p\b.*?>.*?</w:p>)', re.DOTALL)

    def replace_paragraph(match):
        nonlocal level_index, modified

        if level_index >= len(missing_levels):
            return match.group(1)

        para_block = match.group(1)

        # 如果已有 pStyle 就替换
        if "<w:pStyle" in para_block:
            para_block = re.sub(
                r'<w:pStyle[^>]+/>',
                f'<w:pStyle w:val="{style_id_map[missing_levels[level_index]]}"/>',
                para_block,
                count=1
            )
        else:
            # 插入到 <w:pPr> 内
            if "<w:pPr>" in para_block:
                para_block = para_block.replace(
                    "<w:pPr>",
                    f'<w:pPr><w:pStyle w:val="{style_id_map[missing_levels[level_index]]}"/>',
                    1
                )
            else:
                # 没有 pPr 就创建
                para_block = para_block.replace(
                    "<w:p>",
                    f'<w:p><w:pPr><w:pStyle w:val="{style_id_map[missing_levels[level_index]]}"/></w:pPr>',
                    1
                )

        modified = True
        level_index += 1
        return para_block

    new_tail = paragraph_pattern.sub(replace_paragraph, tail)

    return new_tail, modified

def fix_missing_headings_after_toc(files: dict, missing_levels: list[int]) -> bool:
    """
    在 document.xml 中：
    从文档末尾开始反向查找段落，
    仅选择有文本内容的段落，
    设置为缺失的 heading 样式，
    并打印被修改段落的关键字
    """

    if "word/document.xml" not in files:
        print("❌ 未找到 document.xml")
        return False

    content = files["word/document.xml"].decode("utf-8")

    # 匹配所有段落
    paragraph_pattern = re.compile(r'(<w:p\b.*?>.*?</w:p>)', re.DOTALL)
    paragraphs = list(paragraph_pattern.finditer(content))

    if not paragraphs:
        print("⚠️ 未找到任何段落")
        return False

    style_id_map = {
        1: "100",
        2: "101",
        3: "102",
        4: "103"
    }

    modified = False
    level_index = 0

    content_list = list(content)

    # 从后往前遍历段落
    for match in reversed(paragraphs):

        if level_index >= len(missing_levels):
            break

        para_block = match.group(1)

        # 提取纯文本
        text_match = re.findall(r'<w:t[^>]*>(.*?)</w:t>', para_block)
        preview_text = "".join(text_match).strip()

        # ✅ 跳过无文本或纯空白段落
        if not preview_text:
            continue

        preview_print = preview_text[:20]

        new_para = para_block

        # 替换或插入 pStyle
        if "<w:pStyle" in para_block:
            new_para = re.sub(
                r'<w:pStyle[^>]+/>',
                f'<w:pStyle w:val="{style_id_map[missing_levels[level_index]]}"/>',
                para_block,
                count=1
            )
        else:
            if "<w:pPr>" in para_block:
                new_para = para_block.replace(
                    "<w:pPr>",
                    f'<w:pPr><w:pStyle w:val="{style_id_map[missing_levels[level_index]]}"/>',
                    1
                )
            else:
                new_para = para_block.replace(
                    "<w:p>",
                    f'<w:p><w:pPr><w:pStyle w:val="{style_id_map[missing_levels[level_index]]}"/></w:pPr>',
                    1
                )

        start, end = match.span()
        content_list[start:end] = list(new_para)

        print(f"📝 设置 heading {missing_levels[level_index]} → 段落关键字: {preview_print}")

        modified = True
        level_index += 1

    if modified:
        files["word/document.xml"] = "".join(content_list).encode("utf-8")
        print("🛠 document.xml 已更新")
    else:
        print("⚠️ 未找到可用文本段落进行 heading 设置")

    return modified

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

def preprocess_headings(source_file: str) -> str:
    """
    步骤0: 预处理标题样式（严格编号规则 + 顺序比较 + 反向纠错）

    处理规则：
    - 识别“目录”
    - 目录后遇到的第一个“1”级标题忽略
    - 从第二个“1”级标题开始，正式处理正文标题（第二个 1 本身必须被处理）
    - 标题编号必须以数字结尾，"." 后必须接数字，禁止 "3 . xxx" 误判
    """
    import os
    import shutil
    import re
    from docx import Document
    from docx.shared import Pt

    print("-" * 50)
    print("📝 步骤0: 预处理标题样式（正文范围内）")

    try:

        doc = Document(source_file)

        # =========================================================
        # ① 标题正则
        # =========================================================
        heading_pattern = re.compile(
            r'^(\d+(?:\.\d+){0,3})[\s\t]*(?![)）\.])(.+)$'
        )

        # =========================================================
        # ② 标题序号顺序判断
        # =========================================================
        def is_valid_heading_order(prev_num: str, curr_num: str) -> bool:
            A = [int(x) for x in prev_num.split('.')]
            B = [int(x) for x in curr_num.split('.')]

            max_len = max(len(A), len(B))
            A += [0] * (max_len - len(A))
            B += [0] * (max_len - len(B))

            for i in range(max_len):
                if A[i] < B[i]:
                    diff = B[i] - A[i]
                    return diff in (0, 1)
                if A[i] > B[i]:
                    if i > 0 and A[i - 1] + 1 == B[i - 1]:
                        if sum(B[i:]) == 0:
                            return True
                    return False
            return True

        # =========================================================
        # ③ 扫描目录
        # =========================================================
        toc_regex = re.compile(r'目[\s\t]*录')
        after_toc = False
        level1_count = 0
        processing_enabled = False
        prev_heading_number = None

        print("🔎 正在扫描“目录”关键字…")

        for para in doc.paragraphs:
            text = para.text.strip()
            if not text:
                continue

            # ------------------ 扫描目录 ------------------
            if not after_toc:
                if toc_regex.search(text):
                    after_toc = True
                    print(f"✅ 识别到目录段落: '{text}'")
                # 目录前完全不处理标题
                continue

            # ------------------ 目录之后，正文处理 ------------------
            match = heading_pattern.match(text)
            if not match:
                continue

            # -------- 新增括号平衡校验 --------
            num_right = text.count(')') + text.count('）')
            num_left = text.count('(') + text.count('（')
            if num_right - num_left >= 1:
                # 不认为是标题，直接跳过
                print(f"   ⚠️ 段落括号不平衡，忽略标题: '{text}'")
                continue

            curr_number = match.group(1)
            title_text = match.group(2)
            level = curr_number.count('.') + 1

            # 第1个1级标题忽略，第2个开始处理
            if not processing_enabled:
                if curr_number == "1" and level == 1:
                    level1_count += 1
                    if level1_count == 1:
                        print(f"⏭️ 遇到第 1 个 1 级标题，暂不处理: '{text}'")
                        continue
                    elif level1_count == 2:
                        processing_enabled = True
                        prev_heading_number = None
                        print(f"📌 已定位正文起点（第二个 1 级标题）: '{text}'")
                        # ⚠️ 不 continue，让第二个1级标题自己参与处理
                else:
                    # 第1个1级标题和第2个1级标题之间的段落不处理
                    continue

            # ---------- 顺序判断 ----------
            if prev_heading_number is None:
                is_valid = True
            else:
                is_valid = is_valid_heading_order(prev_heading_number, curr_number)

            if not is_valid:
                print(
                    f"   ❌ 不是标题 {curr_number} {title_text}，"
                    f"与前一个标题 {prev_heading_number} 顺序不匹配"
                )
                continue

            print(
                f"   🔍 发现标题: {curr_number} {title_text} "
                f"(级别: {level}) ✅"
            )

            # ---------- 设置标题样式 ----------
            if level == 1:
                para.style = 'Heading 1'
            elif level == 2:
                para.style = 'Heading 2'
            elif level == 3:
                para.style = 'Heading 3'
            else:
                para.style = 'Heading 4'

            para.paragraph_format.space_before = Pt(0)
            para.paragraph_format.space_after = Pt(0)

            prev_heading_number = curr_number

        # =========================================================
        # ④ 反向纠错
        # =========================================================
        heading_styles = ['Heading 1', 'Heading 2', 'Heading 3', 'Heading 4']

        for para in doc.paragraphs:
            if para.style.name in heading_styles:
                text = para.text.strip()
                if not heading_pattern.match(text):
                    para.style = 'Normal'
                    para.paragraph_format.space_before = Pt(0)
                    para.paragraph_format.space_after = Pt(0)
                    if text:
                        print(f"⚠️ 原文档格式可疑处: '{text[:30]}...'")

        doc.save(source_file)
        print("✅ 标题预处理完成!")
        return source_file

    except Exception as e:
        print(f"❌ 标题预处理过程中发生错误: {str(e)}")
        return source_file


if __name__ == "__main__":
    path = r"D:\ASE_develop_file\doc-convertor\result\test.docx"
    fix_docx_heading_styles(path)
    if test_open_with_com(path):
        print("Word COM 打开成功，文件未损坏。")
    else:
        print("Word COM 打开失败，文件可能损坏。")
    from utils.remove_auto_numbering import remove_auto_numbering_inplace
    remove_auto_numbering_inplace(path)
    print("自动编号已转换为文本。")
    preprocess_headings(path)
    print("标题预处理完成!")
