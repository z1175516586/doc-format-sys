# -*- coding: utf-8 -*-
"""
功能：
将 Word 文档中的“自动编号”全部转换为普通文本（等价于 VBA: ActiveDocument.ConvertNumbersToText）

使用方式：
1. 修改 main 中的 file_path 为你的文件路径
2. 直接在 PyCharm 点击运行
"""

import os
import win32com.client
import pythoncom

def convert_auto_number_to_text(file_path: str):
    """
    将 Word 文档中的自动编号转换为普通文本，并原地保存

    :param file_path: Word 文件绝对路径（.docx / .docm）
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"文件不存在: {file_path}")

    pythoncom.CoInitialize()  # ✅ 必加

    word = None
    doc = None

    try:
        # 启动 Word
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # 不弹框

        # 打开文档
        doc = word.Documents.Open(file_path)

        # ✅ 核心操作（等价 VBA 一行）
        doc.ConvertNumbersToText()

        # 保存（原地覆盖）
        doc.Save()

        print("处理完成：自动编号已转换为文本")

    except Exception as e:
        print(f"处理失败: {e}")

    finally:
        # 关闭文档
        if doc:
            doc.Close(False)

        # 退出 Word
        if word:
            word.Quit()

        pythoncom.CoUninitialize()  # ✅ 释放


# =========================
# ✅ 本地测试入口（直接点运行）
# =========================
if __name__ == "__main__":
    # 👉 改成你的 Word 文件路径
    file_path = r"D:\ASE_develop_file\排版资料1014\tmpl52t7b8k - 2026年4月30日.docx"

    convert_auto_number_to_text(file_path)