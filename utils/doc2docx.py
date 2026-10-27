import os
import win32com.client as win32
import pythoncom


import os
import pythoncom
import win32com.client


def convert_doc_to_docx(doc_path: str) -> bool:
    """
    将 .doc 转换为 .docx
    成功返回 True，失败返回 False
    """

    if not os.path.exists(doc_path):
        return False

    if not doc_path.lower().endswith(".doc"):
        return False

    docx_path = os.path.splitext(doc_path)[0] + ".docx"

    if os.path.exists(docx_path):
        return True

    word = None

    try:
        pythoncom.CoInitialize()

        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        doc = word.Documents.Open(doc_path)
        doc.SaveAs(docx_path, FileFormat=16)  # 16 = wdFormatDocumentDefault (.docx)
        doc.Close()

        return os.path.exists(docx_path)

    except Exception:
        return False

    finally:
        if word:
            try:
                word.Quit()
            except Exception:
                pass

        pythoncom.CoUninitialize()


if __name__ == "__main__":
    target_file = r"D:\doc-convertor\app\src_files\test\11600035_鞍钢联众炼钢一线炉余热回收项目可研0000.doc"
    convert_doc_to_docx(target_file)
