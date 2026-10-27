from __future__ import annotations

import os
import time
from pathlib import Path
from typing import Iterable, List, Tuple

import win32com.client as win32

try:
    from PIL import ImageGrab
except ImportError:
    print("错误：未找到 Pillow (PIL) 库，请运行 `pip install Pillow` 后重试。")
    raise SystemExit(1)

try:
    import pythoncom
except ImportError:
    pythoncom = None  # type: ignore[assignment]

# Word COM 常量
WD_INLINE_SHAPE_EMBEDDED_OLE_OBJECT = 1

# 默认图片输出目录：项目根目录下 temp/formula_images
PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_IMAGE_DIR = PROJECT_ROOT / "temp" / "formula_images"


def _ensure_path(path: os.PathLike[str] | str) -> Path:
    """将路径转换为绝对 Path。"""
    return Path(path).expanduser().resolve()

from typing import List, Dict, Any, Tuple
def replace_ole_with_images(
    docx_path: os.PathLike[str] | str,
    output_docx_path: os.PathLike[str] | str,
    output_img_dir: os.PathLike[str] | str | None = None,
    *,
    word_visible: bool = False,
        word = None
) -> Tuple[List[str], List[Dict[str, Any]]]:
    """
    将 Word 文档中的 OLE 内联对象（如公式）转换成嵌入式图片。

    Args:
        docx_path: 源文档路径。
        output_docx_path: 输出文档路径。
        output_img_dir: 图片输出目录，默认写入 temp/formula_images。
        word_visible: 是否显示 Word UI，默认后台执行。

    Returns:
        List[str]: 所有生成图片的绝对路径。
    """
    if word is None:
        print("word is None")
    formula_info_list = []
    src = _ensure_path(docx_path)
    dst = _ensure_path(output_docx_path)
    image_dir = _ensure_path(output_img_dir) if output_img_dir else DEFAULT_IMAGE_DIR

    image_dir.mkdir(parents=True, exist_ok=True)
    dst.parent.mkdir(parents=True, exist_ok=True)

    print(f"[公式转图片] 源文件: {src}")
    print(f"[公式转图片] 输出文件: {dst}")
    print(f"[公式转图片] 图片目录: {image_dir}")

    # word = None
    doc = None
    saved_images: List[str] = []
    pythoncom_initialized = False
    owns_word = False

    try:
        if pythoncom:
            try:
                pythoncom.CoInitialize()
                pythoncom_initialized = True
            except Exception:
                pass

        # 若传入的 word 不可用，则自行创建
        if word is None or not hasattr(word, "Documents"):
            try:
                word = win32.Dispatch("Word.Application")
            except Exception:
                word = win32.DispatchEx("Word.Application")
            word.Visible = word_visible
            word.DisplayAlerts = False
            try:
                word.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
            except Exception:
                pass
            owns_word = True

        doc = word.Documents.Open(str(src))
        print("[公式转图片] Word 已打开文档。")

        shapes = doc.InlineShapes
        shape_count = shapes.Count
        replaced_count = 0

        if shape_count:
            print(f"[公式转图片] 检测到 {shape_count} 个内联对象，开始处理。")
        else:
            print("[公式转图片] 未检测到内联对象。")

        for index in range(shape_count, 0, -1):
            try:
                shape = shapes(index)
            except Exception:
                continue

            try:
                if shape.Type != WD_INLINE_SHAPE_EMBEDDED_OLE_OBJECT:
                    continue

                try:
                    prog_id = shape.OLEFormat.ProgID
                except Exception:
                    prog_id = "Unknown"

                print(f"  > 第 {index} 个对象：OLE 类型（ProgID={prog_id}），准备替换。")

                original_range = shape.Range
                shape.Select()
                word.Selection.CopyAsPicture()
                time.sleep(0.1)  # 等待剪贴板刷新

                image = ImageGrab.grabclipboard()
                if not image:
                    print("    ! 剪贴板未捕获到图片，跳过该公式。")
                    continue

                file_name = f"formula_shape_{index}.png"
                save_path = image_dir / file_name
                image.save(save_path)
                saved_images.append(str(save_path))

                # 记录插入前的段落样式和对齐设置，以便插入图片后恢复
                original_para = original_range.Paragraphs(1)
                original_style = None
                original_alignment = None
                original_left_indent = None
                original_first_line_indent = None
                try:
                    original_style = original_para.Style
                    original_alignment = original_para.Alignment
                    original_left_indent = original_para.LeftIndent
                    original_first_line_indent = original_para.FirstLineIndent
                except Exception:
                    pass

                shape.Delete()

                new_shape = doc.InlineShapes.AddPicture(
                    FileName=str(save_path),
                    LinkToFile=False,
                    SaveWithDocument=True,
                    Range=original_range,
                )
                new_shape.LockAspectRatio = True

                #
                uid = f"FORMULA_{index}"
                if image:
                    width, height = image.size
                    print(f"[INFO]原公式图片尺寸：width={width}px, height={height}px")
                    formula_info_list.append({
                        'uid': uid,
                        'width': width,
                        'height': height
                    })
                print("begin")
                # new_shape.Title = uid
                print("title over")
                new_shape.AlternativeText = uid
                # print(f"DEBUG - 立即验证 Title: {new_shape.Title}")
                print(f"DEBUG - 立即验证 AlternativeText: {new_shape.AlternativeText}")
                doc.Save()

                try:
                    new_para = new_shape.Range.ParagraphFormat
                    if original_style is not None:
                        try:
                            new_para.Style = original_style
                        except Exception:
                            pass
                    # 强制左对齐，消除额外缩进
                    new_para.Alignment = 0  # wdAlignParagraphLeft
                    new_para.LeftIndent = 0
                    new_para.FirstLineIndent = 0
                except Exception:
                    pass

                replaced_count += 1
            except Exception as err:
                print(f"    ! Error processing shape #{index}: {err}")

        if replaced_count:
                print(f"[公式转图片] 共替换 {replaced_count} 个 OLE 对象为图片。")
        else:
            print("[公式转图片] 未替换任何 OLE 对象。")

        print(f"[公式转图片] 正在保存文档到: {dst}")
        doc.SaveAs(str(dst))
    except Exception as exc:
        print(f"[公式转图片] 处理文档时出错: {exc}")
        raise
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass
            print("[公式转图片] 文档已关闭。")

        if owns_word and word is not None:
            try:
                word.Quit()
            except Exception:
                pass
            print("[公式转图片] Word 应用已退出。")

        if pythoncom_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    return saved_images, formula_info_list


def _default_output_path(input_path: Path) -> Path:
    """默认输出路径：原名后缀 _equations."""
    name, ext = os.path.splitext(input_path.name)
    return input_path.with_name(f"{name}_equations{ext}")


def main(argv: Iterable[str] | None = None) -> None:
    import argparse
    import inspect

    parser = argparse.ArgumentParser(description="将 Word 文档中的公式（OLE）转换成图片。")
    parser.add_argument("input", help="源 DOCX 文件路径。")
    parser.add_argument(
        "-o",
        "--output",
        help="转换后的 DOCX 输出路径，默认在原文件名后加 _equations。",
    )
    parser.add_argument(
        "-d",
        "--image-dir",
        help="图片输出目录，默认写入项目 temp/formula_images。",
    )
    parser.add_argument(
        "--visible",
        action="store_true",
        help="是否在前台显示 Word 窗口，默认后台运行。",
    )

    args = parser.parse_args(tuple(argv) if argv is not None else None)

    input_path = _ensure_path(args.input)
    output_path = _ensure_path(args.output) if args.output else _default_output_path(input_path)
    image_dir = _ensure_path(args.image_dir) if args.image_dir else None

    images = replace_ole_with_images(
        docx_path=input_path,
        output_docx_path=output_path,
        output_img_dir=image_dir,
        word_visible=args.visible,
    )

    print("\nSummary")
    print("-------")
    print(f"Processed document : {output_path}")
    print(f"Images saved       : {len(images)}")
    if images:
        print(f"Image directory    : {Path(images[0]).parent}")


if __name__ == "__main__":
    main()
