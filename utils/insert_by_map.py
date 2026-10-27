import argparse
import json
import os
from pathlib import Path
from typing import Optional

try:
    import win32com.client as win32
except Exception as _exc:
    win32 = None
    _WIN32COM_IMPORT_ERROR = _exc


def insert_images_fix_layout(
    target_doc_path: str,
    image_dir: str,
    output_doc_path: Optional[str] = None,
    *,
    word_visible: bool = False,
) -> Optional[str]:
    """
    基于 image_dir/image_map.json 的“图注文本 -> 图片文件名(列表)”映射，
    将图片插入到目标文档中对应图注上方。
    支持：一图多页（多张图片，逐张插入）。
    """
    if os.name != "nt":
        print("错误：insert_by_map 仅支持 Windows（依赖 Word COM）")
        return None

    if win32 is None:
        print(f"错误：win32com不可用: {_WIN32COM_IMPORT_ERROR}")
        return None

    target = Path(target_doc_path).expanduser()
    img_dir = Path(image_dir).expanduser()
    if not target.is_absolute():
        target = (Path.cwd() / target).resolve()
    if not img_dir.is_absolute():
        img_dir = (Path.cwd() / img_dir).resolve()

    json_path = img_dir / "image_map.json"
    if not json_path.exists():
        print(f"错误：找不到 image_map.json: {json_path}")
        return None

    if output_doc_path is None:
        output_doc_path = str(target.with_name(f"{target.stem}_restored_fixed.docx"))
    out_path = Path(output_doc_path).expanduser()
    if not out_path.is_absolute():
        out_path = (Path.cwd() / out_path).resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)

    with open(json_path, "r", encoding="utf-8") as f:
        image_map = json.load(f)

    print(f"目标文档: {target}")
    print(f"映射文件: {json_path}")

    word = win32.DispatchEx("Word.Application")
    word.Visible = bool(word_visible)
    word.DisplayAlerts = 0

    doc = None
    try:
        doc = word.Documents.Open(str(target))

        missing_images = 0
        missing_anchors = 0

        # ==================================================
        # 【关键修改 ①】
        # image_map 的 value 由：
        #   "xxx.png"
        # 变为：
        #   ["xxx_p3.png", "xxx_p4.png", ...]
        # ==================================================
        for caption_text, filenames in image_map.items():

            # 兼容兜底：若仍是字符串，转为列表
            if isinstance(filenames, str):
                filenames = [filenames]

            print(f"--- 正在处理图注: '{caption_text}' ---")

            # 搜索定位图注
            word.Selection.HomeKey(Unit=6)
            find_obj = word.Selection.Find
            find_obj.ClearFormatting()

            if not find_obj.Execute(FindText=caption_text, Forward=True):
                print("    [失败] 未找到锚点文字")
                missing_anchors += 1
                continue

            # 移动到图注上一行，准备插图
            word.Selection.MoveLeft(Unit=1, Count=1)
            word.Selection.TypeParagraph()
            word.Selection.MoveUp(Unit=5, Count=1)

            # 强制解除“固定行距”，避免图片显示被裁剪
            word.Selection.ParagraphFormat.LineSpacingRule = 0
            word.Selection.ParagraphFormat.SpaceBefore = 0
            word.Selection.ParagraphFormat.SpaceAfter = 0

            # ==================================================
            # 【关键修改 ②】
            # 多张图片：逐张插入（每一页一张图）
            # 不拼接，不合并
            # ==================================================
            for idx, filename in enumerate(filenames, start=1):
                image_full_path = img_dir / filename
                if not image_full_path.exists():
                    print(f"    [跳过] 图片文件丢失: {filename}")
                    missing_images += 1
                    continue

                try:
                    word.Selection.InlineShapes.AddPicture(
                        FileName=str(image_full_path),
                        LinkToFile=False,
                        SaveWithDocument=True,
                    )
                    word.Selection.ParagraphFormat.Alignment = 1  # 居中
                    print(f"    > 插入成功: {filename}")

                    # 插入多张图时，图片之间增加换行，避免重叠
                    word.Selection.TypeParagraph()

                except Exception as e:
                    print(f"    [错误] 插入失败 {filename}: {e}")

        doc.SaveAs(str(out_path))
        print(f"\n修复版文档已保存: {out_path}")
        if missing_images or missing_anchors:
            print(f"提示：缺失图片={missing_images}，未找到锚点={missing_anchors}")

        return str(out_path)

    except Exception as e:
        print(f"运行出错: {e}")
        return None

    finally:
        if doc:
            try:
                doc.Close(SaveChanges=0)
            except Exception:
                pass
        if word:
            try:
                word.Quit()
            except Exception:
                pass


def _build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Insert images into docx based on image_map.json captions.")
    parser.add_argument("--input-doc", required=True, help="Target docx file to insert images into")
    parser.add_argument("--image-dir", required=True, help="Directory containing image_map.json and images")
    parser.add_argument("--output-doc", default=None, help="Output docx path (default: <input>_restored_fixed.docx)")
    parser.add_argument("--word-visible", dest="word_visible", action="store_true", default=False, help="Show Word UI while processing")
    return parser


if __name__ == "__main__":
    args = _build_arg_parser().parse_args()
    insert_images_fix_layout(
        args.input_doc,
        args.image_dir,
        output_doc_path=args.output_doc,
        word_visible=args.word_visible,
    )
