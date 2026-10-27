import argparse
import ctypes
import json
import time
from pathlib import Path
from typing import Optional, Tuple, Dict

try:
    import win32com.client as win32
except Exception as _exc:
    win32 = None
    _WIN32COM_IMPORT_ERROR = _exc


# ================= 工具函数 =================
def force_clear_clipboard():
    try:
        if ctypes.windll.user32.OpenClipboard(None):
            ctypes.windll.user32.EmptyClipboard()
            ctypes.windll.user32.CloseClipboard()
    except:
        pass


def save_clipboard_to_image(save_path):
    for _ in range(20):
        try:
            from PIL import ImageGrab
            img = ImageGrab.grabclipboard()
            if img is not None:
                if isinstance(img, list):
                    return False
                img.save(save_path, "PNG")
                return True
        except:
            time.sleep(0.5)
        time.sleep(0.5)
    return False


# ================= 核心逻辑 =================
def extract_and_generate_map(
    src_path: str,
    output_dir: str,
    output_doc_path: Optional[str] = None,
    *,
    word_visible: bool = True,
    word=None
) -> Tuple[str, str, Dict[str, list]]:
    """
    从 docx 中按“图注”定位提取 Shapes。
    支持“一图多页”：同一图注在多页出现时，每一页生成一张图片（不拼接）。
    """

    src = Path(src_path).expanduser().resolve()
    out_dir = Path(output_dir).expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    if output_doc_path is None:
        output_doc_path = str(out_dir / f"{src.stem}_clean.docx")
    out_doc = Path(output_doc_path).expanduser().resolve()
    out_doc.parent.mkdir(parents=True, exist_ok=True)

    # 映射字典：{ "图注原文": [img1.png, img2.png, ...] }
    image_map: Dict[str, list] = {}

    if win32 is None:
        raise ImportError(f"win32com 不可用: {_WIN32COM_IMPORT_ERROR}")

    wdVerticalPositionRelativeToPage = 6
    wdActiveEndPageNumber = 3

    doc = None
    try:
        if word is None:
            word = win32.Dispatch("Word.Application")
            word.Visible = word_visible
        doc = word.Documents.Open(str(src))

        # ================= 1. 扫描图注 =================
        captions = []
        print("正在扫描图注位置...")
        for p in doc.Paragraphs:
            if len(p.Range.Text) > 100:
                continue
            text = p.Range.Text.strip()
            if text.startswith("图") and len(text) > 2:
                try:
                    captions.append({
                        "text": text,
                        "visual_top": p.Range.Information(wdVerticalPositionRelativeToPage),
                        "page": p.Range.Information(wdActiveEndPageNumber)
                    })
                except:
                    pass

        # ================= 2. 按图注处理 Shapes =================
        processed_shapes = set()

        for cap in captions:
            target_text = cap["text"]
            target_page = cap["page"]
            bottom_limit = cap["visual_top"]

            print(f"--- 处理图注: {target_text} ---")

            # --------------------------------------------------
            # 【关键修改 ①】
            # 不再使用单一 shapes_to_select
            # 而是：按页收集 Shapes
            # page_shapes_map = { page_num: [shape_name, ...] }
            # --------------------------------------------------
            page_shapes_map: Dict[int, list] = {}

            for shape in doc.Shapes:
                try:
                    if shape.Name in processed_shapes:
                        continue

                    shape_page = shape.Anchor.Information(wdActiveEndPageNumber)

                    # 只处理图注所在页及其之前的页
                    if shape_page > target_page:
                        continue

                    # 同页：只取图注上方
                    if shape_page == target_page:
                        visual_y = shape.Anchor.Information(wdVerticalPositionRelativeToPage)
                        if not (0 < visual_y < (bottom_limit - 2)):
                            continue

                    page_shapes_map.setdefault(shape_page, []).append(shape.Name)
                except:
                    pass

            if not page_shapes_map:
                continue

            # --------------------------------------------------
            # 【关键修改 ②】
            # 按页逐页生成图片
            # 每一页一张图，不拼接、不跨页合并
            # --------------------------------------------------
            for page_num in sorted(page_shapes_map.keys()):
                shapes = page_shapes_map[page_num]
                if not shapes:
                    continue

                try:
                    doc.Activate()
                    doc.Shapes.Range(shapes).Select()

                    force_clear_clipboard()
                    time.sleep(0.5)
                    word.Selection.CopyAsPicture()
                    time.sleep(1.5)

                    safe_filename = (
                        target_text
                        .replace(" ", "_")
                        .replace("/", "")
                        .replace(":", "")
                        .replace("\n", "")
                        + f"_p{page_num}.png"
                    )
                    save_path = str(out_dir / safe_filename)

                    if save_clipboard_to_image(save_path):
                        print(f"    > 保存成功: {safe_filename}")

                        image_map.setdefault(target_text, []).append(safe_filename)

                        # 删除当前页 Shapes
                        for shape_name in shapes:
                            try:
                                doc.Shapes(shape_name).Delete()
                                processed_shapes.add(shape_name)
                            except:
                                pass
                    else:
                        print(f"    [警告] 页面 {page_num} 保存失败")

                except Exception as e:
                    print(f"    [错误] 页面 {page_num}: {e}")

        # ================= 3. 保存映射文件 =================
        json_path = str(out_dir / "image_map.json")
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(image_map, f, ensure_ascii=False, indent=4)

        print(f"\n映射文件已生成: {json_path}")

        # ================= 4. 保存清理版文档 =================
        doc.SaveAs(str(out_doc))
        print(f"清理版文档已保存: {out_doc}")

        return str(out_doc), json_path, image_map

    finally:
        force_clear_clipboard()
        if doc:
            try:
                doc.Close(SaveChanges=0)
            except:
                pass


def _build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Extract docx Shapes to images and generate caption->image map.")
    parser.add_argument("--input", required=True, help="Input docx path")
    parser.add_argument("--output-dir", required=True, help="Directory to save images and image_map.json")
    parser.add_argument("--output-doc", default=None, help="Output cleaned docx path")
    parser.add_argument("--word-visible", dest="word_visible", action="store_true", default=True)
    parser.add_argument("--word-hidden", dest="word_visible", action="store_false")
    return parser


if __name__ == "__main__":
    args = _build_arg_parser().parse_args()
    extract_and_generate_map(
        args.input,
        args.output_dir,
        output_doc_path=args.output_doc,
        word_visible=args.word_visible,
    )
