import win32com.client
import os
import time
import ctypes
from PIL import ImageGrab


# ================= 工具函数 =================
def force_clear_clipboard():
    """强制清空系统剪贴板"""
    try:
        # 尝试打开剪贴板
        if ctypes.windll.user32.OpenClipboard(None):
            # 清空
            ctypes.windll.user32.EmptyClipboard()
            # 关闭
            ctypes.windll.user32.CloseClipboard()
    except Exception:
        pass


def save_clipboard_to_image(save_path):
    """保存剪贴板内容为图片，带重试机制"""
    # 重试 20 次，给 Word 渲染大图留足时间
    for i in range(20):
        try:
            img = ImageGrab.grabclipboard()
            if img is not None:
                # 排除文件列表，确保是图片
                if isinstance(img, list): return False
                img.save(save_path, 'PNG')
                return True
        except:
            time.sleep(0.5)
        time.sleep(0.5)
    return False


# ================= 核心逻辑 =================
def extract_aggressive_v3(src_path, output_dir):
    # 1. 路径标准化处理
    base_dir = os.getcwd()
    if not os.path.isabs(src_path): src_path = os.path.abspath(os.path.join(base_dir, src_path))
    if not os.path.isabs(output_dir): output_dir = os.path.abspath(os.path.join(base_dir, output_dir))

    if not os.path.exists(output_dir): os.makedirs(output_dir)

    print(f"处理文档: {src_path}")
    print(f"输出目录: {output_dir}")

    # 启动 Word
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True  # 建议保持 True 以便观察
    word.DisplayAlerts = 0  # 暂时屏蔽弹窗

    # Word 常量
    wdVerticalPositionRelativeToPage = 6
    wdActiveEndPageNumber = 3

    doc = None

    try:
        doc = word.Documents.Open(src_path)

        # --- 第一步：扫描所有图注 ---
        captions = []
        print("正在扫描图注位置...")

        for p in doc.Paragraphs:
            # 过滤长段落
            if len(p.Range.Text) > 100: continue
            text = p.Range.Text.strip()

            # 【匹配规则】以“图”开头
            if text.startswith("图") and len(text) > 2:
                try:
                    visual_top = p.Range.Information(wdVerticalPositionRelativeToPage)
                    page_num = p.Range.Information(wdActiveEndPageNumber)
                    captions.append({
                        "text": text,
                        "visual_top": visual_top,
                        "page": page_num
                    })
                    print(f"  - 锁定锚点: [{text}] (第{page_num}页)")
                except:
                    pass

        # --- 第二步：遍历每个图注进行处理 ---
        processed_shapes = set()

        for cap in captions:
            target_text = cap['text']
            target_page = cap['page']
            bottom_limit = cap['visual_top']

            print(f"--- 正在处理: {target_text} ---")

            shapes_to_select = []

            # 全文档扫描形状
            for shape in doc.Shapes:
                try:
                    if shape.Name in processed_shapes: continue

                    # A. 必须在同一页
                    if shape.Anchor.Information(wdActiveEndPageNumber) != target_page:
                        continue

                    # B. 必须在图注上方 (0 < 高度 < 图注高度)
                    # 只要是这一页上方的，全部选中
                    visual_y = shape.Anchor.Information(wdVerticalPositionRelativeToPage)
                    if 0 < visual_y < (bottom_limit - 2):
                        shapes_to_select.append(shape.Name)
                except:
                    pass

            if not shapes_to_select:
                print("    [警告] 本页上方未发现形状。")
                continue

            print(f"    > 选中 {len(shapes_to_select)} 个零件，准备提取...")

            try:
                # 激活文档
                doc.Activate()
                # 批量选中
                doc.Shapes.Range(shapes_to_select).Select()

                # 记录已处理
                for n in shapes_to_select: processed_shapes.add(n)

                # 1. 强制清空剪贴板（防止残留）
                force_clear_clipboard()
                time.sleep(0.5)

                # 2. 复制为图片
                word.Selection.CopyAsPicture()
                print("    > 已复制，正在渲染...")
                time.sleep(1.5)

                # 3. 保存图片
                safe_name = target_text.replace(" ", "_").replace("/", "").replace(":", "").replace("\n", "")
                save_path = os.path.join(output_dir, f"{safe_name}.png")

                if save_clipboard_to_image(save_path):
                    print(f"    > [成功] 图片已保存: {save_path}")

                    # 4. 逐个删除原部件 (强力删除模式)
                    print(f"    > 正在清除原图表组件...")
                    delete_count = 0
                    for shape_name in shapes_to_select:
                        try:
                            doc.Shapes(shape_name).Delete()
                            delete_count += 1
                        except:
                            pass
                    print(f"    > 清理完成: {delete_count}/{len(shapes_to_select)}")
                else:
                    print(f"    [失败] 图片保存失败，跳过删除步骤")

            except Exception as e:
                print(f"    [错误] 操作异常: {e}")

        # --- 第三步：另存为清理后的文档 ---
        new_doc_path = os.path.splitext(src_path)[0] + "_clean.docx"
        doc.SaveAs(new_doc_path)
        print(f"\n全部流程结束！")
        print(f"清理版文档已保存: {new_doc_path}")

    except Exception as e:
        print(f"程序运行出错: {e}")

    finally:
        # ==========================================
        #        关键修改：解决退出弹窗问题
        # ==========================================
        print("正在清理资源...")

        # 1. 【核心】退出前强制清空剪贴板
        # 只要剪贴板是空的，Word 就不会问"是否保留"
        force_clear_clipboard()

        # 2. 关闭文档 (不保存，因为前面已经 SaveAs 了)
        if doc:
            try:
                doc.Close(SaveChanges=0)
            except:
                pass

        # 3. 退出 Word
        if word:
            try:
                word.Quit()
            except:
                pass
            # 释放 COM 对象
            del word

        print("资源清理完毕。")


# ================= 配置区 =================
src_file = r"../docx/二烧翻车机可研汇总.docx"
out_folder = r"../docx/extracted_images_v3"

if __name__ == "__main__":
    extract_aggressive_v3(src_file, out_folder)