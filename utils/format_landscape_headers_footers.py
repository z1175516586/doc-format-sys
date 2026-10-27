# test_format_headers_footers.py
import os
import time
import zipfile
from lxml import etree

class DocxProcessor:
    def __init__(self):
        # 固定输入文件
        self.input_file = r"C:\DevOps\doc-convertor\result\aaa.docx"

        # 输出文件夹
        timestamp = time.strftime("%H%M%S")
        self.output_dir = r"C:\DevOps\doc-convertor\result\2026年1月23日"
        os.makedirs(self.output_dir, exist_ok=True)

        # 输出文件路径
        self.output_file = os.path.join(self.output_dir, f"converted_{timestamp}.docx")

    def step14_format_landscape_headers_footers(self, output_file: str):
        """
        步骤14：以【最后一个竖版小节】的页眉页脚为模板，
               覆盖其后所有【横板小节】所引用的 header/footer XML 文件内容
        """

        import zipfile
        from lxml import etree
        from copy import deepcopy

        W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        NSMAP = {"w": W_NS, "r": R_NS}

        def log(msg):
            print(f"[Step14] {msg}")

        def is_landscape(sectPr):
            pgSz = sectPr.find("w:pgSz", NSMAP)
            return pgSz is not None and pgSz.get(f"{{{W_NS}}}orient") == "landscape"

        try:
            # ---------- 1. 解包 docx ----------
            with zipfile.ZipFile(output_file, "r") as zin:
                files = {name: zin.read(name) for name in zin.namelist()}

            # ---------- 2. 解析 document.xml ----------
            doc_tree = etree.fromstring(files["word/document.xml"])
            body = doc_tree.find("w:body", NSMAP)
            sectPrs = body.xpath(".//w:sectPr", namespaces=NSMAP)

            log(f"检测到 sectPr 数量: {len(sectPrs)}")

            def has_default_header_footer(sect):
                has_header = False
                has_footer = False

                for h in sect.findall("w:headerReference", NSMAP):
                    if h.get(f"{{{W_NS}}}type") == "default":
                        has_header = True
                        break

                for f in sect.findall("w:footerReference", NSMAP):
                    if f.get(f"{{{W_NS}}}type") == "default":
                        has_footer = True
                        break

                return has_header and has_footer

            # ---------- 3. 找到最后一个“可用”的竖版小节 ----------
            last_portrait_idx = None
            last_portrait_sect = None

            for i, sect in enumerate(sectPrs):
                if is_landscape(sect):
                    continue

                if not has_default_header_footer(sect):
                    log(f"跳过第 {i} 节：竖版但无 default header/footer")
                    continue

                last_portrait_idx = i
                last_portrait_sect = sect
                log(f"候选竖版小节：第 {i} 节（含 default header/footer）")
                break
            last_portrait_idx = last_portrait_idx + 1
            if last_portrait_idx is None:
                raise RuntimeError("未找到包含 default header/footer 的竖版小节，无法作为复制源")

            portrait_sect = sectPrs[last_portrait_idx]
            landscape_sects = sectPrs[last_portrait_idx + 1:]

            log(f"最后一个竖版小节索引: {last_portrait_idx}")
            log(f"需处理横板小节数量: {len(landscape_sects)}")

            if not landscape_sects:
                return True

            # ---------- 输出最后一个竖版小节的 sectPr ----------
            log("【DEBUG】最后一个竖版小节 sectPr XML 内容如下：")

            sect_xml_str = etree.tostring(
                portrait_sect,
                pretty_print=True,
                encoding="unicode"
            )

            print(sect_xml_str)

            # ---------- 4. 解析 document.xml.rels ----------
            REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
            REL_NSMAP = {"rel": REL_NS}

            rels_tree = etree.fromstring(files["word/_rels/document.xml.rels"])

            rid_to_target = {
                rel.get("Id"): "word/" + rel.get("Target")
                for rel in rels_tree.findall("rel:Relationship", namespaces=REL_NSMAP)
            }

            # ---------- DEBUG: 仅解析竖版小节 default header/footer ----------
            def get_default_header_footer_target(sect):
                default_header_rid = None
                default_header_target = None
                default_footer_rid = None
                default_footer_target = None

                # 查找 default header
                for el in sect.findall("w:headerReference", NSMAP):
                    w_type = el.get(f"{{{W_NS}}}type")
                    if w_type == "default":
                        rid = el.get(f"{{{R_NS}}}id")
                        target = rid_to_target.get(rid)
                        default_header_rid = rid
                        default_header_target = target
                        break  # 只取第一个 default header

                # 查找 default footer
                for el in sect.findall("w:footerReference", NSMAP):
                    w_type = el.get(f"{{{W_NS}}}type")
                    if w_type == "default":
                        rid = el.get(f"{{{R_NS}}}id")
                        target = rid_to_target.get(rid)
                        default_footer_rid = rid
                        default_footer_target = target
                        break  # 只取第一个 default footer

                return default_header_rid, default_header_target, default_footer_rid, default_footer_target

            # ---------- 输出竖版小节 default header/footer 对应的文件 ----------
            default_header_rid, default_header_path, default_footer_rid, default_footer_path = \
                get_default_header_footer_target(portrait_sect)

            if not default_header_path and not default_footer_path:
                log("⚠️ 未找到竖版小节 default header/footer")
                return False

            log(f"[DEBUG] 使用的竖版 header 文件为: {default_header_path}")
            log(f"[DEBUG] 使用的竖版 footer 文件为: {default_footer_path}")

            # ---------- 打印竖版 default header/footer 的全部内容 ----------
            if default_header_path and default_header_path in files:
                header_content = files[default_header_path].decode("utf-8")
                print(f"\n【竖版 default header 内容 - {default_header_path}】\n")
                print(header_content)
            else:
                print(f"\n⚠️ default_header_path 无效或文件不存在: {default_header_path}")

            if default_footer_path and default_footer_path in files:
                footer_content = files[default_footer_path].decode("utf-8")
                print(f"\n【竖版 default footer 内容 - {default_footer_path}】\n")
                print(footer_content)
            else:
                print(f"\n⚠️ default_footer_path 无效或文件不存在: {default_footer_path}")

            # ---------- 复制竖版 pgMar 到横板 sectPr ----------
            portrait_pgMar = portrait_sect.find("w:pgMar", NSMAP)
            if portrait_pgMar is not None:
                log("【DEBUG】最后一个竖版 sectPr 的 pgMar 属性：")
                print(dict(portrait_pgMar.attrib))

            # ---------- 覆盖横板小节 header/footer ----------
            for idx, sect in enumerate(landscape_sects, start=last_portrait_idx + 1):
                log(f"处理第 {idx} 节（横板）")

                # 复制 pgMar
                sect_pgMar = sect.find("w:pgMar", NSMAP)
                if sect_pgMar is not None:
                    log("横板小节原 pgMar 属性：")
                    print(dict(sect_pgMar.attrib))
                    # 替换为竖版 pgMar
                    sect.replace(sect_pgMar, deepcopy(portrait_pgMar))
                else:
                    log("横板小节没有 pgMar，添加竖版 pgMar")
                    sect.append(deepcopy(portrait_pgMar))
                # 打印更新后的属性
                sect_pgMar_new = sect.find("w:pgMar", NSMAP)
                log("横板小节更新后的 pgMar 属性：")
                print(dict(sect_pgMar_new.attrib))

                # 查找该横板节引用的 header/footer
                headers = [rid_to_target[el.get(f"{{{R_NS}}}id")]
                           for el in sect.findall("w:headerReference", NSMAP)
                           if el.get(f"{{{R_NS}}}id") in rid_to_target]
                footers = [rid_to_target[el.get(f"{{{R_NS}}}id")]
                           for el in sect.findall("w:footerReference", NSMAP)
                           if el.get(f"{{{R_NS}}}id") in rid_to_target]

                # 覆盖 header 文件
                for path in headers:
                    if default_header_path and default_header_path in files:
                        files[path] = files[default_header_path]  # 直接覆盖
                        log(f"✔ 覆盖横板 header 文件: {path}")

                # 覆盖 footer 文件
                for path in footers:
                    if default_footer_path and default_footer_path in files:
                        footer_xml = etree.fromstring(files[default_footer_path])

                        # 找到所有文本框 <v:shape>
                        for shape in footer_xml.findall(".//v:shape",
                                                        namespaces={"v": "urn:schemas-microsoft-com:vml"}):
                            style = shape.get("style", "")

                            # 删除 margin-left 属性
                            import re
                            style = re.sub(r"margin-left\s*:[^;]+;?", "", style, flags=re.IGNORECASE)

                            # 删除任何空格和多余分号
                            style = style.strip().rstrip(";") + ";"

                            # 添加居中属性
                            style += "mso-position-horizontal:center;"
                            style += "mso-position-horizontal-relative:margin;"

                            shape.set("style", style)

                        # 写回文件
                        files[path] = etree.tostring(
                            footer_xml,
                            xml_declaration=True,
                            encoding="UTF-8",
                            standalone="yes"
                        )
                        log(f"✔ 覆盖横板 footer 文件（文本框水平居中）: {path}")

            # ---------- 7. 写回 document.xml ----------
            files["word/document.xml"] = etree.tostring(
                doc_tree,
                xml_declaration=True,
                encoding="UTF-8",
                standalone="yes"
            )

            # ---------- 7. 写回 docx ----------
            with zipfile.ZipFile(output_file, "w", zipfile.ZIP_DEFLATED) as zout:
                for name, data in files.items():
                    zout.writestr(name, data)

            log("✅ 横板页眉页脚已真正覆盖生效")
            return True


        except Exception as e:
            log(f"❌ 失败: {e}")
            return False


if __name__ == "__main__":
    processor = DocxProcessor()
    # 复制 input 文件到 output_file 做测试
    import shutil
    shutil.copyfile(processor.input_file, processor.output_file)
    print(f"测试文件已复制到: {processor.output_file}")

    success = processor.step14_format_landscape_headers_footers(processor.output_file)
    if success:
        print("页眉页脚处理完成 ✅")
    else:
        print("页眉页脚处理失败 ⚠️")
