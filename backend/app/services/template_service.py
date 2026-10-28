"""模板管理服务。

用文件系统 + 元数据 JSON 存储，不依赖数据库。
（数据库结构的调整需要单独批准，见 docs/REFACTOR_PLAN.md §8）

模板的作用：提供页眉页脚与正文样式基线。
原实现把模板处理散落在 service/converter.py 的 step0_replace_header_footer 中，
这里收敛为单一职责：把模板的页眉页脚"移植"到目标文档。
"""

from __future__ import annotations

import hashlib
import json
import posixpath
import re
import shutil
from datetime import datetime, timezone
from pathlib import Path

from lxml import etree

from app.core.config import get_settings
from app.core.constants import layout_rules
from app.core.errors import NotFoundError, ValidationError
from app.core.logging import get_logger
from app.core.namespaces import NSMAP, qn, w_attr
from app.domain.ooxml.package import DocxPackage, paragraph_text

logger = get_logger(__name__)

_META_FILENAME = "_meta.json"
_SLUG_PATTERN = re.compile(r"[^a-zA-Z0-9_-]+")

_HF_PREFIXES = ("word/header", "word/footer")


class TemplateService:
    """模板的注册、列举与页眉页脚移植。"""

    def __init__(self, root: Path | None = None) -> None:
        self.root = root or get_settings().template_dir
        self.root.mkdir(parents=True, exist_ok=True)
        self._meta_path = self.root / _META_FILENAME

    # ------------------------------------------------------------------
    # 元数据
    # ------------------------------------------------------------------
    def _load_meta(self) -> dict[str, dict]:
        if not self._meta_path.exists():
            return {}
        try:
            with self._meta_path.open("r", encoding="utf-8") as fp:
                return json.load(fp)
        except (json.JSONDecodeError, OSError) as exc:
            logger.warning("模板元数据损坏，已重置: %s", exc)
            return {}

    def _save_meta(self, meta: dict[str, dict]) -> None:
        tmp = self._meta_path.with_suffix(".tmp")
        with tmp.open("w", encoding="utf-8") as fp:
            json.dump(meta, fp, ensure_ascii=False, indent=2)
        tmp.replace(self._meta_path)

    @staticmethod
    def _make_id(name: str) -> str:
        slug = _SLUG_PATTERN.sub("-", name).strip("-").lower() or "template"
        digest = hashlib.sha1(name.encode("utf-8")).hexdigest()[:8]
        return f"{slug}-{digest}"

    # ------------------------------------------------------------------
    # 查询
    # ------------------------------------------------------------------
    def list_templates(self) -> list[dict]:
        meta = self._load_meta()
        items: list[dict] = []
        for template_id, info in meta.items():
            path = self.root / info.get("filename", "")
            if not path.exists():
                logger.warning("模板文件缺失，跳过: %s", template_id)
                continue
            items.append(
                {
                    "template_id": template_id,
                    "name": info.get("name", template_id),
                    "kind": info.get("kind", "content"),
                    "filename": info.get("filename", ""),
                    "size_bytes": path.stat().st_size,
                    "created_at": info.get("created_at", ""),
                    "is_default": bool(info.get("is_default", False)),
                    "description": info.get("description", ""),
                    "tags": info.get("tags", []),
                }
            )
        items.sort(key=lambda item: (not item["is_default"], item["created_at"]))
        return items

    def get_path(self, template_id: str) -> Path:
        meta = self._load_meta()
        info = meta.get(template_id)
        if info is None:
            raise NotFoundError(f"模板不存在: {template_id}")
        path = self.root / info.get("filename", "")
        if not path.exists():
            raise NotFoundError(f"模板文件已丢失: {template_id}")
        return path

    def builtin_content_template(self) -> Path | None:
        """内置正文模板。调用方没有另选模板时用它。"""
        name = layout_rules()["content_template"]
        path = get_settings().resources_dir / name
        if path.is_file():
            return path
        return None

    def default_template_path(self) -> Path | None:
        """返回默认模板路径；无默认时取第一个 content 模板。"""
        items = self.list_templates()
        for item in items:
            if item["is_default"]:
                return self.root / item["filename"]
        for item in items:
            if item["kind"] == "content":
                return self.root / item["filename"]
        return None

    # ------------------------------------------------------------------
    # 注册
    # ------------------------------------------------------------------
    def register(
        self,
        name: str,
        filename: str,
        data: bytes,
        kind: str = "content",
        description: str = "",
        is_default: bool = False,
    ) -> dict:
        if not filename.lower().endswith(".docx"):
            raise ValidationError("模板必须是 .docx 格式")

        template_id = self._make_id(name)
        stored_name = f"{template_id}.docx"
        target = self.root / stored_name
        target.write_bytes(data)

        # 校验能被解析，避免注册损坏模板
        try:
            with DocxPackage.open(target):
                pass
        except Exception as exc:
            target.unlink(missing_ok=True)
            raise ValidationError(f"模板文件无法解析: {exc}") from exc

        meta = self._load_meta()
        meta[template_id] = {
            "name": name,
            "filename": stored_name,
            "kind": kind,
            "description": description,
            "is_default": is_default,
            "created_at": datetime.now(timezone.utc).isoformat(),
            "tags": [],
        }

        if is_default:
            for key in meta:
                if key != template_id:
                    meta[key]["is_default"] = False

        self._save_meta(meta)
        logger.info("注册模板: %s (%s)", template_id, name)
        return next(item for item in self.list_templates() if item["template_id"] == template_id)

    def delete(self, template_id: str) -> None:
        meta = self._load_meta()
        info = meta.pop(template_id, None)
        if info is None:
            raise NotFoundError(f"模板不存在: {template_id}")
        (self.root / info.get("filename", "")).unlink(missing_ok=True)
        self._save_meta(meta)

    # ------------------------------------------------------------------
    # 页眉页脚移植
    # ------------------------------------------------------------------
    def apply_header_footer(self, target: DocxPackage, template_path: Path) -> list[str]:
        """把模板的页眉页脚内容移植到目标文档。

        策略：按"同类型、同序号"替换。目标文档第 N 个 header 部件接收模板第 N 个
        header 部件的内容，footer 同理。这样目标文档原有的 rels 结构保持有效，
        不需要重建引用关系。

        返回告警列表（媒体冲突等），调用方需将其透出给用户，不得静默吞掉。
        """
        warnings: list[str] = []

        try:
            template = DocxPackage.open(template_path)
        except Exception as exc:
            warnings.append(f"模板无法解析，已跳过页眉页脚移植: {exc}")
            return warnings

        with template:
            for prefix in _HF_PREFIXES:
                target_parts = self._referenced_parts(target, prefix)
                template_parts = self._referenced_parts(template, prefix)

                if not target_parts or not template_parts:
                    continue

                for index, target_part in enumerate(target_parts):
                    if index >= len(template_parts):
                        break
                    source_part = template_parts[index]
                    warnings.extend(
                        self._port_part(target, template, source_part, target_part)
                    )

        return warnings

    def apply_body_header(
        self,
        target: DocxPackage,
        template_path: Path,
        cover_sections: int,
    ) -> list[str]:
        """正文模板的页眉页脚套到封面第三页、目录和正文。

        封面第一、二页留在第一节，继续用封面自己的页眉页脚。
        封面第三页（设计人员）和目录用同一套页眉页脚，并且不编页码。
        目录之后的正文从 1 开始连续编页码。
        """
        warnings: list[str] = []
        try:
            template = DocxPackage.open(template_path)
        except Exception as exc:
            warnings.append(f"正文模板无法解析，已跳过页眉: {exc}")
            return warnings

        toc_section = self._toc_section_index(target)
        if cover_sections >= 2:
            plain_from = cover_sections - 1
        else:
            plain_from = cover_sections
        numbered_from = toc_section + 1 if toc_section >= plain_from else plain_from

        with template:
            source_part = self._logo_header_part(template)
            if source_part is None:
                warnings.append("正文模板没有带图标的页眉")
                return warnings

            part_name = "word/header_body.xml"
            warnings.extend(self._port_part(target, template, source_part, part_name))
            header_rel = self._ensure_header_relationship(target, part_name)
            self._bind_reference_range(target, "w:headerReference", header_rel, plain_from, None)
            self._drop_title_page(target, plain_from)

            footer_part = self._page_number_footer(template)
            if footer_part is not None:
                warnings.extend(self._port_part(target, template, footer_part, "word/footer_body.xml"))
                warnings.extend(self._port_part(target, template, footer_part, "word/footer_front.xml"))
                self._strip_page_fields(target.xml("word/footer_front.xml"))
                front_rel = self._ensure_part_relationship(
                    target,
                    "word/footer_front.xml",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
                    "rIdFrontFooter",
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
                )
                body_rel = self._ensure_part_relationship(
                    target,
                    "word/footer_body.xml",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
                    "rIdBodyFooter",
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
                )
                self._bind_reference_range(
                    target, "w:footerReference", front_rel, plain_from, numbered_from
                )
                self._bind_reference_range(
                    target, "w:footerReference", body_rel, numbered_from, None
                )
            self._copy_header_distance(target, template, plain_from)
            self._restart_page_numbers(target, numbered_from)
        return warnings

    def fill_header_slots(self, target: DocxPackage, left: str, right: str) -> int:
        """页眉从左到右两个文字节点：项目名称、文档类型。只改带横线的正文页眉。"""
        if not left and not right:
            return 0

        filled = 0
        for name in target.part_names:
            if not (name.startswith("word/header") and name.endswith(".xml")):
                continue
            root = target.xml(name)
            if not root.xpath(".//w:pBdr", namespaces=NSMAP):
                continue
            nodes = [node for node in root.iter(qn("w:t")) if (node.text or "").strip()]
            if len(nodes) < 2:
                continue
            if left:
                nodes[0].text = left
                nodes[0].set(qn("xml:space"), "preserve")
            if right:
                nodes[1].text = right
                nodes[1].set(qn("xml:space"), "preserve")
            filled += 1
        if filled:
            logger.info("页眉已填入: %s | %s", left, right)
        return filled

    def _ensure_header_relationship(self, target: DocxPackage, part_name: str) -> str:
        return self._ensure_part_relationship(
            target,
            part_name,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
            "rIdBodyHeader",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
        )

    def _ensure_part_relationship(
        self,
        target: DocxPackage,
        part_name: str,
        rel_type: str,
        rel_id_base: str,
        content_type: str,
    ) -> str:
        rels_name = "word/_rels/document.xml.rels"
        rels = target.xml(rels_name)
        target_rel = posixpath.relpath(part_name, "word")
        for rel in rels.findall(qn("pr:Relationship")):
            if rel.get("Target") == target_rel and rel.get("Type") == rel_type:
                return rel.get("Id") or ""

        used = {rel.get("Id") or "" for rel in rels.findall(qn("pr:Relationship"))}
        rel_id = rel_id_base
        suffix = 1
        while rel_id in used:
            suffix += 1
            rel_id = f"{rel_id_base}{suffix}"
        self._append_rel(rels, rel_id, rel_type, target_rel, False)
        self._ensure_override_content_type(target, part_name, content_type)
        return rel_id

    @staticmethod
    def _ensure_override_content_type(target: DocxPackage, part_name: str, content_type: str) -> None:
        content_types = target.xml("[Content_Types].xml")
        part = "/" + part_name
        for override in content_types.findall(qn("ct:Override")):
            if override.get("PartName") == part:
                return
        node = etree.SubElement(content_types, qn("ct:Override"))
        node.set("PartName", part)
        node.set("ContentType", content_type)

    @staticmethod
    def _toc_section_index(pkg: DocxPackage) -> int:
        """目录标题所在的节序号。找不到时返回 0。"""
        index = 0
        for child in pkg.body:
            if child.tag == qn("w:p"):
                compact = re.sub(r"\s+", "", paragraph_text(child))
                if compact == "目录":
                    return index
                if child.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None:
                    index += 1
            elif child.tag == qn("w:sectPr"):
                index += 1
        return 0

    @staticmethod
    def _drop_title_page(target: DocxPackage, start: int) -> None:
        for index, sect_pr in enumerate(target.section_properties()):
            if index < start:
                continue
            title_pg = sect_pr.find(qn("w:titlePg"))
            if title_pg is not None:
                sect_pr.remove(title_pg)

    @staticmethod
    def _bind_reference_range(
        target: DocxPackage,
        tag: str,
        rel_id: str,
        start: int,
        end: int | None,
    ) -> None:
        for index, sect_pr in enumerate(target.section_properties()):
            if index < start or (end is not None and index >= end):
                continue
            for ref in list(sect_pr.findall(qn(tag))):
                sect_pr.remove(ref)
            ref = etree.Element(qn(tag))
            ref.set(w_attr("type"), "default")
            ref.set(qn("r:id"), rel_id)
            sect_pr.insert(0, ref)

    @staticmethod
    def _strip_page_fields(root: etree._Element) -> None:
        """去掉页脚里的 PAGE 域，目录及之前的页不显示页码。"""
        for paragraph in root.iter(qn("w:p")):
            runs = [child for child in paragraph if child.tag == qn("w:r")]
            removing = False
            drop: list[etree._Element] = []
            field_runs: list[etree._Element] = []
            is_page = False
            for run in runs:
                fld = run.find(qn("w:fldChar"))
                instr = run.find(qn("w:instrText"))
                if fld is not None and (fld.get(w_attr("fldCharType")) or "") == "begin":
                    removing = True
                    is_page = False
                    field_runs = [run]
                    continue
                if removing:
                    field_runs.append(run)
                    if instr is not None and instr.text and "PAGE" in instr.text.upper():
                        is_page = True
                    if fld is not None and (fld.get(w_attr("fldCharType")) or "") == "end":
                        if is_page:
                            drop.extend(field_runs)
                        removing = False
                        field_runs = []
            for run in drop:
                parent = run.getparent()
                if parent is not None:
                    parent.remove(run)

    @staticmethod
    def _page_number_footer(template: DocxPackage) -> str | None:
        """正文模板里带 PAGE 域的页脚，页码靠它自动编号。"""
        for name in template.part_names:
            if not (name.startswith("word/footer") and name.endswith(".xml")):
                continue
            root = template.xml(name)
            for instr in root.iter(qn("w:instrText")):
                if instr.text and "PAGE" in instr.text.upper():
                    return name
        return None

    @staticmethod
    def _copy_header_distance(target: DocxPackage, template: DocxPackage, reference_index: int) -> None:
        """页眉页脚位置以封面第三页为准，只改后面各节的位置。

        第三页先对齐正文模板的页眉距、页脚距，它自己的上下边距保留。
        后面的节（目录、正文、横向页）只同步页眉距、页脚距和上下边距。
        左右边距不动，页眉页脚里的文字也不动。
        """
        sections = target.section_properties()
        if reference_index < 0 or reference_index >= len(sections):
            return
        reference = sections[reference_index].find(qn("w:pgMar"))
        if reference is None:
            return

        template_sections = template.section_properties()
        if template_sections:
            source_mar = template_sections[0].find(qn("w:pgMar"))
            if source_mar is not None:
                for attr in ("header", "footer"):
                    value = source_mar.get(w_attr(attr))
                    if value:
                        reference.set(w_attr(attr), value)

        values = {
            attr: reference.get(w_attr(attr))
            for attr in ("header", "footer", "top", "bottom")
        }
        for index, sect_pr in enumerate(sections):
            if index <= reference_index:
                continue
            margin = sect_pr.find(qn("w:pgMar"))
            if margin is None:
                continue
            for attr, value in values.items():
                if value:
                    margin.set(w_attr(attr), value)

    @staticmethod
    def _restart_page_numbers(target: DocxPackage, skip_sections: int) -> None:
        """正文第一节从 1 开始，后面的节连续编号。"""
        first = True
        for index, sect_pr in enumerate(target.section_properties()):
            numbering = sect_pr.find(qn("w:pgNumType"))
            if index < skip_sections:
                # 目录及之前不编页码，清掉可能残留的起始页码
                if numbering is not None:
                    sect_pr.remove(numbering)
                continue
            if first:
                if numbering is None:
                    numbering = etree.SubElement(sect_pr, qn("w:pgNumType"))
                numbering.set(w_attr("start"), "1")
                first = False
            elif numbering is not None:
                sect_pr.remove(numbering)

    @staticmethod
    def _logo_header_part(template: DocxPackage) -> str | None:
        for name in template.part_names:
            if not (name.startswith("word/header") and name.endswith(".xml")):
                continue
            root = template.xml(name)
            if root.xpath(".//a:blip", namespaces=NSMAP):
                return name
        return None

    @staticmethod
    def _referenced_parts(pkg: DocxPackage, prefix: str) -> list[str]:
        """按文档顺序返回被 sectPr 引用的页眉/页脚部件名。

        顺序依据 headerReference/footerReference 的 type 属性：
        default(1) -> first(2) -> even(3)，与 Word 的约定一致。
        """
        order = {"default": 0, "first": 1, "even": 2}
        refs: list[tuple[int, str]] = []

        root = pkg.xml("word/document.xml")
        rels = pkg.xml("word/_rels/document.xml.rels") if pkg.has("word/_rels/document.xml.rels") else None
        if rels is None:
            return []

        rel_map = {
            rel.get("Id"): rel.get("Target")
            for rel in rels.findall(qn("pr:Relationship"))
        }

        seen: set[str] = set()
        for sect_pr in root.xpath(".//w:sectPr", namespaces=NSMAP):
            for ref in sect_pr:
                if ref.tag not in (qn("w:headerReference"), qn("w:footerReference")):
                    continue
                target = rel_map.get(ref.get(qn("r:id")))
                if not target:
                    continue
                part = posixpath.normpath(posixpath.join("word", target))
                if not part.startswith(prefix) or part in seen:
                    continue
                seen.add(part)
                kind = ref.get(qn("w:type")) or "default"
                refs.append((order.get(kind, 9), part))

        refs.sort(key=lambda item: item[0])
        return [part for _, part in refs]

    def _port_part(
        self,
        target: DocxPackage,
        template: DocxPackage,
        source_part: str,
        target_part: str,
    ) -> list[str]:
        """把单个页眉/页脚部件及其依赖迁入目标文档。"""
        warnings: list[str] = []
        target.add_part(target_part, template.read(source_part))

        source_rels_name = f"word/_rels/{posixpath.basename(source_part)}.rels"
        target_rels_name = f"word/_rels/{posixpath.basename(target_part)}.rels"

        if not template.has(source_rels_name):
            target.remove_part(target_rels_name)
            return warnings

        rels_root = template.xml(source_rels_name)
        new_root = etree.Element(
            qn("pr:Relationships"),
            nsmap={None: "http://schemas.openxmlformats.org/package/2006/relationships"},
        )

        for rel in rels_root.findall(qn("pr:Relationship")):
            rel_type = rel.get("Type") or ""
            rel_target = rel.get("Target") or ""
            external = (rel.get("TargetMode") or "") == "External"

            if external:
                self._append_rel(new_root, rel.get("Id") or "", rel_type, rel_target, True)
                continue

            source_media = posixpath.normpath(posixpath.join("word", rel_target))
            if not template.has(source_media):
                warnings.append(f"模板部件 {source_part} 引用了缺失文件: {rel_target}")
                continue

            data = template.read(source_media)
            new_name = self._unique_part_name(target, source_media, data, warnings)
            relative = posixpath.relpath(new_name, "word")
            self._append_rel(new_root, rel.get("Id") or "", rel_type, relative, False)

        target.replace_xml_bytes(
            target_rels_name,
            etree.tostring(new_root, xml_declaration=True, encoding="UTF-8"),
        )
        return warnings

    @staticmethod
    def _append_rel(
        root: etree._Element,
        rel_id: str,
        rel_type: str,
        target: str,
        external: bool,
    ) -> None:
        if not rel_id:
            return
        rel = etree.SubElement(root, qn("pr:Relationship"))
        rel.set("Id", rel_id)
        rel.set("Type", rel_type)
        rel.set("Target", target)
        if external:
            rel.set("TargetMode", "External")

    def _unique_part_name(
        self,
        target: DocxPackage,
        source_name: str,
        data: bytes,
        warnings: list[str],
    ) -> str:
        """为迁入的媒体文件找不冲突的名字；内容相同则直接复用。"""
        if target.has(source_name):
            if target.read(source_name) == data:
                return source_name
            stem, dot, ext = source_name.rpartition(".")
            directory = posixpath.dirname(source_name)
            for index in range(1, 1000):
                candidate = f"{posixpath.join(directory, posixpath.basename(stem))}_{index}{dot}{ext}"
                if not target.has(candidate):
                    warnings.append(
                        f"媒体文件重名，已另存为 {posixpath.basename(candidate)}"
                    )
                    target.add_part(candidate, data)
                    self._ensure_media_content_type(target, candidate)
                    return candidate
            raise ValidationError(f"无法为 {source_name} 分配不冲突的文件名")

        target.add_part(source_name, data)
        self._ensure_media_content_type(target, source_name)
        return source_name

    @staticmethod
    def _ensure_media_content_type(target: DocxPackage, part_name: str) -> None:
        ext = part_name.rsplit(".", 1)[-1].lower()
        content_types = {
            "png": "image/png",
            "jpg": "image/jpeg",
            "jpeg": "image/jpeg",
            "gif": "image/gif",
            "wmf": "image/x-wmf",
            "emf": "image/x-emf",
            "tif": "image/tiff",
            "tiff": "image/tiff",
        }
        content_type = content_types.get(ext)
        if content_type is None or not target.has("[Content_Types].xml"):
            return
        root = target.xml("[Content_Types].xml")
        for node in root.findall(qn("ct:Default")):
            if (node.get("Extension") or "").lower() == ext:
                return
        default = etree.SubElement(root, qn("ct:Default"))
        default.set("Extension", ext)
        default.set("ContentType", content_type)
