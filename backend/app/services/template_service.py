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
from app.core.errors import NotFoundError, ValidationError
from app.core.logging import get_logger
from app.core.namespaces import NSMAP, qn
from app.domain.ooxml.package import DocxPackage

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
                    return candidate
            raise ValidationError(f"无法为 {source_name} 分配不冲突的文件名")

        target.add_part(source_name, data)
        return source_name
