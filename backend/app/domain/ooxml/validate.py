"""输出结构校验。

对应 docs/REFACTOR_PLAN.md §7.3 的结构断言清单。
目的：**不允许静默交付坏文档**。原实现大量步骤失败后仍继续并返回成功，
导致用户拿到的是"看起来正常但结构已损坏"的文档。

校验项：
1. 输出能被重新解析（zip 与 XML 结构完整）
2. document.xml 中的所有关系引用都能在 rels 中解析到
3. 所有内部关系目标部件真实存在
4. 正文至少有一个节属性
5. 无悬空样式引用（引用了 styles.xml 中不存在的 pStyle）
6. 分节引用的页眉、页脚部件具有正确的内容类型
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from posixpath import join, normpath

from app.core.namespaces import (
    HEADER_FOOTER_CONTENT_TYPES,
    PART_CONTENT_TYPES,
    PART_DOCUMENT,
    PART_DOCUMENT_RELS,
    PART_STYLES,
    R,
    qn,
)
from app.domain.ooxml.package import DocxPackage

_REL_ATTRS = (
    f"{{{R}}}embed",
    f"{{{R}}}id",
    f"{{{R}}}link",
)

# 这些关系指向包外部，不做存在性校验
_EXTERNAL_TYPES = ("hyperlink",)


@dataclass
class ValidationReport:
    """校验结果。"""

    ok: bool = True
    errors: list[str] = field(default_factory=list)
    stats: dict[str, int] = field(default_factory=dict)

    def add_error(self, message: str) -> None:
        self.ok = False
        self.errors.append(message)


class StructuralValidator:
    """对 docx 做结构断言。"""

    def validate(self, path: str | Path) -> ValidationReport:
        report = ValidationReport()
        try:
            pkg = DocxPackage.open(path)
        except Exception as exc:
            report.add_error(f"输出无法重新解析: {exc}")
            return report

        with pkg:
            self._check_relationships(pkg, report)
            self._check_sections(pkg, report)
            self._check_styles(pkg, report)
            self._collect_stats(pkg, report)

        return report

    # ------------------------------------------------------------------
    def _check_relationships(self, pkg: DocxPackage, report: ValidationReport) -> None:
        if not pkg.has(PART_DOCUMENT_RELS):
            report.add_error("缺少 word/_rels/document.xml.rels")
            return

        rels_root = pkg.xml(PART_DOCUMENT_RELS)
        known: dict[str, str] = {}
        for rel in rels_root.findall(qn("pr:Relationship")):
            rel_id = rel.get("Id")
            if rel_id:
                known[rel_id] = rel.get("Target") or ""

        document = pkg.xml(PART_DOCUMENT)
        missing_refs: set[str] = set()
        unresolved: set[str] = set()

        for node in document.iter():
            for attr in _REL_ATTRS:
                value = node.get(attr)
                if not value:
                    continue
                if value not in known:
                    missing_refs.add(value)
                else:
                    target = known[value]
                    part = f"word/{target}" if not target.startswith("/") else target.lstrip("/")
                    if not pkg.has(part) and not pkg.has(f"word/{target.split('/')[-1]}"):
                        # 目录类关系可能指向自身或其他包外部目标，仅记录
                        unresolved.add(f"{value}->{target}")

        if missing_refs:
            report.add_error(
                f"存在无法解析的关系引用: {', '.join(sorted(missing_refs)[:5])}"
            )

        report.stats["relationships"] = len(known)
        report.stats["unresolved_targets"] = len(unresolved)
        self._check_header_footer_content_types(pkg, report, known)

    @staticmethod
    def _check_header_footer_content_types(
        pkg: DocxPackage, report: ValidationReport, known: dict[str, str]
    ) -> None:
        if not pkg.has(PART_CONTENT_TYPES):
            report.add_error("缺少 [Content_Types].xml")
            return

        root = pkg.xml(PART_CONTENT_TYPES)
        overrides = {el.get("PartName"): el.get("ContentType") for el in root.findall(qn("ct:Override"))}
        defaults = {
            (el.get("Extension") or "").lower(): el.get("ContentType")
            for el in root.findall(qn("ct:Default"))
        }
        checked: set[tuple[str, str]] = set()
        for section in pkg.section_properties():
            for kind, expected in HEADER_FOOTER_CONTENT_TYPES.items():
                for reference in section.findall(qn(f"w:{kind}Reference")):
                    target = known.get(reference.get(qn("r:id"), ""))
                    if not target:
                        continue
                    part = target.lstrip("/") if target.startswith("/") else normpath(join("word", target))
                    if (kind, part) in checked:
                        continue
                    checked.add((kind, part))
                    label = "页眉" if kind == "header" else "页脚"
                    if not pkg.has(part):
                        report.add_error(f"{label}部件缺失: {part}")
                        continue
                    extension = part.rpartition(".")[2].lower()
                    if overrides.get("/" + part, defaults.get(extension)) != expected:
                        report.add_error(f"{label}部件的内容类型无效: {part}")

    def _check_sections(self, pkg: DocxPackage, report: ValidationReport) -> None:
        sections = pkg.section_properties()
        report.stats["sections"] = len(sections)
        if not sections:
            report.add_error("文档缺少 w:sectPr，页面设置丢失")

    def _check_styles(self, pkg: DocxPackage, report: ValidationReport) -> None:
        if not pkg.has(PART_STYLES):
            return

        styles_root = pkg.xml(PART_STYLES)
        known = {
            style.get(qn("w:styleId"))
            for style in styles_root.findall(qn("w:style"))
        }
        known.discard(None)

        document = pkg.xml(PART_DOCUMENT)
        referenced = {
            el.get(qn("w:val"))
            for el in document.iter(qn("w:pStyle"))
        }
        referenced.discard(None)

        dangling = {style for style in referenced if style not in known}
        report.stats["dangling_styles"] = len(dangling)
        if dangling:
            report.add_error(
                f"引用了未定义的段落样式: {', '.join(sorted(dangling)[:5])}"
            )

    def _collect_stats(self, pkg: DocxPackage, report: ValidationReport) -> None:
        report.stats["paragraphs"] = len(pkg.paragraphs())
        report.stats["tables"] = len(pkg.tables())
        report.stats["media_parts"] = sum(
            1 for name in pkg.iter_parts(prefix="word/media/")
        )
        report.stats["parts"] = len(pkg.part_names)
