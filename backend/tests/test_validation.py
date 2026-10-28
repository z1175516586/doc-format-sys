"""源文档规范校验的单元测试。"""

from __future__ import annotations

from pathlib import Path

from docx import Document

from app.domain.ooxml.conformance import DocumentValidator
from app.domain.ooxml.headings import HeadingProcessor
from app.domain.ooxml.package import DocxPackage


def _cover_doc(path: Path, stage: str) -> Path:
    doc = Document()
    doc.add_paragraph("某某工程项目")
    doc.add_paragraph(stage)
    doc.add_paragraph("设计人员名单")
    doc.save(path)
    return path


def _heading_doc(path: Path, items: list[tuple[str, int]]) -> Path:
    doc = Document()
    doc.add_paragraph("目      录")
    for text, level in items:
        doc.add_heading(text, level=level)
    doc.save(path)
    return path


def _heading_issues(path: Path):
    with DocxPackage.open(path) as pkg:
        report = HeadingProcessor(pkg).apply_heading_styles()
        return DocumentValidator(pkg).collect_heading_issues(report)


# ----------------------------------------------------------------------
# 封面文档类型
# ----------------------------------------------------------------------
def test_cover_type_accepts_canonical_with_suffix(tmp_path: Path) -> None:
    doc = _cover_doc(tmp_path / "cover.docx", "初步设计说明书")
    with DocxPackage.open(doc) as pkg:
        assert DocumentValidator(pkg).check_document_type("1") == []


def test_cover_type_rejects_nonstandard_wording(tmp_path: Path) -> None:
    doc = _cover_doc(tmp_path / "cover.docx", "可研设计")
    with DocxPackage.open(doc) as pkg:
        issues = DocumentValidator(pkg).check_document_type("2")
    codes = {item.code for item in issues}
    assert "document_type_unknown" in codes
    assert all(item.severity == "error" for item in issues)


def test_cover_type_rejects_mismatch(tmp_path: Path) -> None:
    doc = _cover_doc(tmp_path / "cover.docx", "初步设计")
    with DocxPackage.open(doc) as pkg:
        issues = DocumentValidator(pkg).check_document_type("2")
    assert [item.code for item in issues] == ["document_type_mismatch"]


def test_cover_type_combines_split_secondary_line(tmp_path: Path) -> None:
    """「初步设计」与「（代可研）」分两行写，也要能识别成第 3 类。"""
    doc = Document()
    doc.add_paragraph("某某工程项目")
    doc.add_paragraph("初步设计")
    doc.add_paragraph("（代可研）")
    doc.add_paragraph("设计人员名单")
    path = tmp_path / "cover.docx"
    doc.save(path)

    with DocxPackage.open(path) as pkg:
        assert DocumentValidator(pkg).check_document_type("3") == []
    with DocxPackage.open(path) as pkg:
        issues = DocumentValidator(pkg).check_document_type("1")
    assert [item.code for item in issues] == ["document_type_mismatch"]


# ----------------------------------------------------------------------
# 标题编号
# ----------------------------------------------------------------------
def test_heading_glued_number_is_warning(tmp_path: Path) -> None:
    doc = _heading_doc(tmp_path / "h.docx", [("1 概述", 1), ("1.1建设规模", 2)])
    issues = _heading_issues(doc)
    glued = [item for item in issues if item.code == "heading_number_glued"]
    assert glued and glued[0].severity == "warning"
    assert "1.1" in glued[0].message


def test_heading_jump_is_error(tmp_path: Path) -> None:
    doc = _heading_doc(tmp_path / "h.docx", [("1 概述", 1), ("1.1 规模", 2), ("1.3 内容", 2)])
    issues = _heading_issues(doc)
    errors = [item for item in issues if item.severity == "error"]
    assert [item.code for item in errors] == ["heading_jump_same_chapter"]
    assert errors[0].location


def test_heading_malformed_number_is_error(tmp_path: Path) -> None:
    doc = _heading_doc(tmp_path / "h.docx", [("1 概述", 1), ("1..1 规模", 2)])
    issues = _heading_issues(doc)
    errors = [item for item in issues if item.severity == "error"]
    assert [item.code for item in errors] == ["heading_number_invalid"]


def test_cross_chapter_jump_is_warning(tmp_path: Path) -> None:
    doc = _heading_doc(tmp_path / "h.docx", [("1 概述", 1), ("3 系统方案", 1)])
    issues = _heading_issues(doc)
    assert [item.code for item in issues] == ["heading_jump_cross_chapter"]
    assert issues[0].severity == "warning"


def test_body_list_item_not_reported(tmp_path: Path) -> None:
    doc = Document()
    doc.add_paragraph("目      录")
    doc.add_heading("1 概述", level=1)
    doc.add_paragraph("1）受场地限制，采用烟道反应脱硫。")
    doc.add_paragraph("2）脱硫系统入口烟气参数")
    path = tmp_path / "h.docx"
    doc.save(path)
    # 正文列表项 "1）…" 不是标题，不能被误报
    assert _heading_issues(path) == []
