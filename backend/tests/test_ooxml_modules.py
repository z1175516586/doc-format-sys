"""OOXML 各能力模块的单元测试。"""

from __future__ import annotations

from pathlib import Path

from app.core.namespaces import NSMAP, qn
from app.domain.ooxml.header_footer import HeaderFooterProcessor
from app.domain.ooxml.highlighting import HighlightProcessor
from app.domain.ooxml.package import DocxPackage
from app.domain.ooxml.pictures import PictureProcessor
from app.domain.ooxml.repair import DocumentRepairer
from app.domain.ooxml.sections import SectionProcessor
from app.domain.ooxml.tables import TableFont, TableProcessor
from app.domain.ooxml.toc import TocProcessor


def test_repair_adds_missing_heading_styles(sample_docx: Path, tmp_path: Path) -> None:
    with DocxPackage.open(sample_docx) as pkg:
        # 先删掉一个标题样式，模拟源文档缺失
        styles = pkg.xml("word/styles.xml")
        for style in list(styles.findall(qn("w:style"))):
            name = style.find(qn("w:name"))
            if name is not None and (name.get(qn("w:val")) or "").lower() == "heading 4":
                styles.remove(style)
                break

        added = DocumentRepairer(pkg).ensure_heading_styles()
        assert 4 in added

        output = tmp_path / "repaired.docx"
        pkg.save(output)

    with DocxPackage.open(output) as pkg:
        levels = DocumentRepairer(pkg)._existing_heading_levels(pkg.xml("word/styles.xml"))
        assert 4 in levels


def test_section_orientation_normalized(sample_docx: Path) -> None:
    with DocxPackage.open(sample_docx) as pkg:
        processor = SectionProcessor(pkg)
        processor.normalize_orientation()

        for info in processor.describe():
            if info.orientation == "landscape":
                assert info.width > info.height
            elif info.orientation == "portrait":
                assert info.height >= info.width


def test_clear_footers(sample_docx: Path) -> None:
    with DocxPackage.open(sample_docx) as pkg:
        cleared = HeaderFooterProcessor(pkg).clear_all_footers()
        assert cleared >= 0

        for name in pkg.iter_parts(prefix="word/footer", suffix=".xml"):
            texts = "".join(node.text or "" for node in pkg.xml(name).iter(qn("w:t")))
            assert "旧页脚内容" not in texts


def test_highlight_processor_counts(sample_docx: Path) -> None:
    with DocxPackage.open(sample_docx) as pkg:
        stats = HighlightProcessor(pkg).remove_all()
        assert stats["highlight"] >= 1
        assert pkg.xml("word/document.xml").xpath(".//w:highlight", namespaces=NSMAP) == []


def test_table_font_applied(sample_docx: Path) -> None:
    with DocxPackage.open(sample_docx) as pkg:
        processor = TableProcessor(pkg)
        count = processor.apply_font_to_first_table(TableFont.from_config("table_cover"))
        assert count > 0

        first_table = pkg.tables()[0]
        fonts = first_table.findall(f".//{qn('w:rFonts')}")
        assert fonts
        assert any(font.get(qn("w:eastAsia")) == "宋体" for font in fonts)


def test_cover_table_removal_requires_marker(tmp_path: Path) -> None:
    """缺少标记时不得删除任何表格（防止误删正文表格）。"""
    from docx import Document

    doc = Document()
    doc.add_paragraph("没有标记的文档")
    table = doc.add_table(rows=1, cols=1)
    table.cell(0, 0).text = "正文表格"
    path = tmp_path / "no_marker.docx"
    doc.save(path)

    with DocxPackage.open(path) as pkg:
        removed = TableProcessor(pkg).remove_tables_before_marker()
        assert removed == 0
        assert len(pkg.tables()) == 1


def test_toc_insert_and_detect(sample_docx: Path) -> None:
    with DocxPackage.open(sample_docx) as pkg:
        processor = TocProcessor(pkg)
        assert not processor.has_toc_field()
        assert processor.insert_field()
        assert processor.has_toc_field()


def test_toc_title_replace(sample_docx: Path) -> None:
    with DocxPackage.open(sample_docx) as pkg:
        processor = TocProcessor(pkg)
        assert processor.replace_title("目  录")

        texts = [node.text or "" for node in pkg.body.iter(qn("w:t"))]
        assert "目  录" in texts


def test_picture_processor_no_crash_without_images(sample_docx: Path) -> None:
    with DocxPackage.open(sample_docx) as pkg:
        processor = PictureProcessor(pkg)
        assert processor.count() == 0
        assert processor.reset_indent() == 0
