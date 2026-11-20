from pathlib import Path

from docx import Document
from docx.shared import Pt

from app.domain.ooxml.document_comparison import compare_documents


def test_comparison_aligns_text_and_reports_inherited_and_direct_formatting(tmp_path: Path) -> None:
    before = tmp_path / "before.docx"
    after = tmp_path / "after.docx"
    for path, changed in ((before, False), (after, True)):
        doc = Document()
        doc.styles["Normal"].font.size = Pt(14 if changed else 12)
        doc.add_heading("项目概况", level=1)
        paragraph = doc.add_paragraph()
        paragraph.add_run("设计").bold = not changed
        paragraph.add_run("依据").bold = changed
        doc.add_paragraph("建设规模为十万吨" if changed else "建设规模为五万吨")
        doc.add_paragraph("新增说明" if changed else "旧版备注")
        table = doc.add_table(rows=1, cols=2)
        table.cell(0, 0).text = "能力"
        table.cell(0, 1).text = "十万吨" if changed else "五万吨"
        doc.save(path)

    result = compare_documents(before, after, "原稿.docx", "成品.docx")
    assert result.counts == {"equal": 1, "added": 1, "removed": 1, "modified": 2, "style": 1}
    style_only = next(row for row in result.rows if row.kind == "style")
    styles = {change.label: change for change in style_only.styles}
    assert (styles["字号"].before, styles["字号"].after) == ("12 磅", "14 磅")
    assert "加粗" in styles
    for row in (row for row in result.rows if row.kind == "modified"):
        assert [part.text for part in row.before.text if part.changed] == ["五"]
        assert [part.text for part in row.after.text if part.changed] == ["十"]
    assert result.rows[-1].block_type == "table_row"


def test_same_text_with_different_run_splits_stays_equal(tmp_path: Path) -> None:
    before, after = tmp_path / "before.docx", tmp_path / "after.docx"
    source = Document()
    source.add_paragraph("建设规模")
    source.save(before)
    target = Document()
    paragraph = target.add_paragraph("建设")
    paragraph.add_run("规模")
    target.save(after)
    report = compare_documents(before, after, before.name, after.name)
    assert report.counts["equal"] == 1
