"""DocxPackage 基础能力测试。"""

from __future__ import annotations

from pathlib import Path

from app.domain.ooxml.package import DocxPackage, paragraph_text


def test_open_and_read(sample_docx: Path) -> None:
    with DocxPackage.open(sample_docx) as pkg:
        assert pkg.has("word/document.xml")
        assert len(pkg.part_names) > 5
        assert len(pkg.paragraphs()) > 0


def test_roundtrip_preserves_content(sample_docx: Path, tmp_path: Path) -> None:
    output = tmp_path / "roundtrip.docx"

    with DocxPackage.open(sample_docx) as pkg:
        before = [paragraph_text(p) for p in pkg.paragraphs()]
        pkg.save(output)

    assert output.exists()
    with DocxPackage.open(output) as pkg:
        after = [paragraph_text(p) for p in pkg.paragraphs()]

    assert before == after


def test_clone_is_independent(sample_docx: Path) -> None:
    with DocxPackage.open(sample_docx) as pkg:
        clone = pkg.clone()
        original_count = len(pkg.paragraphs())

        body = clone.body
        for child in list(body):
            if child.tag.endswith("}p"):
                body.remove(child)

        assert len(clone.paragraphs()) == 0
        assert len(pkg.paragraphs()) == original_count


def test_invalid_file_raises(tmp_path: Path) -> None:
    import pytest

    from app.core.errors import DocumentCorruptedError

    broken = tmp_path / "broken.docx"
    broken.write_text("not a zip", encoding="utf-8")

    with pytest.raises(DocumentCorruptedError):
        DocxPackage.open(broken)
