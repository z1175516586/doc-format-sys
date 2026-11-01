"""文档转换树：身份戳、识别、失败叶子、内容被改。"""

from __future__ import annotations

import zipfile
from pathlib import Path
from shutil import copyfile

from lxml import etree

from app.core.namespaces import qn, w_attr
from app.domain.ooxml.lineage import LineageStamp
from app.domain.ooxml.package import DocxPackage
from app.schemas.document import FormatResult
from app.services.history_service import HistoryService


def _service(tmp_path: Path) -> HistoryService:
    return HistoryService(root=tmp_path / "history")


def _touch_text(path: Path, extra: str) -> None:
    with DocxPackage.open(path) as pkg:
        body = pkg.body
        from lxml import etree

        from app.core.namespaces import qn

        paragraph = etree.SubElement(body, qn("w:p"))
        run = etree.SubElement(paragraph, qn("w:r"))
        text = etree.SubElement(run, qn("w:t"))
        text.text = extra
        pkg.save(path)


def test_stamp_roundtrip(sample_docx: Path, tmp_path: Path) -> None:
    target = tmp_path / "stamped.docx"
    copyfile(sample_docx, target)
    LineageStamp().write(target, "abc123")
    assert LineageStamp().read(target) == "abc123"


def test_stamp_lives_in_settings_docvars(sample_docx: Path, tmp_path: Path) -> None:
    """身份必须写在 settings.xml 的 w:docVars，Word/WPS 编辑保存后才不会丢。

    旧方案写在 customXml 部件，Word 另存时会整块删掉，导致 node_id 消失。
    """
    target = tmp_path / "stamped.docx"
    copyfile(sample_docx, target)
    LineageStamp().write(target, "abc123")

    with zipfile.ZipFile(target) as zf:
        settings = etree.fromstring(zf.read("word/settings.xml"))
        names = zf.namelist()

    doc_vars = settings.find(qn("w:docVars"))
    assert doc_vars is not None, "settings.xml 缺少 w:docVars"
    values = {
        var.get(w_attr("name")): var.get(w_attr("val"))
        for var in doc_vars.findall(qn("w:docVar"))
    }
    assert values["DocHistoryId"] == "abc123"

    assert not any("dfsLineage" in name for name in names), "不应再产生 customXml 身份部件"


def test_docvars_follow_settings_schema_order(sample_docx: Path, tmp_path: Path) -> None:
    """w:docVars 必须排在 rsids / mathPr 等元素之前，否则 Word 判定结构非法。"""
    target = tmp_path / "stamped.docx"
    copyfile(sample_docx, target)
    LineageStamp().write(target, "abc123")

    with zipfile.ZipFile(target) as zf:
        settings = etree.fromstring(zf.read("word/settings.xml"))

    tags: list[str] = []
    for child in settings:
        try:
            tags.append(etree.QName(child).localname)
        except ValueError:
            continue
    position = tags.index("docVars")
    for later in ("rsids", "mathPr", "themeFontLang", "decimalSymbol", "listSeparator"):
        if later in tags:
            assert position < tags.index(later), f"docVars 应排在 {later} 之前"


def test_legacy_custom_xml_still_recognized(sample_docx: Path, tmp_path: Path) -> None:
    """旧文档的身份写在 customXml 里，读取仍要识别出来。"""
    target = tmp_path / "legacy.docx"
    copyfile(sample_docx, target)
    stamp = LineageStamp()
    with DocxPackage.open(target) as pkg:
        pkg.remove_part(stamp.settings_part)
        pkg.add_part(
            stamp.part_name,
            f'<?xml version="1.0" encoding="UTF-8"?><lineage xmlns="{stamp.namespace}">'
            f"<nodeId>legacy-999</nodeId></lineage>".encode("utf-8"),
        )
        pkg.save(target)
    assert stamp.read(target) == "legacy-999"


def test_unknown_file_has_empty_tree(sample_docx: Path, tmp_path: Path) -> None:
    view = _service(tmp_path).inspect(sample_docx, sample_docx.name)
    assert view["new_document"] is True
    assert view["tree"] is None


def test_same_source_converted_twice_is_two_children(sample_docx: Path, tmp_path: Path) -> None:
    history = _service(tmp_path)
    source = history.ensure_source(sample_docx, "A.docx")
    parent_id = source["node"]["node_id"]

    first = tmp_path / "A1.docx"
    second = tmp_path / "A2.docx"
    copyfile(sample_docx, first)
    copyfile(sample_docx, second)
    result = FormatResult(success=True, output_path=str(first), message="ok")
    history.register_success(
        parent_id=parent_id,
        output_path=first,
        output_name="A1.docx",
        task_id="t1",
        result=result,
        source_name="A.docx",
        source_hash=history.file_hash(sample_docx),
    )
    result.output_path = str(second)
    history.register_success(
        parent_id=parent_id,
        output_path=second,
        output_name="A2.docx",
        task_id="t2",
        result=result,
        source_name="A.docx",
        source_hash=history.file_hash(sample_docx),
    )

    view = history.inspect(sample_docx, "A.docx")
    children = view["tree"]["children"]
    documents = [item for item in children if item["kind"] == "document"]
    assert len(documents) == 2
    assert documents[0]["node_id"] != documents[1]["node_id"]
    assert LineageStamp().read(first) == documents[0]["node_id"] or LineageStamp().read(first) == documents[1]["node_id"]


def test_failure_is_leaf_without_descendants(sample_docx: Path, tmp_path: Path) -> None:
    history = _service(tmp_path)
    parent_id = history.ensure_source(sample_docx, "A.docx")["node"]["node_id"]
    history.register_failure(
        parent_id=parent_id,
        task_id="fail-1",
        source_name="A.docx",
        source_hash=history.file_hash(sample_docx),
        message="封面文档类型不规范",
    )
    view = history.inspect(sample_docx, "A.docx")
    failed = [item for item in view["tree"]["children"] if item["kind"] == "failed"]
    assert len(failed) == 1
    assert failed[0]["children"] == []
    assert "封面文档类型不规范" in failed[0]["message"]
    assert failed[0]["download_url"] is None


def test_node_id_same_hash_different_is_modified(sample_docx: Path, tmp_path: Path) -> None:
    history = _service(tmp_path)
    node_id = history.ensure_source(sample_docx, "A.docx")["node"]["node_id"]
    stored = tmp_path / "history" / "files" / f"{node_id}.docx"
    edited = tmp_path / "edited.docx"
    copyfile(stored, edited)
    _touch_text(edited, "用户改过的内容")
    assert LineageStamp().read(edited) == node_id
    view = history.inspect(edited, "edited.docx")
    assert view["new_document"] is False
    assert view["modified"] is True
    assert view["current_node_id"] == node_id
    assert view["modified_message"]


def test_missing_file_is_flagged(sample_docx: Path, tmp_path: Path) -> None:
    history = _service(tmp_path)
    node_id = history.ensure_source(sample_docx, "A.docx")["node"]["node_id"]
    stored = Path(history.get_node(node_id)["file_path"])
    stored.unlink()
    view = history.inspect(sample_docx, "A.docx")
    assert view["tree"]["file_available"] is False
    assert view["tree"]["missing_file_message"] == "有记录、文件已删"


def test_child_upload_shows_path_to_root(sample_docx: Path, tmp_path: Path) -> None:
    history = _service(tmp_path)
    parent_id = history.ensure_source(sample_docx, "A.docx")["node"]["node_id"]
    child = tmp_path / "B.docx"
    copyfile(sample_docx, child)
    history.register_success(
        parent_id=parent_id,
        output_path=child,
        output_name="B.docx",
        task_id="t",
        result=FormatResult(success=True, output_path=str(child), message="ok"),
        source_name="A.docx",
        source_hash=history.file_hash(sample_docx),
    )
    view = history.inspect(child, "B.docx")
    names = [item["name"] for item in view["path"]]
    assert names == ["A.docx", "B.docx"]
    assert view["path"][-1]["current"] is True
    assert view["tree"]["children"][0]["current"] is True
