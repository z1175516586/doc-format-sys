"""排版流水线端到端测试。

这些测试的作用是"回归基线"的最小可用版本：
只要流水线被破坏，这里必然失败。后续应扩展为 golden-file 比对。
"""

from __future__ import annotations

from pathlib import Path

import pytest

from app.converters.noop import NoopConverter
from app.core.namespaces import NSMAP, qn
from app.domain.ooxml.package import DocxPackage, paragraph_text
from app.domain.ooxml.toc import TocProcessor
from app.schemas.document import ConvertOptions, FormatRequest
from app.services.format_service import PIPELINE_STEPS, DocumentFormatService


def _run(sample: Path, output: Path, **option_overrides) -> tuple[bool, Path, object]:
    # 单测用合成文档验证流水线机制，不做封面模板替换（模板替换由真实文档回归验证）
    option_overrides.setdefault("use_cover_template", False)
    options = ConvertOptions(**option_overrides)
    request = FormatRequest(
        source_path=str(sample),
        output_path=str(output),
        template_path=None,
        options=options,
    )
    service = DocumentFormatService(converter=NoopConverter())
    result = service.format(request)
    return result.success, output, result


def test_pipeline_succeeds(sample_docx: Path, tmp_path: Path) -> None:
    success, output, result = _run(sample_docx, tmp_path / "out.docx")

    assert success, f"流水线失败: {result.message}"
    assert output.exists()
    assert output.stat().st_size > 0
    assert len(result.steps) == len(PIPELINE_STEPS)
    assert all(step.ok for step in result.steps), [
        (step.name, step.message) for step in result.steps if not step.ok
    ]


def test_toc_field_inserted(sample_docx: Path, tmp_path: Path) -> None:
    success, output, _ = _run(sample_docx, tmp_path / "out.docx", insert_toc=True)
    assert success

    with DocxPackage.open(output) as pkg:
        assert TocProcessor(pkg).has_toc_field()


def test_toc_skipped_when_disabled(sample_docx: Path, tmp_path: Path) -> None:
    success, output, _ = _run(sample_docx, tmp_path / "out.docx", insert_toc=False)
    assert success

    with DocxPackage.open(output) as pkg:
        assert not TocProcessor(pkg).has_toc_field()


def test_footers_cleared(sample_docx: Path, tmp_path: Path) -> None:
    success, output, _ = _run(sample_docx, tmp_path / "out.docx")
    assert success

    with DocxPackage.open(output) as pkg:
        for name in pkg.iter_parts(prefix="word/footer", suffix=".xml"):
            text = "".join(node.text or "" for node in pkg.xml(name).iter(qn("w:t")))
            assert "旧页脚内容" not in text


def test_highlights_removed(sample_docx: Path, tmp_path: Path) -> None:
    success, output, _ = _run(sample_docx, tmp_path / "out.docx", remove_highlights=True)
    assert success

    with DocxPackage.open(output) as pkg:
        root = pkg.xml("word/document.xml")
        assert root.xpath(".//w:highlight", namespaces=NSMAP) == []


def test_cover_table_removed_content_table_kept(sample_docx: Path, tmp_path: Path) -> None:
    """标记「设计人员名单」之前的封面表格应被移除，正文表格应保留。"""
    success, output, _ = _run(sample_docx, tmp_path / "out.docx")
    assert success

    with DocxPackage.open(output) as pkg:
        tables = pkg.tables()
        assert len(tables) == 1, f"应只保留正文表格，实际 {len(tables)} 个"

        texts = [
            paragraph_text(p)
            for p in pkg.paragraphs()
        ]
        # 封面表格内容转为段落保留，不丢信息
        assert any("封面表格" in text for text in texts)


def test_section_break_created(sample_docx: Path, tmp_path: Path) -> None:
    """封面、目录、正文应形成多节结构。"""
    success, output, _ = _run(sample_docx, tmp_path / "out.docx")
    assert success

    with DocxPackage.open(output) as pkg:
        assert len(pkg.section_properties()) >= 2


def test_output_reopenable_by_python_docx(sample_docx: Path, tmp_path: Path) -> None:
    """输出必须能被 python-docx 正常打开，否则 Word 很可能报错。"""
    success, output, _ = _run(sample_docx, tmp_path / "out.docx")
    assert success

    from docx import Document

    doc = Document(str(output))
    assert len(doc.paragraphs) > 0


def test_no_warning_when_page_numbers_left_to_word(sample_docx: Path, tmp_path: Path) -> None:
    """默认配置关掉服务端刷新，页码交给 Word 打开时算，不算风险，不该告警。"""
    success, output, result = _run(sample_docx, tmp_path / "out.docx")
    assert success
    assert not any("目录页码" in warning for warning in result.warnings)

    with DocxPackage.open(output) as pkg:
        settings = pkg.xml("word/settings.xml")
        assert settings.find(qn("w:updateFields")) is not None


def test_warning_when_no_page_number_source(sample_docx: Path, tmp_path: Path) -> None:
    """服务端不刷新、又不让 Word 更新域时，必须显式告警。"""
    success, _, result = _run(
        sample_docx,
        tmp_path / "out.docx",
        update_fields_on_open=False,
        server_refresh_fields=False,
    )
    assert success
    assert any("目录页码" in warning for warning in result.warnings)


def test_page_number_source_options_override_config(sample_docx: Path, tmp_path: Path) -> None:
    """前端选择优先于配置文件默认值。"""
    success, output, result = _run(
        sample_docx,
        tmp_path / "out.docx",
        update_fields_on_open=False,
        server_refresh_fields=False,
    )
    assert success

    with DocxPackage.open(output) as pkg:
        settings = pkg.xml("word/settings.xml")
        assert settings.find(qn("w:updateFields")) is None

    engine = next(step for step in result.steps if step.name == "engine")
    assert engine.stats["skipped"] is True
    assert engine.stats["update_fields"] is False


def test_both_page_sources_rejected() -> None:
    """两个来源同时开启必须被拒绝，不能默默白占全局引擎。"""
    from pydantic import ValidationError as PydanticValidationError

    with pytest.raises(PydanticValidationError):
        ConvertOptions(update_fields_on_open=True, server_refresh_fields=True)


def test_single_page_source_choice_clears_the_other() -> None:
    """只传一个开关时，另一个必须收敛为 False，不能回落到配置默认造出矛盾组合。"""
    from app.api.v1.documents import _resolve_page_source

    # 只选服务端：客户端更新必须关掉（配置默认为 True，不能漏过来）
    assert _resolve_page_source(None, True) == (False, True)
    # 只选客户端
    assert _resolve_page_source(True, None) == (True, False)
    # 都不传：取配置默认
    assert _resolve_page_source(None, None) == (True, False)
    # 都传且不矛盾
    assert _resolve_page_source(False, True) == (False, True)


def test_engine_waiting_event_reaches_progress(sample_docx: Path, tmp_path: Path) -> None:
    """引擎排队时必须把 waiting 推给前端，否则用户以为程序卡死。"""

    class _QueueingConverter(NoopConverter):
        def capabilities(self):
            from app.converters.base import ConverterCapabilities

            return ConverterCapabilities(
                name="fake", available=True, refresh_fields=True
            )

        def refresh_fields(self, path, on_phase=None) -> bool:
            if on_phase is not None:
                on_phase("waiting", "前面有任务，已排队")
                on_phase("running", "已取得引擎")
            return False

    options = ConvertOptions(
        use_cover_template=False,
        server_refresh_fields=True,
        update_fields_on_open=False,
    )
    request = FormatRequest(
        source_path=str(sample_docx),
        output_path=str(tmp_path / "out.docx"),
        options=options,
    )
    events: list[dict[str, object]] = []

    def on_progress(step, index, total, message, data=None):
        if data and data.get("engine_phase"):
            events.append(dict(data))

    DocumentFormatService(converter=_QueueingConverter()).format(
        request, on_progress=on_progress
    )

    phases = [event["engine_phase"] for event in events]
    assert phases == ["waiting", "running"]
    assert events[0]["waiting"] is True
    assert events[1]["waiting"] is False


def test_refresh_gate_serializes_and_reports_phase() -> None:
    """闸门必须串行，且排队方要收到 waiting、拿到后收到 running。"""
    import threading
    import time

    from app.converters.libreoffice import (
        ENGINE_PHASE_RUNNING,
        ENGINE_PHASE_WAITING,
        _RefreshGate,
    )

    gate = _RefreshGate()
    phases: list[str] = []

    gate.acquire()
    worker = threading.Thread(
        target=lambda: (gate.acquire(on_phase=lambda phase, _: phases.append(phase)), gate.release()),
    )
    try:
        worker.start()
        time.sleep(0.2)  # 让 worker 先卡在闸门上
    finally:
        gate.release()
    worker.join(timeout=10)

    assert not worker.is_alive(), "worker 未能在主线程释放后通过闸门"
    assert phases == [ENGINE_PHASE_WAITING, ENGINE_PHASE_RUNNING]


def test_missing_source_fails_cleanly(tmp_path: Path) -> None:
    options = ConvertOptions()
    request = FormatRequest(
        source_path=str(tmp_path / "不存在.docx"),
        output_path=str(tmp_path / "out.docx"),
        options=options,
    )
    result = DocumentFormatService(converter=NoopConverter()).format(request)

    assert not result.success
    assert result.steps[0].name == "open"
    assert not result.steps[0].ok
