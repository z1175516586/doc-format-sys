"""源文档规范校验的单元测试。"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.text import WD_BREAK
import pytest

from app.domain.ooxml.conformance import DocumentValidator
from app.domain.ooxml.cover import CoverMetadataExtractor
from app.domain.ooxml.headings import HeadingProcessor
from app.domain.ooxml.package import DocxPackage, paragraph_text

_ISSUER = "鞍钢集团工程技术有限公司"


def _cover_doc(
    path: Path,
    page_one: list[str],
    page_two: list[str] | None = None,
    *,
    page_break: bool = True,
) -> Path:
    """封面第一页、第二页的段落，最后接设计人员名单。"""
    doc = Document()
    for text in page_one:
        doc.add_paragraph(text)
    if page_two is not None:
        if page_break:
            doc.paragraphs[-1].add_run().add_break(WD_BREAK.PAGE)
        for text in page_two:
            doc.add_paragraph(text)
    doc.add_paragraph("各专业参加设计人员名单")
    doc.save(path)
    return path


def _type_issues(path: Path, selected: str):
    with DocxPackage.open(path) as pkg:
        return DocumentValidator(pkg).check_document_type(selected)


def _heading_doc(path: Path, items: list[tuple[str, int]]) -> Path:
    doc = Document()
    doc.add_paragraph("目      录")
    for text, level in items:
        doc.add_heading(text, level=level)
    doc.save(path)
    return path


def _plain_doc(path: Path, texts: list[str]) -> Path:
    """只用普通段落，模拟"作者没套标题样式、完全靠编号识别"的文档。"""
    doc = Document()
    doc.add_paragraph("目      录")
    for text in texts:
        doc.add_paragraph(text)
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
    doc = _cover_doc(tmp_path / "cover.docx", ["某某工程项目", "初步设计说明书", _ISSUER])
    assert _type_issues(doc, "1") == []


def test_cover_type_accepts_spaced_wording(tmp_path: Path) -> None:
    doc = _cover_doc(
        tmp_path / "cover.docx",
        ["某某工程项目", "可 行 性 研 究", "鞍 钢 集 团 工 程 技 术 有 限 公 司"],
    )
    assert _type_issues(doc, "2") == []


def test_cover_issuer_missing_is_error(tmp_path: Path) -> None:
    doc = _cover_doc(tmp_path / "cover.docx", ["某某工程项目", "初步设计", "某某设计院"])
    issues = _type_issues(doc, "1")
    assert [item.code for item in issues] == ["cover_issuer_missing"]
    assert issues[0].severity == "error"


def test_cover_type_above_issuer_must_be_standard(tmp_path: Path) -> None:
    doc = _cover_doc(tmp_path / "cover.docx", ["某某工程项目", "可研设计", _ISSUER])
    issues = _type_issues(doc, "2")
    assert [item.code for item in issues] == ["front_type_invalid"]
    assert "可研设计" in issues[0].message


def test_mislabelled_front_page_without_page_break(tmp_path: Path) -> None:
    """二烧翻车机：第一页写成「可研设计」，第二页才是「可行性研究」，两页之间没有分页符。"""
    doc = _cover_doc(
        tmp_path / "cover.docx",
        ["鞍钢股份烧结厂二烧2号翻车机新增除尘器项目", "可研设计", _ISSUER, "2024年8月"],
        ["鞍钢股份烧结厂二烧2号翻车机新增除尘器项目", "可行性研究", "总设计师：吕艳平", _ISSUER],
        page_break=False,
    )
    issues = _type_issues(doc, "2")
    assert [item.code for item in issues] == ["front_type_invalid"]
    assert issues[0].severity == "error"


def test_cover_project_missing_is_error(tmp_path: Path) -> None:
    doc = _cover_doc(tmp_path / "cover.docx", ["初步设计", _ISSUER])
    issues = _type_issues(doc, "1")
    assert [item.code for item in issues] == ["front_project_missing"]


def test_cover_type_rejects_mismatch(tmp_path: Path) -> None:
    doc = _cover_doc(tmp_path / "cover.docx", ["某某工程项目", "初步设计", _ISSUER])
    issues = _type_issues(doc, "2")
    assert [item.code for item in issues] == ["document_type_mismatch"]


def test_cover_type_combines_split_secondary_line(tmp_path: Path) -> None:
    """「初步设计」与「（代可研）」分两行写，也要能识别成第 3 类。"""
    doc = _cover_doc(tmp_path / "cover.docx", ["某某工程项目", "初步设计", "（代可研）", _ISSUER])
    assert _type_issues(doc, "3") == []
    assert [item.code for item in _type_issues(doc, "1")] == ["document_type_mismatch"]


def test_page_two_type_mismatch_is_warning(tmp_path: Path) -> None:
    doc = _cover_doc(
        tmp_path / "cover.docx",
        ["某某工程项目", "初步设计", _ISSUER, "2024年8月"],
        ["某某工程项目", "可行性研究", _ISSUER],
    )
    issues = _type_issues(doc, "1")
    assert [item.code for item in issues] == ["page_two_type_mismatch"]
    assert issues[0].severity == "warning"


def test_project_name_comes_from_front_page_only(tmp_path: Path) -> None:
    """项目名只取第一页类型上方的文字：单位名单独一行，坐标残留去掉，不拼到第二页。"""
    doc = _cover_doc(
        tmp_path / "cover.docx",
        ["6032578105", "本溪玉麒麟新材料有限公司", "竖炉节能增效改造EPC总承包工程", "初步设计", _ISSUER],
        ["本溪玉麒麟新材料有限公司", "竖炉节能增效改造EPC总承包工程", "初步设计", _ISSUER],
    )
    with DocxPackage.open(doc) as pkg:
        meta = CoverMetadataExtractor(pkg).extract()
    assert meta.project_name == "本溪玉麒麟新材料有限公司\n竖炉节能增效改造EPC总承包工程"


# ----------------------------------------------------------------------
# 标题编号
# ----------------------------------------------------------------------
def test_heading_glued_number_is_warning(tmp_path: Path) -> None:
    doc = _heading_doc(tmp_path / "h.docx", [("1 概述", 1), ("1.1建设规模", 2)])
    issues = _heading_issues(doc)
    glued = [item for item in issues if item.code == "heading_number_glued"]
    assert glued and glued[0].severity == "warning"
    assert "1.1" in glued[0].message


def test_heading_extra_spaces_is_warning(tmp_path: Path) -> None:
    """同级别对齐被打乱时提示多空格，且不改内容。"""
    doc = _heading_doc(
        tmp_path / "h.docx",
        [("1 概述", 1), ("1.1 设计依据", 2), ("1.2  建设规模", 2)],
    )
    issues = _heading_issues(doc)
    spacing = [item for item in issues if item.code == "heading_number_spacing"]
    assert spacing and spacing[0].severity == "warning"
    assert "2 个空白" in spacing[0].message

    with DocxPackage.open(doc) as pkg:
        texts = [paragraph_text(paragraph) for paragraph in pkg.paragraphs()]
    assert "1.2  建设规模" in texts, "提示归提示，原文不得被改动"


def test_heading_extra_spaces_used_for_alignment_is_silent(tmp_path: Path) -> None:
    """单位数与两位数编号用不同空格数对齐，属有意排版，不该报。

    真实文档就是这么写的：``1  总论``（1+2 空格）与 ``11 机械化设备``（11+1 空格）
    的标题文字都起始于第 4 列。逐条报会刷出十几条噪声。
    """
    doc = _heading_doc(
        tmp_path / "h.docx",
        [
            ("1  总论", 1),
            ("2  建筑", 1),
            ("11 机械化设备", 1),
            ("12 环境保护", 1),
        ],
    )
    issues = _heading_issues(doc)
    assert [item.code for item in issues if "spacing" in item.code] == []


@pytest.mark.parametrize(
    ("gap", "rest", "is_value"),
    [
        # 编号紧贴字母：单位/量纲，必是正文
        ("", "kV电压", True),
        ("", "V低压配电", True),
        ("", "m标高设配电柜", True),
        ("", "SCADA系统", True),
        # 编号紧贴中文计量单位
        ("", "万元", True),
        ("", "倍", True),
        ("", "吨", True),
        # 编号紧贴汉字标题文字：真标题
        ("", "设计依据", False),
        ("", "项目厂址现状", False),
        ("", "年度计划", False),  # 「年」不在单位表内，不能被误杀
        ("", "元素分析", False),  # 「元」不在单位表内，不能被误杀
        # 编号后有空白：不是本函数管辖范围
        (" ", "电气", False),
        (" ", "设计依据", False),
    ],
)
def test_numbered_value_detection(gap: str, rest: str, is_value: bool) -> None:
    """编号紧贴内容时，靠紧跟的字符区分正文与标题。"""
    from app.domain.ooxml.headings import _looks_like_numbered_value

    assert _looks_like_numbered_value(gap, rest) is is_value


def test_numbered_prose_does_not_hijack_heading_chain(tmp_path: Path) -> None:
    """段首 "10kV电压…" 不是标题，更不能把编号链带偏。

    误判的代价不只是多一条目录项：``10kV电压`` 被读成一级标题 10 后，
    编号链被带偏，后面真正的 ``10 电气`` 会与之"同级重复"，
    再往后的标题还可能被判跳号而降级。
    """
    prose = "10kV电压引自厂区配电室，采用电缆桥架敷设至各车间。"
    chapters = [f"{index} 第{index}章" for index in range(1, 10)]
    doc = _plain_doc(
        tmp_path / "h.docx",
        chapters + ["9.1 节", "9.2 节", "9.3 节", prose, "10 电气"],
    )

    with DocxPackage.open(doc) as pkg:
        processor = HeadingProcessor(pkg)
        report = processor.apply_heading_styles()
        styled = [
            paragraph_text(paragraph)
            for paragraph in pkg.paragraphs()
            if processor.styles.resolver.is_heading(
                processor.styles.paragraph_style_id(paragraph) or ""
            )
        ]

    assert prose not in styled, "以数字开头的正文被认成了标题"
    assert "10 电气" in styled, "真标题被正文带偏后降级了"
    assert report.levels.get(1, 0) == 10, (
        f"一级标题应为 1..10 共 10 处，实际 {report.levels.get(1, 0)} 处"
    )


def test_many_extra_spaces_are_merged_into_one_issue(tmp_path: Path) -> None:
    """多处多余空白合并成一条，避免占满问题清单把错误挤出去。"""
    items: list[tuple[str, int]] = [("1.1 甲", 2)]
    items += [(f"1.{index}  第{index}节", 2) for index in range(2, 7)]
    doc = _heading_doc(tmp_path / "h.docx", items)
    issues = _heading_issues(doc)
    spacing = [item for item in issues if item.code == "heading_number_spacing"]
    assert len(spacing) == 1, f"5 处多余空白应合并成 1 条，实际 {len(spacing)} 条"
    assert "共 5 处" in spacing[0].message
    assert spacing[0].location == "全文"


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
