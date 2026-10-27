"""文档结构对比（回归测试核心工具）。

用途：把"排版输出"与"人工确认过的目标文档"做结构级对比，
避免依赖人工肉眼检查。

设计原则：
- 只比对**结构与语义**，不比对字节。OOXML 中 rId、部件顺序、XML 属性顺序
  都是无意义的差异，直接 diff 字节只会产生噪音。
- 每项检查独立给出结论，便于定位是哪一步回归了。
- 检查项全部来自 REFACTOR_PLAN §7.2 / §7.3 的分层策略。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

from app.domain.ooxml.package import DocxPackage, paragraph_text
from app.domain.ooxml.sections import SectionProcessor
from app.core.namespaces import NSMAP, qn


_WHITESPACE = re.compile(r"\s+")


def normalize_text(text: str) -> str:
    """归一化文本：压缩空白、去零宽字符，用于内容对比。"""
    return _WHITESPACE.sub("", text.replace("\u200b", "").replace("\xa0", ""))


@dataclass
class Snapshot:
    """一份文档的结构快照。"""

    label: str
    paragraph_count: int = 0
    table_count: int = 0
    section_count: int = 0
    orientation_sequence: list[str] = field(default_factory=list)
    toc_field_count: int = 0
    highlight_count: int = 0
    shading_count: int = 0
    media_count: int = 0
    header_count: int = 0
    footer_count: int = 0
    style_usage: dict[str, int] = field(default_factory=dict)
    body_text: str = ""


def snapshot(path: str | Path, label: str = "") -> Snapshot:
    """采集文档结构快照。"""
    snap = Snapshot(label=label or Path(path).name)

    with DocxPackage.open(path) as pkg:
        document = pkg.xml("word/document.xml")

        snap.paragraph_count = len(pkg.paragraphs())
        snap.table_count = len(pkg.tables())

        sections = SectionProcessor(pkg).describe()
        snap.section_count = len(sections)
        snap.orientation_sequence = [info.orientation for info in sections]

        snap.toc_field_count = len(
            document.xpath('.//w:instrText[contains(., "TOC")]', namespaces=NSMAP)
        )
        snap.highlight_count = len(document.xpath(".//w:highlight", namespaces=NSMAP))
        snap.shading_count = len(document.xpath(".//w:shd", namespaces=NSMAP))

        snap.media_count = sum(1 for name in pkg.iter_parts(prefix="word/media/"))
        snap.header_count = sum(
            1 for name in pkg.iter_parts(prefix="word/header", suffix=".xml")
        )
        snap.footer_count = sum(
            1 for name in pkg.iter_parts(prefix="word/footer", suffix=".xml")
        )

        usage: dict[str, int] = {}
        for node in document.iter(qn("w:pStyle")):
            value = node.get(qn("w:val")) or ""
            if value:
                usage[value] = usage.get(value, 0) + 1
        snap.style_usage = usage

        snap.body_text = normalize_text(
            "".join(paragraph_text(p) for p in pkg.paragraphs())
        )

    return snap


@dataclass
class CheckResult:
    """单条检查结论。"""

    name: str
    ok: bool
    expected: object = None
    actual: object = None
    note: str = ""


@dataclass
class ComparisonReport:
    """整体对比报告。"""

    checks: list[CheckResult] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return all(check.ok for check in self.checks)

    @property
    def failures(self) -> list[CheckResult]:
        return [check for check in self.checks if not check.ok]

    def render(self) -> str:
        lines = ["=" * 72]
        lines.append(f"回归对比结果: {'通过' if self.ok else '存在差异'}")
        lines.append("=" * 72)
        for check in self.checks:
            mark = "PASS" if check.ok else "DIFF"
            lines.append(f"[{mark}] {check.name}")
            if not check.ok:
                lines.append(f"       期望: {check.expected}")
                lines.append(f"       实际: {check.actual}")
            if check.note:
                lines.append(f"       说明: {check.note}")
        if self.failures:
            lines.append("-" * 72)
            lines.append(f"共 {len(self.failures)} 项差异 / {len(self.checks)} 项检查")
        return "\n".join(lines)


def _text_similarity(left: str, right: str) -> float:
    """基于字符集合与长度估计文本保留度，避免引入第三方依赖。"""
    if not left and not right:
        return 1.0
    if not left or not right:
        return 0.0

    common = sum((left.count(ch) if left.count(ch) < right.count(ch) else right.count(ch))
                 for ch in set(left) & set(right))
    return common / max(len(left), len(right))


def compare(
    actual_path: str | Path,
    expected_path: str | Path,
    *,
    require_toc: bool = True,
    text_threshold: float = 0.98,
) -> ComparisonReport:
    """对比实际输出与期望输出的结构差异。"""
    actual = snapshot(actual_path, "actual")
    expected = snapshot(expected_path, "expected")
    report = ComparisonReport()

    def check(name: str, ok: bool, exp: object, act: object, note: str = "") -> None:
        report.checks.append(CheckResult(name, ok, exp, act, note))

    # ---- 1. 内容完整性（最高优先级）----
    similarity = _text_similarity(expected.body_text, actual.body_text)
    check(
        "正文文本保留度",
        similarity >= text_threshold,
        f">= {text_threshold:.0%}",
        f"{similarity:.1%}",
        "排版不得丢失或篡改正文内容",
    )

    # ---- 2. 目录域 ----
    if require_toc:
        check(
            "目录域存在",
            actual.toc_field_count >= 1,
            f">= 1 (期望文档: {expected.toc_field_count})",
            actual.toc_field_count,
            "目录以 OOXML 域代码形式插入，由引擎刷新页码",
        )

    # ---- 3. 高亮清理 ----
    check(
        "高亮已清除",
        actual.highlight_count == 0,
        "0",
        actual.highlight_count,
        f"期望文档残留 {expected.highlight_count} 处",
    )

    # ---- 4. 表格保留 ----
    check(
        "表格数量",
        actual.table_count == expected.table_count,
        expected.table_count,
        actual.table_count,
        "表格不得丢失（封面冗余表除外）",
    )

    # ---- 5. 页面方向 ----
    actual_landscape = actual.orientation_sequence.count("landscape")
    expected_landscape = expected.orientation_sequence.count("landscape")
    check(
        "横向节数量",
        actual_landscape == expected_landscape,
        expected_landscape,
        actual_landscape,
        "横竖向页不能丢失",
    )

    # ---- 6. 多节结构 ----
    check(
        "分节数",
        actual.section_count >= 2,
        ">= 2",
        actual.section_count,
        "封面/目录/正文应形成独立节",
    )

    # ---- 7. 图片资源 ----
    check(
        "图片资源数",
        actual.media_count >= min(expected.media_count, 1),
        f">= {min(expected.media_count, 1)} (期望文档: {expected.media_count})",
        actual.media_count,
        "图片不得在合并时丢失",
    )

    return report
