"""标题识别与规范化。

对应原 service/converter.py 的 `step0_preprocess_headings_new`（约 190 行）。

这是整个排版流程的**关键前置能力**：源文档（尤其来自 WPS 的稿件）
的标题几乎都是普通段落（如 "1 总论"、"1.1 项目背景"），
没有应用任何 Heading 样式。若不识别并赋样式：

- 目录域会生成空目录（TOC 只收集标题样式段落）
- 页眉页脚的分节判断失效
- 无法做标题级别的格式统一

原算法全部保留，只把实现从 python-docx 对象树改到 lxml，
以便在同一份 DocxPackage 上连续操作、不落地临时文件。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from lxml import etree

from app.core.constants import layout_rules
from app.core.logging import get_logger
from app.core.namespaces import NSMAP, qn, w_attr
from app.domain.ooxml.package import DocxPackage, paragraph_text
from app.domain.ooxml.styles import StyleProcessor

logger = get_logger(__name__)

# 标题编号正则：1 / 1.2 / 1.2.3 ...，且编号后紧跟的非空白字符不能是 ) . 、 -
# 这条排除规则很关键：正文里 "1." "1、" 之类的列表项不应被误判为标题
_HEADING_PATTERN = re.compile(r"^(\d+(?:\.\d+){0,4})[\s\t]*(?![)）\.、\-])(.+)$")

# 目录标题识别：允许"目 录"、"目      录"等任意空白
_TOC_PATTERN = re.compile(r"目[\s\t]*录")

# 中文括号平衡校验用
_BRACKETS = (("(", ")"), ("（", "）"))


def is_valid_heading_order(prev_number: str, curr_number: str) -> bool:
    """判断标题编号顺序是否合法。

    规则（与原实现一致）：
    - 同级或深层递增只能 +1，不能跳级（1 → 3 非法）
    - 上位级别递增时，后续各级必须归零（1.9 → 2.0 合法，1.9 → 2.1 非法）
    """
    try:
        prev = [int(x) for x in prev_number.split(".")]
        curr = [int(x) for x in curr_number.split(".")]
    except ValueError:
        return False

    max_len = max(len(prev), len(curr))
    prev += [0] * (max_len - len(prev))
    curr += [0] * (max_len - len(curr))

    for index in range(max_len):
        if prev[index] < curr[index]:
            return curr[index] - prev[index] in (0, 1)
        if prev[index] > curr[index]:
            if index > 0 and prev[index - 1] + 1 == curr[index - 1]:
                if sum(curr[index:]) == 0:
                    return True
            return False
    return True


def _balanced_brackets(text: str) -> bool:
    """右括号多于左括号即视为正文，不是标题（如 "1) 说明" 这种列表项）。"""
    extra = 0
    for opening, closing in _BRACKETS:
        extra += text.count(closing) - text.count(opening)
    return extra < 1


@dataclass
class HeadingReport:
    """标题识别结果。"""

    detected: int = 0
    reset_to_normal: int = 0
    levels: dict[int, int] = field(default_factory=dict)
    skipped_first_level1: list[str] = field(default_factory=list)
    invalid_order: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)


class HeadingProcessor:
    """按编号规则识别标题并赋予 Heading 样式。"""

    def __init__(self, pkg: DocxPackage) -> None:
        self.pkg = pkg
        self.styles = StyleProcessor(pkg)
        rules = layout_rules()["headings"]
        self.max_level = int(rules["max_level"])
        self.space_before = float(rules["space_before_pt"])
        self.space_after = float(rules["space_after_pt"])
        self._style_ready: dict[int, str] = {}

    # ------------------------------------------------------------------
    def _style_for_level(self, level: int) -> str | None:
        """取该级别的标题 styleId；超过配置上限时回落到最高级。"""
        effective = min(level, self.max_level)
        if effective in self._style_ready:
            return self._style_ready[effective]

        style_id = self.styles.resolver.heading_style_id(effective)
        if style_id is None:
            return None

        self._style_ready[effective] = style_id
        return style_id

    # ------------------------------------------------------------------
    def apply_heading_styles(self) -> HeadingReport:
        """识别并标注标题。

        规则链（与原实现逐条对应）：
        1. 目录之前不处理
        2. 跳过第一个 "1" 级标题（通常是旧目录里的第一条，或正文首页大标题）
        3. 从第二个 "1" 级标题开始正式处理
        4. 编号顺序不合法 → 视为正文
        5. 括号不平衡 → 视为正文
        6. 含 w:pict 的段落（文本框/图）→ 视为正文
        7. 反向纠错：已是标题样式但文本不符合编号规则 → 还原为正文
        """
        report = HeadingReport()

        after_toc = False
        level1_count = 0
        processing_enabled = False
        prev_number: str | None = None

        paragraphs = self.pkg.paragraphs()

        for index, paragraph in enumerate(paragraphs):
            text = paragraph_text(paragraph)
            if not text:
                continue

            if not after_toc:
                if _TOC_PATTERN.search(text):
                    after_toc = True
                    logger.debug("识别到目录段落: %s", text[:30])
                continue

            # 含图形对象的段落不参与标题识别
            if paragraph.find(f".//{qn('w:pict')}") is not None:
                continue

            match = _HEADING_PATTERN.match(text)
            if not match:
                continue

            if not _balanced_brackets(text):
                report.notes.append(f"括号不平衡，跳过: {text[:30]}")
                continue

            number = match.group(1)
            level = number.count(".") + 1

            # 定位正文起点：跳过第 1 个一级标题，处理从第 2 个开始
            if not processing_enabled:
                if number == "1" and level == 1:
                    level1_count += 1
                    if level1_count == 1:
                        report.skipped_first_level1.append(text[:40])
                        continue
                    processing_enabled = True
                    prev_number = None
                else:
                    continue

            if prev_number is not None and not is_valid_heading_order(prev_number, number):
                report.invalid_order.append(f"{number} (前 {prev_number})")
                self.styles.set_paragraph_style(
                    paragraph,
                    self._normal_style_id(),
                    space_before_pt=self.space_before,
                    space_after_pt=self.space_after,
                )
                continue

            style_id = self._style_for_level(level)
            if style_id is None:
                report.notes.append(f"样式表缺少 Heading {level}，跳过: {text[:30]}")
                continue

            self.styles.set_paragraph_style(
                paragraph,
                style_id,
                space_before_pt=self.space_before,
                space_after_pt=self.space_after,
            )
            report.detected += 1
            report.levels[level] = report.levels.get(level, 0) + 1
            prev_number = number

        report.reset_to_normal = self._reverse_correct()

        logger.info(
            "标题识别完成: 标注 %d 处, 纠错 %d 处, 级别分布 %s",
            report.detected,
            report.reset_to_normal,
            report.levels,
        )
        if not report.detected:
            logger.warning(
                "未识别到任何标题，目录将为空。"
                "请检查文档是否含「目录」标题以及编号是否符合 1 / 1.1 / 1.1.1 格式"
            )
        return report

    # ------------------------------------------------------------------
    def _normal_style_id(self) -> str:
        return self.styles.resolver.id_of("normal") or self.styles.resolver.id_of("正文") or "Normal"

    def _reverse_correct(self) -> int:
        """反向纠错：已是标题样式但文本不符合编号规则的段落还原为正文。"""
        corrected = 0
        for paragraph in self.pkg.paragraphs():
            style_id = self.styles.paragraph_style_id(paragraph)
            if not style_id or not self.styles.resolver.is_heading(style_id):
                continue
            if _HEADING_PATTERN.match(paragraph_text(paragraph)):
                continue

            self.styles.set_paragraph_style(
                paragraph,
                self._normal_style_id(),
                space_before_pt=self.space_before,
                space_after_pt=self.space_after,
            )
            corrected += 1
        return corrected

    # ------------------------------------------------------------------
    def heading_count(self) -> dict[int, int]:
        """统计当前文档中各级标题数量（供校验与报告使用）。"""
        counts: dict[int, int] = {}
        for paragraph in self.pkg.paragraphs():
            style_id = self.styles.paragraph_style_id(paragraph)
            if not style_id:
                continue
            level = self.styles.resolver.heading_level(style_id)
            if level:
                counts[level] = counts.get(level, 0) + 1
        return counts
