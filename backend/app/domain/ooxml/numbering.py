"""自动编号转文本。

对应原 utils/remove_auto_numbering.py（NumberingConverter）。

保留原算法思路：解析 numbering.xml 的 abstractNum / num 映射，
按 ilvl 递增计数器还原出 "%1.%2" 之类的可见文本，写入 run 开头并移除 w:numPr。

改进点：
- 原实现用正则逐个匹配 <w:p> 块，遇到嵌套段落或属性换行会漏处理；
  这里改为 lxml 结构化遍历
- 计数器独立成 NumCounter 类，便于单测
"""

from __future__ import annotations

from dataclasses import dataclass, field

from lxml import etree

from app.core.logging import get_logger
from app.core.namespaces import PART_NUMBERING, qn, w_attr
from app.domain.ooxml.package import DocxPackage

logger = get_logger(__name__)

_MAX_LEVELS = 9


@dataclass
class _LevelDef:
    start: int = 1
    num_fmt: str = "decimal"
    lvl_text: str = "%1"


@dataclass
class _NumCounter:
    """按 numId 维护的多级计数器。"""

    values: dict[str, list[int]] = field(default_factory=dict)

    def next_value(self, num_id: str, level: int, definition: _LevelDef) -> int:
        counters = self.values.setdefault(num_id, [0] * _MAX_LEVELS)

        # 重置所有更深的层级，保证 1.1 → 1.2 而非 1.1 → 1.2.1 递增
        for deeper in range(level + 1, _MAX_LEVELS):
            counters[deeper] = 0

        if counters[level] == 0:
            counters[level] = definition.start
        else:
            counters[level] += 1
        return counters[level]


class NumberingProcessor:
    """把自动编号固化为普通文本，避免跨模板合并后编号错乱。"""

    def __init__(self, pkg: DocxPackage) -> None:
        self.pkg = pkg
        self._abstract: dict[str, dict[int, _LevelDef]] = {}
        self._num_to_abstract: dict[str, str] = {}
        self._counter = _NumCounter()
        self._parse_numbering()

    # ------------------------------------------------------------------
    # 解析
    # ------------------------------------------------------------------
    def _parse_numbering(self) -> None:
        if not self.pkg.has(PART_NUMBERING):
            return

        root = self.pkg.xml(PART_NUMBERING)

        for abstract in root.findall(qn("w:abstractNum")):
            abstract_id = abstract.get(w_attr("abstractNumId"))
            if abstract_id is None:
                continue
            levels: dict[int, _LevelDef] = {}
            for lvl in abstract.findall(qn("w:lvl")):
                ilvl_raw = lvl.get(w_attr("ilvl"))
                try:
                    ilvl = int(ilvl_raw or 0)
                except ValueError:
                    continue
                levels[ilvl] = self._parse_level(lvl)
            self._abstract[abstract_id] = levels

        for num in root.findall(qn("w:num")):
            num_id = num.get(w_attr("numId"))
            ref = num.find(qn("w:abstractNumId"))
            if num_id and ref is not None:
                self._num_to_abstract[num_id] = ref.get(w_attr("val")) or ""

    @staticmethod
    def _parse_level(lvl: etree._Element) -> _LevelDef:
        start_el = lvl.find(qn("w:start"))
        fmt_el = lvl.find(qn("w:numFmt"))
        text_el = lvl.find(qn("w:lvlText"))

        start = 1
        if start_el is not None:
            try:
                start = max(1, int(start_el.get(w_attr("val")) or 1))
            except ValueError:
                start = 1

        return _LevelDef(
            start=start,
            num_fmt=(fmt_el.get(w_attr("val")) if fmt_el is not None else None) or "decimal",
            lvl_text=(text_el.get(w_attr("val")) if text_el is not None else None) or "%1",
        )

    # ------------------------------------------------------------------
    # 转换
    # ------------------------------------------------------------------
    def build_number_text(self, num_id: str, level: int) -> str:
        abstract_id = self._num_to_abstract.get(num_id)
        if abstract_id is None:
            return ""

        levels = self._abstract.get(abstract_id)
        if not levels:
            return ""

        definition = levels.get(level)
        if definition is None:
            return ""

        current = self._counter.next_value(num_id, level, definition)

        if definition.num_fmt == "bullet":
            return f"{definition.lvl_text} "

        # 补齐上级编号（如 2.1 中的 2）
        counters = self._counter.values.setdefault(num_id, [0] * _MAX_LEVELS)
        for upper in range(level):
            if counters[upper] == 0:
                upper_def = levels.get(upper)
                counters[upper] = upper_def.start if upper_def else 1

        text = definition.lvl_text
        for index in range(1, _MAX_LEVELS + 1):
            placeholder = f"%{index}"
            if placeholder in text:
                text = text.replace(placeholder, str(counters[index - 1]))
        return f"{text} "

    def convert(self) -> int:
        """把正文段落中的自动编号转成文本，返回转换段落数。"""
        if not self._num_to_abstract:
            return 0

        converted = 0
        for paragraph in self.pkg.paragraphs():
            ppr = paragraph.find(qn("w:pPr"))
            if ppr is None:
                continue
            num_pr = ppr.find(qn("w:numPr"))
            if num_pr is None:
                continue

            num_id_el = num_pr.find(qn("w:numId"))
            ilvl_el = num_pr.find(qn("w:ilvl"))
            if num_id_el is None:
                continue

            num_id = num_id_el.get(w_attr("val")) or ""
            try:
                level = int(ilvl_el.get(w_attr("val")) or 0) if ilvl_el is not None else 0
            except ValueError:
                level = 0

            number_text = self.build_number_text(num_id, level)

            ppr.remove(num_pr)
            if number_text:
                self._prepend_text_run(paragraph, number_text)
            converted += 1

        if converted:
            logger.info("自动编号转文本 %d 段", converted)
        return converted

    @staticmethod
    def _prepend_text_run(paragraph: etree._Element, text: str) -> None:
        """在第一个 run 之前插入编号文本。"""
        run = etree.Element(qn("w:r"))
        t = etree.SubElement(run, qn("w:t"))
        t.set(qn("xml:space"), "preserve")
        t.text = text

        first_run = paragraph.find(qn("w:r"))
        if first_run is not None:
            first_run.addprevious(run)
        else:
            paragraph.append(run)
