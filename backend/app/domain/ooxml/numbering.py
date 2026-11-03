"""自动编号转文本。

对应原 utils/remove_auto_numbering.py（NumberingConverter）。

保留原算法思路：解析 numbering.xml 的 abstractNum / num 映射，
按 ilvl 递增计数器还原出 "%1.%2" 之类的可见文本，写入 run 开头并移除 w:numPr。

改进点：
- 原实现用正则逐个匹配 <w:p> 块，遇到嵌套段落或属性换行会漏处理；
    10|  这里改为 lxml 结构化遍历
- 计数器独立成 NumCounter 类，便于单测

**还原口径以 Word 实际渲染为准**（用 Word COM 读取 ListString 逐段对照过），
因此下面几条 Word 语义必须实现，否则还原出来的编号和用户看到的不一样：

1. 编号状态挂在 **abstractNumId** 上，不是 numId。多个 numId 指向同一套抽象编号时
   共享同一条序列。典型场景：标题样式挂 numId=1、四级标题段落挂 numId=7，
   两者都指向 abstract 2，所以「11.2.3」下的四级标题会接着用 7.2.9 用过的 1~6，
   显示为 11.2.3.7，而不是重新从 1 起。
2. `w:lvlRestart` 控制重置：缺省表示上层变了就重置；`val=0` 表示**永不重置**
   （上面的 11.2.3.7 就是靠这个继续累加的）。
3. `w:numFmt` 决定数字形态：decimalEnclosedCircleChinese 是 ①，
   lowerLetter / lowerRoman / chineseCounting 各有一套写法，`none` 表示不显示编号。
4. `w:numId val=0` 表示这一段**显式取消编号**，不能再回落到样式上的编号。
5. `w:suff` 决定编号与正文之间的分隔：`nothing` 不加分隔符，其余按配置补一个空格。
"""

from __future__ import annotations

from dataclasses import dataclass, field

from lxml import etree

from app.core.constants import layout_rules
from app.core.logging import get_logger
from app.core.namespaces import PART_NUMBERING, PART_STYLES, qn, w_attr
from app.domain.ooxml.package import DocxPackage

logger = get_logger(__name__)

_MAX_LEVELS = 9

# 编号形态的字符表。这些是 OOXML numFmt 的固定枚举，不属于可调业务配置。
_CIRCLED_DIGITS = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"
_CN_DIGITS = "零一二三四五六七八九"
_ROMAN_UNITS = (
    (1000, "M"),
    (900, "CM"),
    (500, "D"),
    (400, "CD"),
    (100, "C"),
    (90, "XC"),
    (50, "L"),
    (40, "XL"),
    (10, "X"),
    (9, "IX"),
    (5, "V"),
    (4, "IV"),
    (1, "I"),
)
_CHINESE_FMTS = frozenset(
    {
        "chineseCounting",
        "chineseCountingThousand",
        "chineseLegalSimplified",
        "japaneseCounting",
        "ideographDigital",
        "taiwaneseCounting",
        "taiwaneseCountingThousand",
    }
)
_ROMAN_FMTS = frozenset({"lowerRoman", "upperRoman"})
_LETTER_FMTS = frozenset({"lowerLetter", "upperLetter"})
_CIRCLED_FMTS = frozenset({"decimalEnclosedCircle", "decimalEnclosedCircleChinese"})


def _chinese_number(value: int) -> str:
    """1 → 一，11 → 十一，21 → 二十一。超出 99 时退回阿拉伯数字。"""
    if value < 0 or value > 99:
        return str(value)
    if value < 10:
        return _CN_DIGITS[value]
    tens, ones = divmod(value, 10)
    head = "十" if tens == 1 else f"{_CN_DIGITS[tens]}十"
    return head if ones == 0 else f"{head}{_CN_DIGITS[ones]}"


def _roman_number(value: int, upper: bool) -> str:
    if value < 1 or value > 3999:
        return str(value)
    rest, result = value, []
    for amount, symbol in _ROMAN_UNITS:
        while rest >= amount:
            rest -= amount
            result.append(symbol)
    text = "".join(result)
    return text if upper else text.lower()


def _letter_number(value: int, upper: bool) -> str:
    """1 → a，26 → z，27 → aa（Word 的重复字母写法）。"""
    if value < 1:
        return str(value)
    letter = chr(ord("a") + (value - 1) % 26)
    repeats = (value - 1) // 26 + 1
    text = letter * repeats
    return text.upper() if upper else text


def format_number_value(value: int, num_fmt: str) -> str:
    """按 OOXML numFmt 把计数器数值渲染成 Word 里看到的形态。"""
    if num_fmt in ("none", ""):
        return ""
    if num_fmt in _CIRCLED_FMTS:
        if 1 <= value <= len(_CIRCLED_DIGITS):
            return _CIRCLED_DIGITS[value - 1]
        return str(value)
    if num_fmt in _CHINESE_FMTS:
        return _chinese_number(value)
    if num_fmt in _ROMAN_FMTS:
        return _roman_number(value, num_fmt == "upperRoman")
    if num_fmt in _LETTER_FMTS:
        return _letter_number(value, num_fmt == "upperLetter")
    return str(value)


def format_number_label(number: str, suff: str | None = None) -> str:
    """编号和后面的文字之间只留配置里的那一个分隔符。

    编号定义的 lvlText 经常已经带一个空格，Word 的 suff 还会再加一个。
    这里先去掉末尾空白，再只补一次分隔符，避免标题变成「1  总论」。

    ``suff == "nothing"`` 表示 Word 里编号与正文紧贴（如「（1）配料室」），
    此时不补分隔符，保持原样。
    """
    cleaned = number.rstrip(" \t")
    if not cleaned:
        return ""
    if suff == "nothing":
        return cleaned
    separator = str(layout_rules()["headings"].get("number_separator", " "))
    return f"{cleaned}{separator}"


@dataclass
class _LevelDef:
    start: int = 1
    num_fmt: str = "decimal"
    lvl_text: str = "%1"
    # None 表示 w:lvlRestart 缺省（上层变化即重置），0 表示永不重置
    lvl_restart: int | None = None
    suff: str | None = None


@dataclass
class _NumCounter:
    """按 abstractNumId 维护的多级计数器（Word 的编号状态挂在抽象编号上）。"""

    values: dict[str, list[int]] = field(default_factory=dict)
    start_overrides: dict[tuple[str, int], int] = field(default_factory=dict)

    def next_value(
        self,
        abstract_id: str,
        num_id: str,
        level: int,
        levels: dict[int, "_LevelDef"],
    ) -> list[int]:
        counters = self.values.setdefault(abstract_id, [0] * _MAX_LEVELS)
        definition = levels[level]

        if counters[level] == 0:
            override = self.start_overrides.get((num_id, level))
            counters[level] = override if override is not None else definition.start
        else:
            counters[level] += 1

        self._reset_deeper(counters, levels, level)
        return counters

    @staticmethod
    def _reset_deeper(counters: list[int], levels: dict[int, "_LevelDef"], level: int) -> None:
        """上层编号变化时重置更深的层级，规则由 w:lvlRestart 决定。"""
        for deeper in range(level + 1, _MAX_LEVELS):
            definition = levels.get(deeper)
            restart = definition.lvl_restart if definition is not None else None
            if restart == 0:
                continue
            # 缺省：出现更浅的层级就重置；显式 val=N：0 基层级 < N 时重置
            limit = restart if restart is not None else deeper
            if level < limit:
                counters[deeper] = 0


class NumberingProcessor:
    """把自动编号固化为普通文本，避免跨模板合并后编号错乱。"""

    def __init__(self, pkg: DocxPackage) -> None:
        self.pkg = pkg
        self._abstract: dict[str, dict[int, _LevelDef]] = {}
        self._num_to_abstract: dict[str, str] = {}
        self._counter = _NumCounter()
        self._style_numpr: dict[str, tuple[str, int] | None] = {}
        self._style_based: dict[str, str | None] = {}
        self._consumed_styles: set[str] = set()
        self._parse_numbering()
        self._parse_style_numbering()

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
            if not num_id or ref is None:
                continue
            self._num_to_abstract[num_id] = ref.get(w_attr("val")) or ""
            # w:lvlOverride/w:startOverride：这一套编号从指定值重新开始
            for override in num.findall(qn("w:lvlOverride")):
                try:
                    level = int(override.get(w_attr("ilvl")) or 0)
                except ValueError:
                    continue
                start = override.find(qn("w:startOverride"))
                if start is None:
                    continue
                try:
                    self._counter.start_overrides[(num_id, level)] = int(
                        start.get(w_attr("val")) or ""
                    )
                except ValueError:
                    continue

    def _parse_style_numbering(self) -> None:
        """有些文档把编号定义挂在样式上（如标题 1 → 自动编号），段落本身没有 numPr。"""
        if not self.pkg.has(PART_STYLES):
            return

        root = self.pkg.xml(PART_STYLES)
        for style in root.findall(qn("w:style")):
            style_id = style.get(w_attr("styleId")) or ""
            if not style_id:
                continue

            based = style.find(qn("w:basedOn"))
            self._style_based[style_id] = (
                None if based is None else based.get(w_attr("val"))
            )

            num_pr = style.find(f"{qn('w:pPr')}/{qn('w:numPr')}")
            if num_pr is None:
                continue
            num_id_el = num_pr.find(qn("w:numId"))
            ilvl_el = num_pr.find(qn("w:ilvl"))
            if num_id_el is None:
                continue
            num_id = num_id_el.get(w_attr("val")) or ""
            if not num_id or num_id == "0":
                continue
            try:
                level = int(ilvl_el.get(w_attr("val")) or 0) if ilvl_el is not None else 0
            except ValueError:
                level = 0
            self._style_numpr[style_id] = (num_id, level)

    def _resolve_style_numpr(self, style_id: str) -> tuple[str, int] | None:
        """沿 basedOn 链找样式上的编号定义。"""
        seen: set[str] = set()
        current: str | None = style_id
        while current and current not in seen:
            seen.add(current)
            if current in self._style_numpr:
                return self._style_numpr[current]
            current = self._style_based.get(current)
        return None

    @staticmethod
    def _parse_level(lvl: etree._Element) -> _LevelDef:
        start_el = lvl.find(qn("w:start"))
        fmt_el = lvl.find(qn("w:numFmt"))
        text_el = lvl.find(qn("w:lvlText"))
        restart_el = lvl.find(qn("w:lvlRestart"))
        suff_el = lvl.find(qn("w:suff"))

        start = 1
        if start_el is not None:
            try:
                start = max(1, int(start_el.get(w_attr("val")) or 1))
            except ValueError:
                start = 1

        restart: int | None = None
        if restart_el is not None:
            try:
                restart = max(0, int(restart_el.get(w_attr("val")) or 0))
            except ValueError:
                restart = None

        return _LevelDef(
            start=start,
            num_fmt=(fmt_el.get(w_attr("val")) if fmt_el is not None else None) or "decimal",
            lvl_text=(text_el.get(w_attr("val")) if text_el is not None else None) or "%1",
            lvl_restart=restart,
            suff=(suff_el.get(w_attr("val")) if suff_el is not None else None),
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

        if definition.num_fmt == "bullet":
            return format_number_label(definition.lvl_text, definition.suff)
        if definition.num_fmt == "none":
            return ""

        counters = self._counter.next_value(abstract_id, num_id, level, levels)

        # 补齐上级编号（如 2.1 中的 2）
        for upper in range(level):
            if counters[upper] == 0:
                upper_def = levels.get(upper)
                counters[upper] = upper_def.start if upper_def else 1

        text = definition.lvl_text
        for index in range(1, _MAX_LEVELS + 1):
            placeholder = f"%{index}"
            if placeholder not in text:
                continue
            upper_def = levels.get(index - 1)
            fmt = upper_def.num_fmt if upper_def is not None else "decimal"
            text = text.replace(
                placeholder, format_number_value(counters[index - 1], fmt)
            )
        return format_number_label(text, definition.suff)

    def convert(self) -> int:
        """把正文段落中的自动编号转成文本，返回转换段落数。

        编号可能直接写在段落上，也可能来自段落所用的样式（标题样式常见）。
        样式提供的编号要把样式上的 numPr 一并摘掉，否则会重复编号。

        必须**按文档顺序单次遍历**：计数器依赖先后顺序，分两趟走（先段落级、
        再样式级）会把编号序列算乱。
        """
        if not self._num_to_abstract:
            return 0

        converted = self._walk()
        self._strip_style_numbering()

        if converted:
            logger.info("自动编号转文本 %d 段", converted)
        return converted

    def _walk(self) -> int:
        converted = 0
        for paragraph in self.pkg.body.iter(qn("w:p")):
            ppr = paragraph.find(qn("w:pPr"))
            if ppr is None:
                continue

            num_pr = ppr.find(qn("w:numPr"))
            if num_pr is not None:
                num_id, level = self._numpr_values(num_pr)
                ppr.remove(num_pr)
                # numId=0 是显式取消编号，不能回落到样式上的编号
                if num_id and num_id != "0":
                    number_text = self.build_number_text(num_id, level)
                    if number_text:
                        self._prepend_text_run(paragraph, number_text)
                        converted += 1
                continue

            style_el = ppr.find(qn("w:pStyle"))
            if style_el is None:
                continue
            style_id = style_el.get(w_attr("val")) or ""
            resolved = self._resolve_style_numpr(style_id)
            if resolved is None:
                continue

            num_id, level = resolved
            number_text = self.build_number_text(num_id, level)
            if not number_text:
                continue
            self._prepend_text_run(paragraph, number_text)
            self._consumed_styles.add(style_id)
            converted += 1
        return converted

    def _strip_style_numbering(self) -> None:
        """已经把编号写成文本，样式上的编号必须摘掉，否则会重复。"""
        if not self._consumed_styles or not self.pkg.has(PART_STYLES):
            return

        root = self.pkg.xml(PART_STYLES)
        for style in root.findall(qn("w:style")):
            if (style.get(w_attr("styleId")) or "") not in self._consumed_styles:
                continue
            ppr = style.find(qn("w:pPr"))
            if ppr is None:
                continue
            num_pr = ppr.find(qn("w:numPr"))
            if num_pr is not None:
                ppr.remove(num_pr)

    @staticmethod
    def _numpr_values(num_pr: etree._Element) -> tuple[str, int]:
        """取出 numPr 上的 numId 与 ilvl。"""
        num_id_el = num_pr.find(qn("w:numId"))
        ilvl_el = num_pr.find(qn("w:ilvl"))
        num_id = num_id_el.get(w_attr("val")) or "" if num_id_el is not None else ""
        try:
            level = int(ilvl_el.get(w_attr("val")) or 0) if ilvl_el is not None else 0
        except ValueError:
            level = 0
        return num_id, level

    @staticmethod
    def _strip_leading_gap(paragraph: etree._Element) -> None:
        """编号已经自带分隔符，正文开头的空格要去掉，否则会叠成两个空格。"""
        for node in paragraph.iter(qn("w:t")):
            if not node.text:
                continue
            node.text = node.text.lstrip(" \t")
            if node.text:
                return

    @staticmethod
    def _prepend_text_run(paragraph: etree._Element, text: str) -> None:
        """在第一个 run 之前插入编号文本。"""
        NumberingProcessor._strip_leading_gap(paragraph)
        run = etree.Element(qn("w:r"))
        t = etree.SubElement(run, qn("w:t"))
        t.set(qn("xml:space"), "preserve")
        t.text = text

        first_run = paragraph.find(qn("w:r"))
        if first_run is not None:
            first_run.addprevious(run)
        else:
            paragraph.append(run)
