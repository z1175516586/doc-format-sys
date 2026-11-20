"""为左右对比提取 DOCX 文本和常用排版属性。"""

from __future__ import annotations

from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
from typing import Literal

from lxml import etree

from app.core.namespaces import PART_STYLES, qn, w_attr
from app.domain.ooxml.diff import normalize_text
from app.domain.ooxml.package import DocxPackage
from app.domain.ooxml.styles import StyleResolver
from app.schemas.comparison import (
    ComparisonRow,
    ComparisonSide,
    DocumentComparison,
    StyleChange,
    TextFragment,
)

_PARAGRAPH_FIELDS = {
    "jc": {"val": "对齐"},
    "outlineLvl": {"val": "大纲级别"},
    "spacing": {"before": "段前", "after": "段后", "line": "行距", "lineRule": "行距规则"},
    "ind": {"left": "左缩进", "right": "右缩进", "firstLine": "首行缩进", "hanging": "悬挂缩进"},
    "shd": {"fill": "段落底纹"},
    "numPr": {},
}
_RUN_FIELDS = {
    "rFonts": {"eastAsia": "中文字体", "ascii": "西文字体", "eastAsiaTheme": "中文主题字体",
               "asciiTheme": "西文主题字体"},
    "sz": {"val": "字号"},
    "color": {"val": "文字颜色"},
    "highlight": {"val": "高亮"},
    "shd": {"fill": "文字底纹"},
    "vertAlign": {"val": "文字位置"},
    "spacing": {"val": "字距"},
}
_RUN_SWITCHES = {"b": "加粗", "i": "斜体", "u": "下划线", "strike": "删除线"}
_FALSE = {"0", "false", "off", "none"}
_DEFAULT_RUN = {label: "无" for label in _RUN_SWITCHES.values()}
_DEFAULT_RUN["高亮"] = "无"
_DEFAULT_RUN["文字颜色"] = "auto"
_DEFAULT_PARAGRAPH = dict.fromkeys(
    ("段前", "段后", "左缩进", "右缩进", "首行缩进", "悬挂缩进"), "0"
) | {"对齐": "left", "大纲级别": "9", "行距规则": "auto", "编号方式": "无"}
_ALIGNMENTS = {"left": "左对齐", "start": "左对齐", "center": "居中", "right": "右对齐",
               "end": "右对齐", "both": "两端对齐", "justify": "两端对齐",
               "distribute": "分散对齐"}
_LINE_RULES = {"auto": "多倍", "exact": "固定值", "atLeast": "最小值"}
_TEXT_NODES = {qn("w:t"), qn("w:tab"), qn("w:br"), qn("w:cr")}
_SKIP_CONTAINERS = {qn(tag) for tag in ("w:del", "w:moveFrom", "w:txbxContent")}
_MAX_INLINE_DIFF = 800


def _find(node: etree._Element | None, path: str) -> etree._Element | None:
    return node.find("/".join(qn(tag) for tag in path.split("/"))) if node is not None else None


def _val(node: etree._Element | None, attr: str = "val") -> str | None:
    return node.get(w_attr(attr)) if node is not None else None


def _points(value: str, divisor: int) -> str:
    try:
        return f"{int(value) / divisor:g} 磅"
    except ValueError:
        return value


def _properties(node: etree._Element | None, *, run: bool = False) -> dict[str, str]:
    fields = _RUN_FIELDS if run else _PARAGRAPH_FIELDS
    values: dict[str, str] = {}
    for tag, attributes in fields.items():
        child = _find(node, f"w:{tag}")
        if child is None:
            continue
        for attr, label in attributes.items():
            if (value := _val(child, attr)) is not None:
                values[label] = value
        if tag == "numPr":
            number_id = _val(_find(child, "w:numId"))
            if number_id is not None:
                level = _val(_find(child, "w:ilvl")) or "0"
                values["编号方式"] = "无" if number_id == "0" else f"自动编号（层级 {level}）"

    if run:
        for tag, label in _RUN_SWITCHES.items():
            flag = _find(node, f"w:{tag}")
            if flag is not None:
                value = _val(flag) or "single"
                values[label] = "无" if value in _FALSE else (
                    "有" if tag != "u" or value == "single" else value
                )
    return values


def _display_style(label: str, value: str) -> str:
    if value == "未设置":
        return value
    if label in {"段前", "段后", "左缩进", "右缩进", "首行缩进", "悬挂缩进", "字距"}:
        return _points(value, 20)
    if label == "字号":
        return _points(value, 2)
    if label == "对齐":
        return _ALIGNMENTS.get(value, value)
    if label == "行距规则":
        return _LINE_RULES.get(value, value)
    if label == "大纲级别":
        return "正文" if value == "9" else f"级别 {int(value) + 1}" if value.isdecimal() else value
    if label in {"文字颜色", "段落底纹", "文字底纹"}:
        return "自动" if value.lower() == "auto" else f"#{value.upper()}"
    return "无" if label == "高亮" and value == "none" else value


class _StyleReader:
    def __init__(self, package: DocxPackage) -> None:
        self.resolver = StyleResolver(package)
        self.styles: dict[str, etree._Element] = {}
        self.default_id = ""
        self.style_cache: dict[str, tuple[dict[str, str], dict[str, str]]] = {}
        root = package.xml(PART_STYLES) if package.has(PART_STYLES) else None
        if root is not None:
            for style in root.findall(qn("w:style")):
                if style_id := _val(style, "styleId"):
                    self.styles[style_id] = style
                    if _val(style, "type") == "paragraph" and _val(style, "default") in ("1", "true", "on"):
                        self.default_id = style_id
        if not self.default_id:
            self.default_id = self.resolver.id_of("Normal") or ""
        self.paragraph_defaults = _DEFAULT_PARAGRAPH | _properties(
            _find(root, "w:docDefaults/w:pPrDefault/w:pPr")
        )
        self.run_defaults = _DEFAULT_RUN | _properties(
            _find(root, "w:docDefaults/w:rPrDefault/w:rPr"), run=True
        )

    def _chain(self, style_id: str) -> tuple[dict[str, str], dict[str, str]]:
        if style_id in self.style_cache:
            return self.style_cache[style_id]
        chain = []
        seen = set()
        current = style_id
        while current in self.styles and current not in seen:
            style = self.styles[current]
            chain.append(style)
            seen.add(current)
            current = _val(_find(style, "w:basedOn")) or ""
        paragraph, run = {}, {}
        for style in reversed(chain):
            paragraph.update(_properties(_find(style, "w:pPr")))
            run.update(_properties(_find(style, "w:rPr"), run=True))
        self.style_cache[style_id] = (paragraph, run)
        return paragraph, run

    def paragraph(self, node: etree._Element) -> tuple[dict[str, str], dict[str, str]]:
        direct = _find(node, "w:pPr")
        style_id = _val(_find(direct, "w:pStyle")) or self.default_id
        inherited, run = self._chain(style_id)
        properties = self.paragraph_defaults | inherited | _properties(direct)
        properties["段落样式"] = self.resolver.name_of(style_id).lower() or style_id
        if line := properties.get("行距"):
            if properties.get("行距规则", "auto") == "auto":
                try:
                    properties["行距"] = f"{int(line) / 240:g} 倍"
                except ValueError:
                    pass
            else:
                properties["行距"] = _points(line, 20)
        return properties, self.run_defaults | run

    def run(self, node: etree._Element, inherited: dict[str, str]) -> dict[str, str]:
        direct = _find(node, "w:rPr")
        style_id = _val(_find(direct, "w:rStyle")) or ""
        _, style = self._chain(style_id)
        return inherited | style | _properties(direct, run=True)


@dataclass
class _Block:
    kind: Literal["paragraph", "table_row"]
    text: str
    styles: dict[str, tuple[tuple[int, str], ...]]


def _visible_runs(paragraph: etree._Element):
    for run in paragraph.iter(qn("w:r")):
        parent = run.getparent()
        while parent is not None and parent is not paragraph and parent.tag not in _SKIP_CONTAINERS:
            parent = parent.getparent()
        if parent is paragraph:
            yield run


def _run_text(run: etree._Element) -> str:
    return "".join(
        (child.text or "") if child.tag == qn("w:t") else "\t" if child.tag == qn("w:tab") else "\n"
        for child in run if child.tag in _TEXT_NODES
    )


def _style_ranges(parts: list[tuple[str, dict[str, str]]]) -> dict[str, tuple[tuple[int, str], ...]]:
    """合并相同格式的相邻 run，避免 Word 拆分 XML 时产生样式差异。"""
    formatted_parts = [(text, style) for text, style in parts if style]
    labels = dict.fromkeys(label for _, style in formatted_parts for label in style)
    result = {}
    for label in labels:
        ranges: list[tuple[int, str]] = []
        for text, style in formatted_parts:
            value = _display_style(label, style.get(label, "未设置"))
            if ranges and ranges[-1][1] == value:
                ranges[-1] = (ranges[-1][0] + len(text), value)
            else:
                ranges.append((len(text), value))
        result[label] = tuple(ranges)
    return result


def _paragraph_parts(paragraph: etree._Element, styles: _StyleReader) -> list[tuple[str, dict[str, str]]]:
    paragraph_style, inherited = styles.paragraph(paragraph)
    return [
        (text, paragraph_style | styles.run(run, inherited))
        for run in _visible_runs(paragraph)
        if (text := _run_text(run))
    ]


def _read_blocks(path: Path) -> list[_Block]:
    def make_block(kind: Literal["paragraph", "table_row"], parts: list[tuple[str, dict[str, str]]]) -> _Block:
        return _Block(kind, "".join(text for text, _ in parts), _style_ranges(parts))

    blocks: list[_Block] = []
    with DocxPackage.open(path) as package:
        styles = _StyleReader(package)
        for node in package.body.iter():
            if node.tag not in (qn("w:p"), qn("w:tr")):
                continue
            # 表格以行对齐，文本框、删除修订不与主文档段落混排。
            if any(parent.tag in _SKIP_CONTAINERS for parent in node.iterancestors()):
                continue
            if node.tag == qn("w:p"):
                if any(True for _ in node.iterancestors(qn("w:tbl"))):
                    continue
                parts = _paragraph_parts(node, styles)
                if any(text.strip() for text, _ in parts):
                    blocks.append(make_block("paragraph", parts))
            else:
                table = node.getparent()
                while table is not None and table.tag != qn("w:tbl"):
                    table = table.getparent()
                table_style = _val(_find(table, "w:tblPr/w:tblStyle"))
                table_name = (styles.resolver.name_of(table_style) or table_style) if table_style else ""
                parts = []
                for cell_index, cell in enumerate(node.findall(qn("w:tc"))):
                    if cell_index:
                        parts.append((" │ ", {}))
                    for index, paragraph in enumerate(cell.findall(qn("w:p"))):
                        if index:
                            parts.append(("\n", {}))
                        parts.extend(_paragraph_parts(paragraph, styles))
                block = make_block("table_row", parts)
                if table_name:
                    block.styles["表格样式"] = ((len(block.text), table_name),)
                if block.text:
                    blocks.append(block)
    return blocks


def _style_changes(before: _Block, after: _Block) -> list[StyleChange]:
    changes = []
    for label in dict.fromkeys((*before.styles, *after.styles)):
        left = before.styles.get(label, ((0, "未设置"),))
        right = after.styles.get(label, ((0, "未设置"),))
        left_values = tuple(dict.fromkeys(value for _, value in left))
        right_values = tuple(dict.fromkeys(value for _, value in right))
        if left_values != right_values or (before.text == after.text and left != right):
            changes.append(StyleChange(
                label=label, before=" / ".join(left_values), after=" / ".join(right_values),
                range_changed=left_values == right_values,
            ))
    return changes


def _fragments(before: str, after: str) -> tuple[list[TextFragment], list[TextFragment]]:
    if before == after:
        return [TextFragment(text=before)], [TextFragment(text=after)]
    if max(len(before), len(after)) > _MAX_INLINE_DIFF:
        return [TextFragment(text=before, changed=True)], [TextFragment(text=after, changed=True)]
    left, right = [], []
    for tag, a, b, c, d in SequenceMatcher(None, before, after, autojunk=False).get_opcodes():
        if a < b:
            left.append(TextFragment(text=before[a:b], changed=tag != "equal"))
        if c < d:
            right.append(TextFragment(text=after[c:d], changed=tag != "equal"))
    return left, right


def _paired(before: _Block, after: _Block, i: int, j: int) -> ComparisonRow:
    styles = _style_changes(before, after)
    left, right = _fragments(before.text, after.text)
    return ComparisonRow(
        kind="modified" if before.text != after.text else "style" if styles else "equal",
        block_type=before.kind,
        before=ComparisonSide(number=i + 1, text=left),
        after=ComparisonSide(number=j + 1, text=right),
        styles=styles,
    )


def _only(block: _Block, index: int, *, added: bool) -> ComparisonRow:
    side = ComparisonSide(number=index + 1, text=[TextFragment(text=block.text, changed=True)])
    return ComparisonRow(kind="added" if added else "removed", block_type=block.kind,
                         before=None if added else side, after=side if added else None)


def compare_documents(before_path: Path, after_path: Path, source_name: str, output_name: str) -> DocumentComparison:
    """先按正文和表格文字对齐，再逐行比较可见文字和样式。"""
    before, after = _read_blocks(before_path), _read_blocks(after_path)
    left = [(block.kind, normalize_text(block.text)) for block in before]
    right = [(block.kind, normalize_text(block.text)) for block in after]
    rows = []
    for tag, i, end_i, j, end_j in SequenceMatcher(None, left, right).get_opcodes():
        while i < end_i and j < end_j:
            if tag == "equal" or (
                before[i].kind == after[j].kind
                and SequenceMatcher(None, left[i][1][:_MAX_INLINE_DIFF],
                                    right[j][1][:_MAX_INLINE_DIFF]).ratio() >= 0.45
            ):
                rows.append(_paired(before[i], after[j], i, j))
            else:
                rows.extend((_only(before[i], i, added=False), _only(after[j], j, added=True)))
            i += 1
            j += 1
        rows.extend(_only(before[k], k, added=False) for k in range(i, end_i))
        rows.extend(_only(after[k], k, added=True) for k in range(j, end_j))
    counts = dict.fromkeys(("equal", "added", "removed", "modified", "style"), 0)
    for row in rows:
        counts[row.kind] += 1
    return DocumentComparison(source_name=source_name, output_name=output_name, rows=rows, counts=counts)
