"""封面处理。

对应原 utils/cover_replace.py + utils/cover_table_replace.py。

原实现把"从封面抽取元数据 → 选模板 → 回填占位符 → 移植设计人员表"
四件事混在 1500 行文件里，并且依赖 win32com 与远端 Dify 服务，
还针对 4 文档类型 × 2 日期型 × 3 长度档维护了 24 份几乎相同的模板。

重构后：

- ``CoverMetadataExtractor``：从原始封面抽取项目名/日期/库号/人员（纯本地规则）
- ``CoverProcessor``：在**唯一**的封面模板上做阶段文字、字段、人员回填，
  再移植设计人员表、调整字间距

模板选择与流程编排放在 ``app/services/cover_service.py``。
所有可调参数来自 ``cover_config()``（backend/config/cover_config.json）。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Iterator

from lxml import etree

from app.core.constants import cover_config, pt_to_half_points
from app.core.logging import get_logger
from app.core.namespaces import qn, w_attr
from app.domain.ooxml.document import RelationshipRemapper, rewrite_relationship_ids
from app.domain.ooxml.package import DocxPackage, element_text, paragraph_text

logger = get_logger(__name__)

# 封面元数据里出现即视为"图形/文本框"，不参与文本识别
_SKIP_TAGS = ("w:drawing", "w:pict", "w:txbxContent")


@dataclass
class CoverMetadata:
    """从原始封面抽取到的元数据。"""

    project_name: str = ""
    date_text: str = ""
    library_number: str = ""
    personnel: list[tuple[str, str]] = field(default_factory=list)

    @property
    def is_empty(self) -> bool:
        return not (self.project_name or self.date_text or self.library_number or self.personnel)


def _compact(text: str) -> str:
    """去所有空白（含全角空格），用于标题/职务/姓名的规范比较。"""
    return re.sub(r"\s+", "", text)


def iter_personnel_lines(
    paragraphs: list[etree._Element],
    role_keywords: tuple[str, ...],
) -> Iterator[tuple[etree._Element, str, str]]:
    """按文档顺序产出「职务：姓名」段落。

    职务命中关键词即视为人员行（用于源封面与模板封面两侧的定位）。
    """
    for paragraph in paragraphs:
        text = paragraph_text(paragraph)
        if not text:
            continue
        parts = re.split(r"[:：]", text, maxsplit=1)
        if len(parts) != 2:
            continue
        role = _compact(parts[0])
        if not any(keyword in role for keyword in role_keywords):
            continue
        name = _compact(parts[1])
        if not name:
            continue
        yield paragraph, role, name


# ----------------------------------------------------------------------
# 元数据抽取
# ----------------------------------------------------------------------
class CoverMetadataExtractor:
    """按排版规范从封面段落中抽取元数据。

    只取标记（设计人员名单）之前的段落：其后是目录与正文，
    把它们纳入识别会误抓正文里的项目名。
    """

    def __init__(self, pkg: DocxPackage) -> None:
        self.pkg = pkg
        config = cover_config()
        self.marker = config["designers_marker"]
        rules = config["metadata"]
        self.company_suffix = re.compile(rules["company_suffix_pattern"])
        self.project_suffix = rules["project_name_suffix"]
        self.library_pattern = re.compile(rules["library_number_pattern"])
        self.date_day = re.compile(rules["date_year_month_day_pattern"])
        self.date_month = re.compile(rules["date_year_month_pattern"])
        self.role_keywords = tuple(rules["personnel_role_keywords"])
        self.skip_keywords = tuple(rules["design_filter_keywords"])
        self.merge_company_prefix = bool(rules["merge_company_prefix"])

    # ------------------------------------------------------------------
    def extract(self) -> CoverMetadata:
        texts = self._cover_paragraph_texts()
        meta = CoverMetadata()

        for text in texts:
            if not text:
                continue
            if any(keyword in text for keyword in self.skip_keywords):
                continue

            if not meta.library_number:
                match = self.library_pattern.search(text)
                if match:
                    meta.library_number = match.group(1).strip()

            if not meta.date_text:
                meta.date_text = self._pick_date(text)

            if self._is_personnel(text):
                meta.personnel.append(self._split_personnel(text))
                continue

            if not meta.project_name and self._is_project_name(text):
                meta.project_name = text

        if meta.project_name and self.merge_company_prefix:
            meta.project_name = self._prepend_company(meta.project_name, texts)

        return meta

    # ------------------------------------------------------------------
    def _cover_paragraph_texts(self) -> list[str]:
        """取标记段落之前的所有非空段落文本（含表格内段落）。"""
        texts: list[str] = []
        for element in self.pkg.body:
            if element.tag == qn("w:p"):
                text = paragraph_text(element)
                if self.marker in _compact(text):
                    return texts
                if text:
                    texts.append(text)
            elif element.tag == qn("w:tbl"):
                for paragraph in element.iter(qn("w:p")):
                    text = paragraph_text(paragraph)
                    if text:
                        texts.append(text)
        return texts

    # ------------------------------------------------------------------
    def _pick_date(self, text: str) -> str:
        """日期优先取"年月日"，其次"年月"。"""
        normalized = _compact(text)
        match = self.date_day.search(normalized)
        if match:
            return match.group(1)
        match = self.date_month.search(normalized)
        if match:
            return match.group(1)
        return ""

    def _is_personnel(self, text: str) -> bool:
        """人员行形如"职务：姓名"，职务（去空格）需命中关键词。

        注意：原稿的职务可能被拆成 "总 设 计 师"，必须先压缩空格再匹配。
        """
        parts = re.split(r"[:：]", text, maxsplit=1)
        if len(parts) != 2:
            return False
        role = _compact(parts[0])
        return any(keyword in role for keyword in self.role_keywords)

    @staticmethod
    def _split_personnel(text: str) -> tuple[str, str]:
        parts = re.split(r"[:：]", text, maxsplit=1)
        role = _compact(parts[0])
        name = _compact(parts[1]) if len(parts) > 1 else ""
        return role, name

    def _is_project_name(self, text: str) -> bool:
        """项目名：以"项目"结尾且不含日期，也不是单位名。"""
        stripped = text.strip()
        if not stripped.endswith(self.project_suffix):
            return False
        if self._pick_date(stripped):
            return False
        return not self.company_suffix.search(stripped)

    def _prepend_company(self, project_name: str, texts: list[str]) -> str:
        """封面项目名常拆成"单位 / 项目"两行，模板只有一个占位段落。

        规则显式写在配置里（merge_company_prefix），可关闭。
        """
        if not project_name:
            return project_name
        for text in texts:
            stripped = text.strip()
            if not stripped or stripped == project_name:
                continue
            if self.company_suffix.search(stripped) and not self._pick_date(stripped):
                merged = f"{stripped}\n{project_name}"
                logger.info("封面项目名按两行合并: %s / %s", stripped, project_name)
                return merged
        return project_name


# ----------------------------------------------------------------------
# 封面加工
# ----------------------------------------------------------------------
class CoverProcessor:
    """在唯一封面模板上做回填与设计人员表移植。"""

    def __init__(self, pkg: DocxPackage) -> None:
        self.pkg = pkg
        config = cover_config()
        self.marker = config["designers_marker"]
        self.stage_primary = tuple(config["stage_primary_variants"])
        self.stage_secondary = tuple(config["stage_secondary_variants"])
        self.role_keywords = tuple(config["metadata"]["personnel_role_keywords"])
        fmt = config["format"]
        self.title_keywords = tuple(fmt["title_spacing_keywords"])
        self.character_spacing_pt = float(fmt["character_spacing_pt"])

    # ------------------------------------------------------------------
    def apply_stage(self, primary: str, secondary: str) -> dict[str, int]:
        """设置阶段文字（初步设计 / 可行性研究 及可选的 （代可研）/（代初设））。

        阶段主/副文字在模板里是独立段落；副文字为空（文档类型 1/2）时整段删除。
        """
        replaced = 0
        removed = 0
        for paragraph in list(self.pkg.paragraphs()):
            text = _compact(paragraph_text(paragraph))
            if text in self.stage_primary:
                self._set_paragraph_text(paragraph, primary)
                replaced += 1
            elif text in self.stage_secondary:
                if secondary:
                    self._set_paragraph_text(paragraph, secondary)
                else:
                    parent = paragraph.getparent()
                    if parent is not None:
                        parent.remove(paragraph)
                        removed += 1
        return {"stage_primary": replaced, "stage_secondary_removed": removed}

    def replace_fields(self, mapping: dict[str, str]) -> int:
        """回填项目名/日期/库号等字段，保留 run 字体格式。

        占位符常被 Word 拆在多个 run 里，因此必须做跨 run 替换。
        """
        valid = {key: value for key, value in mapping.items() if key and value}
        if not valid:
            return 0

        replaced = 0
        for paragraph in self._all_paragraphs():
            replaced += self._replace_in_paragraph(paragraph, valid)
        return replaced

    def replace_personnel(self, personnel: list[tuple[str, str]]) -> int:
        """按顺序回填人员（职务：姓名）。

        模板里的人员行是占位，原稿里的人员才是真值；双方按顺序一一对应，
        职务以原稿为准（例如模板写"副总工程师"，原稿是"事业部总经理"）。
        """
        slots = list(iter_personnel_lines(self.pkg.paragraphs(), self.role_keywords))
        replaced = 0
        for index, (paragraph, _role, _name) in enumerate(slots):
            if index < len(personnel):
                role, name = personnel[index]
                self._set_paragraph_text(paragraph, f"{role}：{name}")
                replaced += 1
            else:
                self._set_paragraph_text(paragraph, "")
        return replaced

    # ------------------------------------------------------------------
    def _all_paragraphs(self) -> list[etree._Element]:
        return list(self.pkg.body.iter(qn("w:p")))

    def _replace_in_paragraph(self, paragraph: etree._Element, mapping: dict[str, str]) -> int:
        nodes = self._text_nodes(paragraph)
        if not nodes:
            return 0

        full = "".join(node.text or "" for node in nodes)
        updated = full
        for key, value in mapping.items():
            if key in updated:
                updated = updated.replace(key, value)

        if updated == full:
            return 0

        if "\n" in updated:
            self._write_lines(paragraph, nodes, updated.split("\n"))
        else:
            nodes[0].text = updated
            for node in nodes[1:]:
                node.text = ""
        return 1

    @staticmethod
    def _write_lines(paragraph: etree._Element, nodes: list[etree._Element], lines: list[str]) -> None:
        """把多行文本写进段落，行与行之间用换行符，不碰图形 run。"""
        if not nodes:
            return
        first_run = nodes[0].getparent()
        if first_run is None:
            return

        nodes[0].text = lines[0]
        nodes[0].set(qn("xml:space"), "preserve")
        for node in nodes[1:]:
            node.text = ""

        parent = first_run.getparent()
        if parent is None:
            return
        index = parent.index(first_run)
        rpr = first_run.find(qn("w:rPr"))

        for line in lines[1:]:
            break_run = etree.Element(qn("w:r"))
            if rpr is not None:
                break_run.append(etree.fromstring(etree.tostring(rpr)))
            etree.SubElement(break_run, qn("w:br"))

            text_run = etree.Element(qn("w:r"))
            if rpr is not None:
                text_run.append(etree.fromstring(etree.tostring(rpr)))
            text_node = etree.SubElement(text_run, qn("w:t"))
            text_node.set(qn("xml:space"), "preserve")
            text_node.text = line

            index += 1
            parent.insert(index, break_run)
            index += 1
            parent.insert(index, text_run)

    @staticmethod
    def _text_nodes(paragraph: etree._Element) -> list[etree._Element]:
        """只取普通 run 内的 w:t，跳过图形/文本框里的文本。

        图形内部文本参与拼接会把坐标数字带进正文（实测模板里
        会出现 "2763520344805库号：..." 这种脏数据）。
        """
        nodes: list[etree._Element] = []
        for run in paragraph.iter(qn("w:r")):
            if _inside_skipped(paragraph, run):
                continue
            for node in run.findall(qn("w:t")):
                nodes.append(node)
        return nodes

    @staticmethod
    def _set_paragraph_text(paragraph: etree._Element, text: str) -> None:
        """把整段文本替换为 text，保留首个 run 的字体格式。"""
        nodes: list[etree._Element] = []
        for run in paragraph.iter(qn("w:r")):
            for node in run.findall(qn("w:t")):
                nodes.append(node)
        if not nodes:
            run = etree.SubElement(paragraph, qn("w:r"))
            node = etree.SubElement(run, qn("w:t"))
            node.set(qn("xml:space"), "preserve")
            nodes = [node]
        nodes[0].text = text
        for node in nodes[1:]:
            node.text = ""

    # ------------------------------------------------------------------
    def transplant_table_after_marker(self, source: DocxPackage) -> bool:
        """把源文档标记段落后的第一个表格，替换掉封面模板里的同名表格。

        设计人员表是逐项目变化的，必须保留原稿数据；模板里只是占位。
        """
        source_table = _first_table_after_marker(source, self.marker)
        if source_table is None:
            logger.warning("源文档未找到「%s」后的表格，跳过封面表格移植", self.marker)
            return False

        target_table = _first_table_after_marker(self.pkg, self.marker)
        if target_table is None:
            logger.warning("封面模板未找到「%s」后的表格，跳过封面表格移植", self.marker)
            return False

        imported = etree.fromstring(etree.tostring(source_table))

        remapper = RelationshipRemapper(self.pkg, source, only_types=("image",))
        mapping = remapper.remap()
        rewrite_relationship_ids([imported], mapping)

        parent = target_table.getparent()
        parent.replace(target_table, imported)
        logger.info("封面设计人员表已从源文档移植 (%d 行)", len(imported.findall(qn("w:tr"))))
        return True

    # ------------------------------------------------------------------
    def apply_title_spacing(self) -> int:
        """给"初步设计/可行性研究"标题加字间距。

        模板里的大字号标题需要字间距才有版式效果；同时清掉落单的行距，
        否则大字号会被裁切。原实现用 python-docx 遍历所有段落，
        这里同样覆盖表格与页眉页脚。
        """
        spacing = str(int(self.character_spacing_pt * 20))
        touched = 0

        for paragraph in self._all_paragraphs():
            text = paragraph_text(paragraph)
            if not text or not any(keyword in text for keyword in self.title_keywords):
                continue

            ppr = paragraph.find(qn("w:pPr"))
            if ppr is not None:
                spacing_el = ppr.find(qn("w:spacing"))
                if spacing_el is None:
                    spacing_el = etree.SubElement(ppr, qn("w:spacing"))
                spacing_el.set(qn("w:before"), "0")
                spacing_el.set(qn("w:after"), "0")
                spacing_el.set(qn("w:line"), "240")
                spacing_el.set(qn("w:lineRule"), "auto")

            for run in paragraph.iter(qn("w:r")):
                run_text = "".join(node.text or "" for node in run.findall(qn("w:t")))
                if len(text) >= 20 and not any(k in run_text for k in self.title_keywords):
                    continue
                rpr = run.find(qn("w:rPr"))
                if rpr is None:
                    rpr = etree.Element(qn("w:rPr"))
                    run.insert(0, rpr)
                spacing_el = rpr.find(qn("w:spacing"))
                if spacing_el is None:
                    spacing_el = etree.SubElement(rpr, qn("w:spacing"))
                spacing_el.set(qn("w:val"), spacing)
            touched += 1

        return touched

    # ------------------------------------------------------------------
    def embed_images(self) -> int:
        """封面图片改为嵌入型（wp:inline），不再使用浮动锚定。"""
        converted = 0
        for drawing in self.pkg.body.iter(qn("w:drawing")):
            if drawing.find(f".//{qn('a:blip')}") is None:
                continue
            anchor = drawing.find(qn("wp:anchor"))
            if anchor is None:
                continue
            inline = etree.Element(qn("wp:inline"))
            for tag in ("wp:extent", "wp:effectExtent", "wp:docPr", "wp:cNvGraphicFramePr", "a:graphic"):
                node = anchor.find(qn(tag))
                if node is not None:
                    inline.append(node)
            parent = anchor.getparent()
            if parent is not None:
                parent.replace(anchor, inline)
                converted += 1
        if converted:
            logger.info("封面图片已改为嵌入型: %d 处", converted)
        return converted

    def format_library_number(self) -> bool:
        """「库号：」黑体三号，编号 Times New Roman 三号。不改动同段里的图形。"""
        spec = cover_config()["format"]["library_number"]
        size = str(pt_to_half_points(float(spec["size_pt"])))

        for paragraph in self._all_paragraphs():
            nodes = self._text_nodes(paragraph)
            if not nodes:
                continue
            full = "".join(node.text or "" for node in nodes)
            match = re.search(r"(库号[:：])\s*(\S+)", full)
            if not match:
                continue

            label, value = match.group(1), match.group(2)
            first_run = nodes[0].getparent()
            if first_run is None or first_run.getparent() is None:
                continue

            parent = first_run.getparent()
            index = parent.index(first_run)
            for node in nodes:
                node.text = ""

            label_run = _font_run(label, spec["label_east_asia"], spec["label_ascii"], size, bold=True)
            value_run = _font_run(value, spec["value_east_asia"], spec["value_ascii"], size, bold=True)
            parent.insert(index + 1, label_run)
            parent.insert(index + 2, value_run)
            logger.info("库号已按黑体 / Times New Roman 分开: %s%s", label, value)
            return True

        return False

    def body_text(self) -> str:
        return _compact(element_text(self.pkg.body))


def _font_run(text: str, east_asia: str, ascii_font: str, size: str, *, bold: bool) -> etree._Element:
    run = etree.Element(qn("w:r"))
    rpr = etree.SubElement(run, qn("w:rPr"))
    fonts = etree.SubElement(rpr, qn("w:rFonts"))
    fonts.set(w_attr("ascii"), ascii_font)
    fonts.set(w_attr("hAnsi"), ascii_font)
    fonts.set(w_attr("eastAsia"), east_asia)
    for tag in ("w:sz", "w:szCs"):
        node = etree.SubElement(rpr, qn(tag))
        node.set(w_attr("val"), size)
    if bold:
        etree.SubElement(rpr, qn("w:b"))
        etree.SubElement(rpr, qn("w:bCs"))
    text_node = etree.SubElement(run, qn("w:t"))
    text_node.set(qn("xml:space"), "preserve")
    text_node.text = text
    return run


def _inside_skipped(root: etree._Element, node: etree._Element) -> bool:
    """判断节点是否位于图形/文本框子树内。"""
    skipped = {qn(tag) for tag in _SKIP_TAGS}
    current = node.getparent()
    while current is not None and current is not root:
        if current.tag in skipped:
            return True
        current = current.getparent()
    return False


def _first_table_after_marker(pkg: DocxPackage, marker: str) -> etree._Element | None:
    """在 body 顶层找到含 marker 的段落之后的第一张表格。"""
    normalized_marker = _compact(marker)
    found = False
    for element in pkg.body:
        if element.tag == qn("w:p"):
            if found:
                continue
            text = _compact(paragraph_text(element))
            if normalized_marker in text:
                found = True
        elif element.tag == qn("w:tbl") and found:
            return element
    return None
