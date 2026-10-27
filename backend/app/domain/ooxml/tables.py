"""表格处理。

对应原 utils/handle_first_table.py、utils/table_font_setter.py、
utils/remove_tables_before_designers.py、utils/table_font_recorder.py。

原实现基于 python-docx 的 doc.tables / row.height，只能处理顶层表格且无法控制
run 级字体继承。这里统一在 OOXML 层处理，支持任意嵌套层级。
"""

from __future__ import annotations

from dataclasses import dataclass

from lxml import etree

from app.core.constants import cm_to_twips, cover_config, layout_rules, pt_to_half_points
from app.core.logging import get_logger
from app.core.namespaces import qn, w_attr
from app.domain.ooxml.package import DocxPackage, element_text

logger = get_logger(__name__)


@dataclass(frozen=True)
class TableFont:
    """表格字体定义。"""

    east_asia: str
    ascii: str
    size_pt: float

    @classmethod
    def from_config(cls, key: str) -> "TableFont":
        config = layout_rules()["fonts"][key]
        return cls(
            east_asia=config["east_asia"],
            ascii=config["ascii"],
            size_pt=float(config["size_pt"]),
        )


class TableProcessor:
    """表格结构、字体与封面清理。"""

    def __init__(self, pkg: DocxPackage) -> None:
        self.pkg = pkg

    # ------------------------------------------------------------------
    # 字体
    # ------------------------------------------------------------------
    def apply_font_to_first_table(self, font: TableFont) -> int:
        """把第一个表格内所有 run 的字体设为指定值，返回处理的 run 数量。"""
        tables = self.pkg.tables()
        if not tables:
            logger.info("文档无表格，跳过字体设置")
            return 0
        return self._apply_font(tables[0], font)

    def apply_font_to_all_tables(self, font: TableFont) -> int:
        total = 0
        for table in self.pkg.tables():
            total += self._apply_font(table, font)
        return total

    def _apply_font(self, table: etree._Element, font: TableFont) -> int:
        size = str(pt_to_half_points(font.size_pt))
        count = 0

        for run in table.iter(qn("w:r")):
            rpr = run.find(qn("w:rPr"))
            if rpr is None:
                rpr = etree.Element(qn("w:rPr"))
                run.insert(0, rpr)

            fonts = rpr.find(qn("w:rFonts"))
            if fonts is None:
                fonts = etree.SubElement(rpr, qn("w:rFonts"))
            fonts.set(w_attr("ascii"), font.ascii)
            fonts.set(w_attr("hAnsi"), font.ascii)
            fonts.set(w_attr("eastAsia"), font.east_asia)

            for tag in ("w:sz", "w:szCs"):
                size_el = rpr.find(qn(tag))
                if size_el is None:
                    size_el = etree.SubElement(rpr, qn(tag))
                size_el.set(w_attr("val"), size)
            count += 1

        return count

    # ------------------------------------------------------------------
    # 行高
    # ------------------------------------------------------------------
    def shrink_first_table_rows(self, height_cm: float | None = None) -> int | None:
        """把第一个表格所有行压到指定行高，返回原始第二行行高（twips）。

        返回 None 表示无表格或行数不足。主要用于封面：压缩表格高度避免封面溢出，
        排版末尾再调用 restore_first_table_row_height 还原。
        """
        tables = self.pkg.tables()
        if not tables:
            return None

        table = tables[0]
        rows = table.findall(qn("w:tr"))
        if len(rows) < 2:
            return None

        original = self._read_row_height(rows[1])

        target = cm_to_twips(height_cm or cover_config()["first_table_row_height_cm"])
        for row in rows:
            self._write_row_height(row, target, "atLeast")

        logger.info("封面表格行高压缩为 %d twips", target)
        return original

    def restore_first_table_row_height(self, original: int | None) -> bool:
        """还原第一个表格行高。"""
        if original is None:
            return False

        tables = self.pkg.tables()
        if not tables:
            return False

        for row in tables[0].findall(qn("w:tr")):
            self._write_row_height(row, original, "atLeast")
        return True

    @staticmethod
    def _read_row_height(row: etree._Element) -> int | None:
        tr_pr = row.find(qn("w:trPr"))
        if tr_pr is None:
            return None
        height = tr_pr.find(qn("w:trHeight"))
        if height is None:
            return None
        try:
            return int(height.get(w_attr("val")) or 0) or None
        except ValueError:
            return None

    @staticmethod
    def _write_row_height(row: etree._Element, value: int, rule: str) -> None:
        tr_pr = row.find(qn("w:trPr"))
        if tr_pr is None:
            tr_pr = etree.Element(qn("w:trPr"))
            row.insert(0, tr_pr)

        height = tr_pr.find(qn("w:trHeight"))
        if height is None:
            height = etree.SubElement(tr_pr, qn("w:trHeight"))
        height.set(w_attr("val"), str(value))
        height.set(w_attr("hRule"), rule)

    # ------------------------------------------------------------------
    # 封面前置表格
    # ------------------------------------------------------------------
    def remove_tables_before_marker(self, marker: str | None = None) -> int:
        """删除正文中位于标记文本之前的所有表格，内容转为普通段落。

        设计人员名单之前的表格是封面冗余信息，保留会污染正文结构。
        原实现只保留纯文本，这里同样保留文本以不丢信息。

        安全约束：**标记不存在时不做任何删除**。
        否则在结构异常的文档上会把正文章节内的表格全部删除。
        """
        target = marker or cover_config()["designers_marker"]
        body = self.pkg.body

        if not any(target in element_text(child) for child in body):
            logger.info("未找到标记「%s」，跳过封面前置表格清理", target)
            return 0

        pending: list[etree._Element] = []
        for child in list(body):
            if child.tag == qn("w:tbl"):
                pending.append(child)
                continue
            if target in element_text(child):
                break

        for table in pending:
            parent = table.getparent()
            if parent is None:
                continue
            index = parent.index(table)
            for offset, text in enumerate(self._table_to_lines(table)):
                parent.insert(index + offset, self._line_to_paragraph(text))
            parent.remove(table)

        if pending:
            logger.info("删除封面前置表格 %d 个", len(pending))
        return len(pending)

    @staticmethod
    def _table_to_lines(table: etree._Element) -> list[str]:
        lines: list[str] = []
        for row in table.findall(qn("w:tr")):
            cells: list[str] = []
            for cell in row.findall(qn("w:tc")):
                cell_paras = [
                    element_text(p) for p in cell.findall(qn("w:p"))
                ]
                cells.append(" ".join(filter(None, cell_paras)))
            line = "\t".join(cells).strip()
            if line:
                lines.append(line)
        return lines

    @staticmethod
    def _line_to_paragraph(text: str) -> etree._Element:
        p = etree.Element(qn("w:p"))
        run = etree.SubElement(p, qn("w:r"))
        t = etree.SubElement(run, qn("w:t"))
        t.set(qn("xml:space"), "preserve")
        t.text = text
        return p

    # ------------------------------------------------------------------
    # 统计
    # ------------------------------------------------------------------
    def count(self) -> int:
        return len(self.pkg.tables())
