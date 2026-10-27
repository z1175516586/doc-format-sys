"""转换引擎抽象层。

设计目的：把"需要外部排版引擎才能完成"的操作隔离在一个接口后面。
业务代码只依赖 DocConverter，不依赖 LibreOffice，也不依赖 COM。

这样带来两个能力：
1. Linux 上默认用 LibreOffice，Windows 上可临时挂 COM 实现，业务零改动
2. 将来若改用商业库（如 Aspose.Words），只需新增一个实现类
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from pathlib import Path


@dataclass(frozen=True)
class ConverterCapabilities:
    """引擎能力声明。

    必须如实声明，不得为了"看起来支持"而虚报。
    """

    name: str
    available: bool
    doc_to_docx: bool = False
    refresh_fields: bool = False
    export_pdf: bool = False
    page_count: bool = False
    render_page: bool = False
    notes: tuple[str, ...] = field(default_factory=tuple)

    def to_dict(self) -> dict:
        return {
            "name": self.name,
            "available": self.available,
            "doc_to_docx": self.doc_to_docx,
            "refresh_fields": self.refresh_fields,
            "export_pdf": self.export_pdf,
            "page_count": self.page_count,
            "render_page": self.render_page,
            "notes": list(self.notes),
        }


class DocConverter(ABC):
    """排版引擎接口。"""

    @abstractmethod
    def capabilities(self) -> ConverterCapabilities:
        """返回引擎能力，用于启动自检与降级决策。"""

    @abstractmethod
    def convert(self, src: Path, dst: Path, target_format: str) -> Path:
        """格式转换（如 doc -> docx, docx -> pdf）。"""

    @abstractmethod
    def refresh_fields(self, path: Path) -> bool:
        """刷新域（目录页码、交叉引用）。

        返回 False 表示引擎不支持，调用方必须据此在结果中标记"目录需手动更新"，
        不得静默当作成功。
        """

    @abstractmethod
    def page_count(self, path: Path) -> int | None:
        """返回页数，不支持时返回 None。"""

    @abstractmethod
    def export_pdf(self, src: Path, dst: Path) -> Path | None:
        """导出 PDF，不支持时返回 None。"""
