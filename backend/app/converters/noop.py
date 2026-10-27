"""纯文本/无引擎实现。

作用：在没有安装 LibreOffice 的环境（如 CI、开发机）也能跑通完整业务流程。
所有需要真实排版计算的能力一律如实声明为不支持，由上层决定如何提示。
"""

from __future__ import annotations

from pathlib import Path

from app.converters.base import ConverterCapabilities, DocConverter
from app.core.errors import EngineUnavailableError
from app.core.logging import get_logger

logger = get_logger(__name__)


class NoopConverter(DocConverter):
    """不做任何外部调用的引擎占位实现。"""

    def capabilities(self) -> ConverterCapabilities:
        return ConverterCapabilities(
            name="none",
            available=False,
            notes=(
                "未检测到可用排版引擎",
                "目录页码与域结果不会被刷新，需在 Word/WPS 中按 F9 更新",
                "不支持 .doc 转换与 PDF 导出",
            ),
        )

    def convert(self, src: Path, dst: Path, target_format: str) -> Path:
        raise EngineUnavailableError("当前环境未配置排版引擎，无法执行格式转换")

    def refresh_fields(self, path: Path) -> bool:
        logger.info("无可用引擎，跳过域刷新: %s", path.name)
        return False

    def page_count(self, path: Path) -> int | None:
        return None

    def export_pdf(self, src: Path, dst: Path) -> Path | None:
        return None
