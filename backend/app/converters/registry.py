"""引擎注册与选择。"""

from __future__ import annotations

from functools import lru_cache

from app.converters.base import DocConverter
from app.converters.libreoffice import LibreOfficeConverter
from app.converters.noop import NoopConverter
from app.core.config import get_settings
from app.core.logging import get_logger

logger = get_logger(__name__)


@lru_cache
def get_converter() -> DocConverter:
    """按配置选择排版引擎。

    layout_engine 取值：
    - auto：优先 LibreOffice，不可用则降级为 Noop
    - libreoffice：强制 LibreOffice，不可用直接启动失败（生产环境推荐）
    - none：始终 Noop（用于测试与纯 XML 场景）
    """
    settings = get_settings()
    mode = (settings.layout_engine or "auto").lower()

    if mode == "none":
        return NoopConverter()

    if mode in ("auto", "libreoffice"):
        engine = LibreOfficeConverter()
        if engine.capabilities().available:
            return engine
        if mode == "libreoffice":
            raise RuntimeError(
                "配置要求使用 LibreOffice，但未找到 soffice。"
                "请安装 LibreOffice 或设置 DFS_LIBREOFFICE_BINARY。"
            )
        logger.warning("未找到 LibreOffice，降级为无引擎模式（目录页码不会刷新）")
        return NoopConverter()

    raise RuntimeError(f"未知的 layout_engine 配置: {mode}")


def reset_converter_cache() -> None:
    """测试用：清空引擎缓存。"""
    get_converter.cache_clear()
