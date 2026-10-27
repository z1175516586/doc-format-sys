"""健康检查。"""

from __future__ import annotations

from fastapi import APIRouter

from app.converters.registry import get_converter
from app.core.config import get_settings

router = APIRouter()


@router.get("/health", summary="健康检查")
def health() -> dict:
    """返回服务状态与依赖概况，供探针与前端启动自检使用。"""
    settings = get_settings()
    capabilities = get_converter().capabilities()
    return {
        "status": "ok",
        "app": settings.app_name,
        "version": settings.app_version,
        "engine": capabilities.to_dict(),
    }
