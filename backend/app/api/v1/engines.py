"""排版引擎能力查询。

前端用它决定是否展示"自动更新目录页码"相关的提示，
避免后端不支持却让用户以为已经生效。
"""

from __future__ import annotations

from fastapi import APIRouter

from app.converters.registry import get_converter

router = APIRouter()


@router.get("", summary="查询当前排版引擎能力")
def engine_capabilities() -> dict:
    return get_converter().capabilities().to_dict()
