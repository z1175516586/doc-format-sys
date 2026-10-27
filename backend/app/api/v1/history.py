"""转换历史接口。"""

from __future__ import annotations

from typing import Any

from fastapi import APIRouter

from app.api.deps import get_history_service

router = APIRouter()


@router.get("", summary="历史记录列表")
def list_history(limit: int = 20, offset: int = 0) -> dict[str, Any]:
    service = get_history_service()
    return {
        "total": service.count(),
        "items": service.list(limit=limit, offset=offset),
    }


@router.get("/{record_id}", summary="历史记录详情")
def get_history(record_id: str) -> dict[str, Any]:
    return get_history_service().get(record_id)


@router.get("/by-hash/{source_hash}", summary="按源文件哈希查同源记录")
def find_by_hash(source_hash: str) -> dict[str, Any]:
    records = get_history_service().find_by_hash(source_hash)
    return {"total": len(records), "items": records}
