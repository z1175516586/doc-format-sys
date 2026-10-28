"""转换历史接口。"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from fastapi import APIRouter
from fastapi.responses import FileResponse

from app.api.deps import get_history_service
from app.core.errors import NotFoundError

router = APIRouter()

_DOCX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


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


@router.get("/{record_id}/file", summary="下载历史产出文件", response_class=FileResponse)
def download_history_file(record_id: str) -> FileResponse:
    """按历史记录下载当时的排版结果。

    旧前端用 /result/{node_code} 下载历史文件。任务下载只在任务存活期间有效，
    历史下载走记录里保存的 output_path。
    """
    record = get_history_service().get(record_id)
    raw = record.get("output_path")
    if not raw:
        raise NotFoundError("该记录没有可下载的结果文件")

    path = Path(str(raw))
    if not path.is_file():
        raise NotFoundError("结果文件已被清理")

    name = record.get("output_name") or path.name
    return FileResponse(path, filename=str(name), media_type=_DOCX_MEDIA_TYPE)
