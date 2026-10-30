"""转换历史接口。"""

from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Any

from fastapi import APIRouter, File, UploadFile
from fastapi.responses import FileResponse

from app.api.deps import get_history_service
from app.core.constants import lineage_config
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


@router.post("/inspect", summary="根据上传文件识别转换树")
async def inspect_lineage(file: UploadFile = File(..., description="待识别的文档")) -> dict[str, Any]:
    """上传任意一份文档，返回它在系统里的来时之路和整棵转换树。"""
    suffix = Path(file.filename or "upload.docx").suffix or ".docx"
    payload = await file.read()
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(payload)
        tmp_path = Path(tmp.name)
    try:
        return get_history_service().inspect(tmp_path, file.filename or tmp_path.name)
    finally:
        tmp_path.unlink(missing_ok=True)


@router.get("/nodes/{node_id}/file", summary="下载转换树中的文档节点", response_class=FileResponse)
def download_node_file(node_id: str) -> FileResponse:
    path, name = get_history_service().node_file(node_id)
    return FileResponse(path, filename=name, media_type=_DOCX_MEDIA_TYPE)


@router.get("/by-hash/{source_hash}", summary="按源文件哈希查同源记录")
def find_by_hash(source_hash: str) -> dict[str, Any]:
    records = get_history_service().find_by_hash(source_hash)
    return {"total": len(records), "items": records}


@router.get("/{record_id}", summary="历史记录详情")
def get_history(record_id: str) -> dict[str, Any]:
    return get_history_service().get(record_id)


@router.get("/{record_id}/file", summary="下载历史产出文件", response_class=FileResponse)
def download_history_file(record_id: str) -> FileResponse:
    """按历史记录下载当时的排版结果。

    优先走节点保存的副本。记录没有文件或文件被手工删除时给出明确错误。
    """
    record = get_history_service().get(record_id)
    child_id = record.get("child_id")
    if child_id:
        path, name = get_history_service().node_file(str(child_id))
        return FileResponse(path, filename=name, media_type=_DOCX_MEDIA_TYPE)

    raw = record.get("output_path")
    if not raw:
        raise NotFoundError("该记录没有可下载的结果文件")

    path = Path(str(raw))
    if not path.is_file():
        raise NotFoundError(str(lineage_config()["missing_file_message"]))

    name = record.get("output_name") or path.name
    return FileResponse(path, filename=str(name), media_type=_DOCX_MEDIA_TYPE)
