"""任务查询、进度订阅与结果下载。"""

from __future__ import annotations

import json
from pathlib import Path

from fastapi import APIRouter
from fastapi.responses import FileResponse, StreamingResponse

from app.core.errors import NotFoundError, ValidationError
from app.domain.ooxml.document_comparison import compare_documents
from app.schemas.comparison import DocumentComparison
from app.schemas.task import TaskDetail, TaskState, TaskSummary
from app.services.task_service import task_service

router = APIRouter()

_DOCX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


@router.get("", summary="任务列表", response_model=list[TaskSummary])
def list_tasks(limit: int = 20) -> list[TaskSummary]:
    return task_service.list(limit=limit)


@router.get("/{task_id}", summary="任务详情", response_model=TaskDetail)
def get_task(task_id: str) -> TaskDetail:
    return task_service.get(task_id).detail()


@router.get("/{task_id}/events", summary="订阅任务进度（SSE）")
async def stream_events(task_id: str) -> StreamingResponse:
    """Server-Sent Events 进度流。

    选择 SSE 而非 WebSocket：进度是单向流，SSE 走普通 HTTP，
    天然兼容 Nginx 反代与断线重连，运维成本更低。
    """
    task_service.get(task_id)  # 提前校验存在性，404 直接返回

    async def event_source():
        async for event in task_service.subscribe(task_id):
            payload = json.dumps(event.model_dump(), ensure_ascii=False)
            yield f"data: {payload}\n\n"

    return StreamingResponse(
        event_source(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            # 关闭 Nginx 缓冲，否则进度会被攒着一起发
            "X-Accel-Buffering": "no",
        },
    )


@router.get("/{task_id}/download", summary="下载排版结果", response_class=FileResponse)
def download(task_id: str) -> FileResponse:
    task = task_service.get(task_id)
    if not task.output_path:
        raise NotFoundError("任务尚未产出结果文件")

    path = Path(task.output_path)
    if not path.exists():
        raise NotFoundError("结果文件已被清理")

    return FileResponse(
        path,
        filename=task.output_name or path.name,
        media_type=_DOCX_MEDIA_TYPE,
    )


@router.get("/{task_id}/comparison", summary="对比源文档与排版结果")
def comparison(task_id: str) -> DocumentComparison:
    task = task_service.get(task_id)
    if task.state is not TaskState.SUCCEEDED or not task.output_path:
        raise ValidationError("转换成功后可查看文档对比")
    if not task.source_path:
        raise NotFoundError("原始文档不可用，请重新转换")
    source = Path(task.source_path)
    output = Path(task.output_path)
    # .doc 任务打开时已另存为同目录同名的 .docx，作为实际排版前的输入。
    if source.suffix.lower() == ".doc":
        source = source.with_suffix(".docx")
    if not source.is_file() or not output.is_file():
        raise NotFoundError("原始文档或转换结果已被清理，请重新转换")
    return compare_documents(source, output, task.source_name, task.output_name or output.name)
