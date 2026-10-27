"""文档排版接口。"""

from __future__ import annotations

import asyncio
import uuid
from pathlib import Path

from fastapi import APIRouter, File, Form, UploadFile
from fastapi.concurrency import run_in_threadpool
from fastapi.responses import FileResponse

from app.api.deps import get_history_service, get_template_service
from app.core.constants import layout_rules
from app.core.errors import NotFoundError, ValidationError
from app.core.logging import get_logger
from app.schemas.document import ConvertOptions, FormatRequest
from app.services.format_service import DocumentFormatService, build_output_path
from app.services.upload_service import UploadService
from app.workers.executor import get_executor

logger = get_logger(__name__)
router = APIRouter()

_DOCX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


def _build_options(
    toc_title: str | None,
    header_text: str,
    header_left_text: str,
    header_right_text: str,
    document_type: int,
    insert_toc: bool,
    remove_highlights: bool,
    convert_numbering: bool,
    apply_template: bool,
) -> ConvertOptions:
    """把表单字段组装为选项对象，默认值统一来自配置文件。"""
    default_title = layout_rules()["toc"]["default_title"]
    return ConvertOptions(
        toc_title=toc_title or default_title,
        header_text=header_text,
        header_left_text=header_left_text,
        header_right_text=header_right_text,
        document_type=document_type,
        insert_toc=insert_toc,
        remove_highlights=remove_highlights,
        convert_numbering=convert_numbering,
        apply_template=apply_template,
    )


async def _prepare_request(
    file: UploadFile,
    template_file: UploadFile | None,
    template_id: str | None,
    options: ConvertOptions,
) -> FormatRequest:
    uploads = UploadService()
    source_path = await uploads.save(file, prefix="src")

    template_path: Path | None = None
    if template_file is not None and template_file.filename:
        template_path = await uploads.save(template_file, prefix="tpl")
    elif template_id:
        try:
            template_path = get_template_service().get_path(template_id)
        except NotFoundError as exc:
            source_path.unlink(missing_ok=True)
            raise ValidationError(str(exc)) from exc

    output_path = build_output_path(file.filename or source_path.name)

    return FormatRequest(
        source_path=str(source_path),
        output_path=str(output_path),
        template_path=str(template_path) if template_path else None,
        options=options,
    )


@router.post("/convert", summary="提交文档排版任务（异步）")
async def convert_async(
    file: UploadFile = File(..., description="待排版文档（docx）"),
    template_file: UploadFile | None = File(None, description="自定义模板（docx）"),
    template_id: str | None = Form(None, description="已注册模板 ID"),
    toc_title: str | None = Form(None),
    header_text: str = Form(""),
    header_left_text: str = Form(""),
    header_right_text: str = Form(""),
    document_type: int = Form(1),
    insert_toc: bool = Form(True),
    remove_highlights: bool = Form(True),
    convert_numbering: bool = Form(True),
    apply_template: bool = Form(True),
) -> dict:
    """提交任务后立即返回 task_id，进度通过 SSE 订阅。

    适合大文档：HTTP 请求不会被长时间占用，前端可实时展示步骤进度。
    """
    options = _build_options(
        toc_title,
        header_text,
        header_left_text,
        header_right_text,
        document_type,
        insert_toc,
        remove_highlights,
        convert_numbering,
        apply_template,
    )
    request = await _prepare_request(file, template_file, template_id, options)

    history = get_history_service()
    source_hash = history.file_hash(Path(request.source_path))

    task = get_executor().tasks.create(file.filename or "未命名文档")
    task.bind_loop(asyncio.get_running_loop())
    get_executor().submit(request, task.task_id, source_hash)

    logger.info("任务已提交: %s -> %s", task.task_id, request.output_path)
    return {
        "task_id": task.task_id,
        "state": task.state,
        "events_url": f"/api/v1/tasks/{task.task_id}/events",
    }


@router.post(
    "/convert-sync",
    summary="同步执行排版并直接返回文档",
    response_class=FileResponse,
)
async def convert_sync(
    file: UploadFile = File(..., description="待排版文档（docx）"),
    template_file: UploadFile | None = File(None),
    template_id: str | None = Form(None),
    toc_title: str | None = Form(None),
    header_text: str = Form(""),
    header_left_text: str = Form(""),
    header_right_text: str = Form(""),
    document_type: int = Form(1),
    insert_toc: bool = Form(True),
    remove_highlights: bool = Form(True),
    convert_numbering: bool = Form(True),
    apply_template: bool = Form(True),
) -> FileResponse:
    """同步排版，直接返回 docx 文件。

    用于联调与脚本化调用；不推送进度，请求会阻塞至完成。
    """
    options = _build_options(
        toc_title,
        header_text,
        header_left_text,
        header_right_text,
        document_type,
        insert_toc,
        remove_highlights,
        convert_numbering,
        apply_template,
    )
    request = await _prepare_request(file, template_file, template_id, options)

    history = get_history_service()
    source_hash = history.file_hash(Path(request.source_path))

    service = DocumentFormatService()
    result = await run_in_threadpool(service.format, request)

    history.record_format_result(
        task_id=uuid.uuid4().hex,
        source_name=file.filename or Path(request.source_path).name,
        source_hash=source_hash,
        result=result,
    )

    if not result.success or not result.output_path:
        raise ValidationError(result.message)

    output = Path(result.output_path)
    if not output.exists():
        raise ValidationError("排版结果文件丢失")

    return FileResponse(
        output,
        filename=output.name,
        media_type=_DOCX_MEDIA_TYPE,
    )
