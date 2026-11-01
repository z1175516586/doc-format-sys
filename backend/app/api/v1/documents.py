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
    update_fields_on_open: bool | None = None,
    server_refresh_fields: bool | None = None,
) -> ConvertOptions:
    """把表单字段组装为选项对象，默认值统一来自配置文件。"""
    if update_fields_on_open and server_refresh_fields:
        # 二选一：两种方式是同一件事的两种实现，同时开启只是白占全局引擎
        raise ValidationError(
            "update_fields_on_open 与 server_refresh_fields 只能选一个",
            detail={"update_fields_on_open": True, "server_refresh_fields": True},
        )

    open_on_open, server_refresh = _resolve_page_source(
        update_fields_on_open, server_refresh_fields
    )

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
        update_fields_on_open=open_on_open,
        server_refresh_fields=server_refresh,
    )


def _resolve_page_source(
    update_fields_on_open: bool | None,
    server_refresh_fields: bool | None,
) -> tuple[bool, bool]:
    """把两个可缺省开关收敛成一对互斥值。

    - 都不传：取配置默认（默认是客户端更新）
    - 只传一个：另一个必为 False，因为「选了这种方式」本身就排除了另一种

    只传 ``server_refresh_fields=True`` 而让另一个回落到配置默认 ``True`` 会
    造出矛盾组合，这里必须显式收敛，不能交给默认值碰运气。
    """
    if update_fields_on_open is None and server_refresh_fields is None:
        toc_rules = layout_rules()["toc"]
        return bool(toc_rules["update_fields_on_open"]), bool(toc_rules["server_refresh_fields"])
    if server_refresh_fields is not None:
        return update_fields_on_open is True, server_refresh_fields
    return bool(update_fields_on_open), False


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
        source_name=file.filename or source_path.name,
        options=options,
    )


def _attach_source_node(request: FormatRequest) -> FormatRequest:
    """转换前把源文档挂到转换树上。"""
    history = get_history_service()
    source_path = Path(request.source_path)
    ensured = history.ensure_source(source_path, request.source_name or source_path.name)
    request.source_node_id = ensured["node"]["node_id"]
    return request


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
    update_fields_on_open: bool | None = Form(
        None, description="打开文档时让 Word/WPS 自动更新域（页码由客户端算）"
    ),
    server_refresh_fields: bool | None = Form(
        None, description="服务器端用 LibreOffice 预刷新目录页码（内部加锁串行）"
    ),
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
        update_fields_on_open,
        server_refresh_fields,
    )
    request = await _prepare_request(file, template_file, template_id, options)
    request = _attach_source_node(request)

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
    update_fields_on_open: bool | None = Form(None),
    server_refresh_fields: bool | None = Form(None),
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
        update_fields_on_open,
        server_refresh_fields,
    )
    request = await _prepare_request(file, template_file, template_id, options)
    request = _attach_source_node(request)

    history = get_history_service()
    source_hash = history.file_hash(Path(request.source_path))

    service = DocumentFormatService()
    result = await run_in_threadpool(service.format, request)

    history.record_format_result(
        task_id=uuid.uuid4().hex,
        source_name=request.source_name or file.filename or Path(request.source_path).name,
        source_hash=source_hash,
        result=result,
        parent_id=request.source_node_id,
    )

    if not result.success or not result.output_path:
        raise ValidationError(result.message, detail=result.issues or None)

    output = Path(result.output_path)
    if not output.exists():
        raise ValidationError("排版结果文件丢失")

    return FileResponse(
        output,
        filename=output.name,
        media_type=_DOCX_MEDIA_TYPE,
    )
