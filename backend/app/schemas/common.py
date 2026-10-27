"""通用 Schema。"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field


class ErrorResponse(BaseModel):
    """统一错误响应。"""

    code: str = Field(description="错误码，前端据此做分支处理")
    message: str = Field(description="可直接展示给用户的错误说明")
    detail: Any | None = Field(default=None, description="调试用附加信息")


class PageMeta(BaseModel):
    """分页元信息。"""

    total: int = 0
    page: int = 1
    size: int = 20


class ProgressEvent(BaseModel):
    """进度事件（SSE 数据体）。"""

    seq: int = Field(description="事件序号，用于断线续传与去重")
    type: Literal["progress", "log", "done", "error"] = "progress"
    step: str | None = None
    index: int = 0
    total: int = 0
    percent: int = 0
    message: str = ""
    data: dict[str, Any] | None = None
