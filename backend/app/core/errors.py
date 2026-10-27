"""统一异常体系。

约定：所有可预期的业务失败都必须抛出 AppError 子类，
由 API 层统一转换为标准错误响应，禁止裸抛 Exception。
"""

from __future__ import annotations


class AppError(Exception):
    """应用异常基类。"""

    code: str = "internal_error"
    http_status: int = 500
    message: str = "服务内部错误"

    def __init__(self, message: str | None = None, *, detail: object = None) -> None:
        self.message = message or self.message
        self.detail = detail
        super().__init__(self.message)

    def to_dict(self) -> dict:
        payload: dict = {"code": self.code, "message": self.message}
        if self.detail is not None:
            payload["detail"] = self.detail
        return payload


class ValidationError(AppError):
    code = "validation_error"
    http_status = 400
    message = "请求参数不合法"


class UnsupportedFormatError(ValidationError):
    code = "unsupported_format"
    message = "不支持的文件格式"


class DocumentCorruptedError(AppError):
    code = "document_corrupted"
    http_status = 422
    message = "文档结构损坏，无法解析"


class LayoutError(AppError):
    code = "layout_error"
    http_status = 422
    message = "文档排版处理失败"


class EngineUnavailableError(AppError):
    code = "engine_unavailable"
    http_status = 503
    message = "排版引擎不可用"


class NotFoundError(AppError):
    code = "not_found"
    http_status = 404
    message = "资源不存在"


class TaskNotFoundError(NotFoundError):
    code = "task_not_found"
    message = "任务不存在"


class TaskFailedError(AppError):
    code = "task_failed"
    http_status = 422
    message = "任务执行失败"
