"""接口契约（Pydantic Schema）。"""

from app.schemas.common import ErrorResponse, PageMeta, ProgressEvent
from app.schemas.document import ConvertOptions, FormatRequest, FormatResult, StepReport
from app.schemas.task import TaskDetail, TaskStatus, TaskSummary, TaskState
from app.schemas.template import TemplateInfo, TemplateKind

__all__ = [
    "ConvertOptions",
    "ErrorResponse",
    "FormatRequest",
    "FormatResult",
    "PageMeta",
    "ProgressEvent",
    "StepReport",
    "TaskDetail",
    "TaskState",
    "TaskStatus",
    "TaskSummary",
    "TemplateInfo",
    "TemplateKind",
]
