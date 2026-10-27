"""任务相关 Schema。"""

from __future__ import annotations

from enum import Enum
from typing import Any

from pydantic import BaseModel, Field

from app.schemas.common import ProgressEvent
from app.schemas.document import StepReport


class TaskState(str, Enum):
    """任务状态机：PENDING -> RUNNING -> SUCCEEDED / FAILED / CANCELLED。"""

    PENDING = "pending"
    RUNNING = "running"
    SUCCEEDED = "succeeded"
    FAILED = "failed"
    CANCELLED = "cancelled"

    @property
    def is_terminal(self) -> bool:
        return self in (TaskState.SUCCEEDED, TaskState.FAILED, TaskState.CANCELLED)


class TaskStatus(BaseModel):
    """任务当前进度快照。"""

    task_id: str
    state: TaskState
    percent: int = 0
    current_step: str | None = None
    step_index: int = 0
    step_total: int = 0
    message: str = ""


class TaskSummary(BaseModel):
    """任务列表项。"""

    task_id: str
    state: TaskState
    percent: int = 0
    source_name: str = ""
    output_name: str | None = None
    created_at: str = ""
    finished_at: str | None = None


class TaskDetail(TaskSummary):
    """任务详情。"""

    steps: list[StepReport] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)
    engine: str = "none"
    error: dict[str, Any] | None = None
    events: list[ProgressEvent] = Field(default_factory=list)
    download_url: str | None = None
