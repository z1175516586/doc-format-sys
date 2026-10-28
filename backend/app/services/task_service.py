"""任务服务：状态机 + 进度事件总线。

设计要点：
- 任务状态保存在内存，附带 JSONL 落盘（见 history_service），不依赖数据库
- 进度通过 SSE 推送，单向流，无需 WebSocket 的握手与心跳
- 线程安全：worker 在独立线程运行，事件通过 loop.call_soon_threadsafe 投递到事件循环
"""

from __future__ import annotations

import asyncio
import threading
import uuid
from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Any, AsyncIterator

from app.core.errors import TaskNotFoundError
from app.core.logging import get_logger
from app.schemas.common import ProgressEvent
from app.schemas.document import StepReport
from app.schemas.task import TaskDetail, TaskState, TaskSummary

logger = get_logger(__name__)


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()


@dataclass
class Task:
    """单个排版任务。"""

    task_id: str
    source_name: str = ""
    state: TaskState = TaskState.PENDING
    percent: int = 0
    current_step: str | None = None
    step_index: int = 0
    step_total: int = 0
    message: str = ""
    output_name: str | None = None
    output_path: str | None = None
    created_at: str = field(default_factory=_now)
    finished_at: str | None = None
    steps: list[StepReport] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    engine: str = "none"
    error: dict[str, Any] | None = None
    events: list[ProgressEvent] = field(default_factory=list)
    _seq: int = 0
    _listeners: list[asyncio.Queue] = field(default_factory=list)
    _loop: asyncio.AbstractEventLoop | None = None

    # ---- 事件 ----
    def bind_loop(self, loop: asyncio.AbstractEventLoop) -> None:
        self._loop = loop

    def emit(self, event: ProgressEvent) -> None:
        """写事件并广播给所有订阅者（可从任意线程调用）。"""
        self._seq += 1
        event.seq = self._seq
        self.events.append(event)

        for queue in list(self._listeners):
            try:
                if self._loop is not None and self._loop.is_running():
                    self._loop.call_soon_threadsafe(queue.put_nowait, event)
                else:
                    queue.put_nowait(event)
            except Exception:
                logger.debug("事件投递失败（订阅者已断开）", exc_info=True)

    # ---- 快照 ----
    def summary(self) -> TaskSummary:
        return TaskSummary(
            task_id=self.task_id,
            state=self.state,
            percent=self.percent,
            source_name=self.source_name,
            output_name=self.output_name,
            created_at=self.created_at,
            finished_at=self.finished_at,
        )

    def detail(self) -> TaskDetail:
        return TaskDetail(
            task_id=self.task_id,
            state=self.state,
            percent=self.percent,
            source_name=self.source_name,
            output_name=self.output_name,
            created_at=self.created_at,
            finished_at=self.finished_at,
            steps=self.steps,
            warnings=self.warnings,
            engine=self.engine,
            error=self.error,
            events=self.events[-200:],
            download_url=(
                f"/api/v1/tasks/{self.task_id}/download"
                if self.state is TaskState.SUCCEEDED and self.output_path
                else None
            ),
        )


class TaskService:
    """任务注册表与事件分发。"""

    def __init__(self, max_tasks: int = 500) -> None:
        self._tasks: dict[str, Task] = {}
        self._lock = threading.RLock()
        self._max_tasks = max_tasks

    # ------------------------------------------------------------------
    # 创建与查询
    # ------------------------------------------------------------------
    def create(self, source_name: str) -> Task:
        task = Task(task_id=uuid.uuid4().hex, source_name=source_name)
        with self._lock:
            self._tasks[task.task_id] = task
            self._evict_if_needed()
        logger.info("创建任务 %s (%s)", task.task_id, source_name)
        return task

    def get(self, task_id: str) -> Task:
        with self._lock:
            task = self._tasks.get(task_id)
        if task is None:
            raise TaskNotFoundError(f"任务不存在: {task_id}")
        return task

    def list(self, limit: int = 20) -> list[TaskSummary]:
        with self._lock:
            tasks = sorted(self._tasks.values(), key=lambda t: t.created_at, reverse=True)
        return [task.summary() for task in tasks[:limit]]

    def _evict_if_needed(self) -> None:
        if len(self._tasks) <= self._max_tasks:
            return
        finished = [t for t in self._tasks.values() if t.state.is_terminal]
        finished.sort(key=lambda t: t.created_at)
        for task in finished[: len(self._tasks) - self._max_tasks]:
            self._tasks.pop(task.task_id, None)

    # ------------------------------------------------------------------
    # 状态更新
    # ------------------------------------------------------------------
    def mark_running(self, task_id: str, step_total: int) -> None:
        task = self.get(task_id)
        task.state = TaskState.RUNNING
        task.step_total = step_total
        task.message = "开始排版"

    def report_progress(
        self,
        task_id: str,
        step: str,
        index: int,
        total: int,
        message: str,
    ) -> None:
        task = self.get(task_id)
        task.current_step = step
        task.step_index = index
        task.step_total = total
        task.percent = int(index / total * 100) if total else 0
        task.message = message
        task.emit(
            ProgressEvent(
                seq=0,
                type="progress",
                step=step,
                index=index,
                total=total,
                percent=task.percent,
                message=message,
            )
        )

    def log(self, task_id: str, message: str) -> None:
        task = self.get(task_id)
        task.emit(ProgressEvent(seq=0, type="log", message=message))

    def mark_succeeded(self, task_id: str, output_path: str, output_name: str) -> None:
        task = self.get(task_id)
        task.state = TaskState.SUCCEEDED
        task.percent = 100
        task.output_path = output_path
        task.output_name = output_name
        task.message = "排版完成"
        task.finished_at = _now()
        task.emit(
            ProgressEvent(
                seq=0,
                type="done",
                percent=100,
                message="排版完成",
                data={
                    "download_url": f"/api/v1/tasks/{task_id}/download",
                    "output_name": output_name,
                },
            )
        )

    def mark_failed(
        self,
        task_id: str,
        code: str,
        message: str,
        *,
        issues: list[dict[str, object]] | None = None,
    ) -> None:
        task = self.get(task_id)
        task.state = TaskState.FAILED
        task.error = {"code": code, "message": message}
        if issues:
            task.error["issues"] = issues
        task.message = message
        task.finished_at = _now()
        task.emit(
            ProgressEvent(seq=0, type="error", message=message, data={"code": code})
        )

    # ------------------------------------------------------------------
    # 订阅（SSE）
    # ------------------------------------------------------------------
    async def subscribe(self, task_id: str, replay: bool = True) -> AsyncIterator[ProgressEvent]:
        """订阅任务事件流。

        先回放已有事件（断线重连不丢进度），再实时推送直至任务终态。
        """
        task = self.get(task_id)
        queue: asyncio.Queue = asyncio.Queue()

        if replay:
            for event in list(task.events):
                yield event

        if task.state.is_terminal:
            return

        task._listeners.append(queue)
        try:
            while True:
                try:
                    event = await asyncio.wait_for(queue.get(), timeout=15.0)
                except asyncio.TimeoutError:
                    # 心跳：保持连接存活，避免代理层断开
                    yield ProgressEvent(seq=0, type="log", message="")
                    if task.state.is_terminal:
                        break
                    continue

                yield event
                if event.type in ("done", "error"):
                    break
        finally:
            with self._lock:
                if queue in task._listeners:
                    task._listeners.remove(queue)


task_service = TaskService()
