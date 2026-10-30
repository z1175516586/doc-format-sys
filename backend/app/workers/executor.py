"""后台任务执行器。

文档排版是 CPU 密集、分钟级任务，必须与 Web 进程的事件循环隔离：
- 用线程池执行（python-docx / lxml 大量时间在 C 层，GIL 影响可控）
- 并发数由配置控制，默认 2，避免 LibreOffice 实例互相争抢内存
- 失败必须落库到任务状态，不允许静默丢弃异常
"""

from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from app.core.config import get_settings
from app.core.errors import AppError
from app.core.logging import get_logger
from app.schemas.document import FormatRequest
from app.services.format_service import PIPELINE_STEPS, DocumentFormatService
from app.services.history_service import HistoryService
from app.services.task_service import TaskService, task_service

logger = get_logger(__name__)


class TaskExecutor:
    """任务提交与执行。"""

    def __init__(
        self,
        tasks: TaskService | None = None,
        history: HistoryService | None = None,
        max_workers: int | None = None,
    ) -> None:
        settings = get_settings()
        self.tasks = tasks or task_service
        self.history = history or HistoryService()
        self._pool = ThreadPoolExecutor(
            max_workers=max_workers or settings.worker_max_workers,
            thread_name_prefix="dfs-worker",
        )

    # ------------------------------------------------------------------
    def submit(
        self,
        request: FormatRequest,
        task_id: str,
        source_hash: str,
    ) -> None:
        self._pool.submit(self._run, request, task_id, source_hash)

    # ------------------------------------------------------------------
    def _run(self, request: FormatRequest, task_id: str, source_hash: str) -> None:
        service = DocumentFormatService()
        self.tasks.mark_running(task_id, len(PIPELINE_STEPS))

        def on_progress(step: str, index: int, total: int, message: str) -> None:
            self.tasks.report_progress(task_id, step, index, total, message)

        try:
            result = service.format(request, on_progress=on_progress)
        except AppError as exc:
            issues = getattr(exc, "issues", None)
            self._record_failure(task_id, source_hash, request, exc.code, exc.message, issues)
            self.tasks.mark_failed(task_id, exc.code, exc.message, issues=issues)
            return
        except Exception as exc:
            logger.exception("任务 %s 未预期异常", task_id)
            self._record_failure(task_id, source_hash, request, "internal_error", str(exc))
            self.tasks.mark_failed(task_id, "internal_error", f"任务异常: {exc}")
            return

        task = self.tasks.get(task_id)
        task.steps = result.steps
        task.warnings = result.warnings
        task.engine = result.engine

        output_path = Path(result.output_path) if result.output_path else None

        self.history.record_format_result(
            task_id=task_id,
            source_name=request.source_name or Path(request.source_path).name,
            source_hash=source_hash,
            result=result,
            parent_id=request.source_node_id,
        )

        if result.success and output_path is not None:
            self.tasks.mark_succeeded(task_id, str(output_path), output_path.name)
        else:
            self.tasks.mark_failed(
                task_id, "layout_failed", result.message, issues=result.issues
            )

    def _record_failure(
        self,
        task_id: str,
        source_hash: str,
        request: FormatRequest,
        code: str,
        message: str,
        issues: list[dict[str, object]] | None = None,
    ) -> None:
        try:
            warnings = [code]
            for item in issues or []:
                text = str(item.get("message") or "")
                if text:
                    warnings.append(text)
            source_name = request.source_name or Path(request.source_path).name
            if request.source_node_id:
                self.history.register_failure(
                    parent_id=request.source_node_id,
                    task_id=task_id,
                    source_name=source_name,
                    source_hash=source_hash,
                    message=message,
                    warnings=warnings,
                )
            else:
                self.history.record(
                    task_id=task_id,
                    source_name=source_name,
                    source_hash=source_hash,
                    output_name=None,
                    output_path=None,
                    success=False,
                    message=message,
                    engine="unknown",
                    duration_ms=0,
                    warnings=warnings,
                )
        except Exception:
            logger.exception("写历史失败")

    # ------------------------------------------------------------------
    def shutdown(self) -> None:
        self._pool.shutdown(wait=False, cancel_futures=True)


_executor: TaskExecutor | None = None


def get_executor() -> TaskExecutor:
    """惰性获取执行器单例。

    不在模块导入时创建：应用停止后需要能被重建（测试、热重载场景）。
    """
    global _executor
    if _executor is None:
        _executor = TaskExecutor()
    return _executor


def shutdown_executor() -> None:
    global _executor
    if _executor is not None:
        _executor.shutdown()
        _executor = None
