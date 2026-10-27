"""转换历史服务。

以 JSONL 追加文件存储，不引入数据库。
理由：数据库结构调整需要单独评审（见 docs/REFACTOR_PLAN.md §8），
而历史记录是纯追加型数据，JSONL 足以支撑查询与审计，且部署零依赖。

原实现把历史写入 doc_record / doc_node 两张表，并通过修改 docx 的
settings.xml 写入 DocHistoryId 来串联版本链。这里保留"版本链"语义，
但链路信息记录在历史文件里，不再侵入文档本身——避免每次转换都改写文档。
"""

from __future__ import annotations

import hashlib
import json
import threading
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import TYPE_CHECKING, Any

from app.core.config import get_settings
from app.core.errors import NotFoundError
from app.core.logging import get_logger

if TYPE_CHECKING:
    from app.schemas.document import FormatResult

logger = get_logger(__name__)

_RECORDS_FILENAME = "records.jsonl"


class HistoryService:
    """转换历史的写入与查询。"""

    def __init__(self, root: Path | None = None) -> None:
        self.root = root or get_settings().history_dir
        self.root.mkdir(parents=True, exist_ok=True)
        self._path = self.root / _RECORDS_FILENAME
        self._lock = threading.Lock()

    # ------------------------------------------------------------------
    # 写入
    # ------------------------------------------------------------------
    def record(
        self,
        *,
        task_id: str,
        source_name: str,
        source_hash: str,
        output_name: str | None,
        output_path: str | None,
        success: bool,
        message: str,
        engine: str,
        duration_ms: int,
        steps: list[dict[str, Any]] | None = None,
        warnings: list[str] | None = None,
        parent_id: str | None = None,
    ) -> dict[str, Any]:
        entry = {
            "record_id": uuid.uuid4().hex,
            "task_id": task_id,
            "created_at": datetime.now(timezone.utc).isoformat(),
            "source_name": source_name,
            "source_hash": source_hash,
            "output_name": output_name,
            "output_path": output_path,
            "success": success,
            "message": message,
            "engine": engine,
            "duration_ms": duration_ms,
            "steps": steps or [],
            "warnings": warnings or [],
            "parent_id": parent_id,
        }

        with self._lock:
            with self._path.open("a", encoding="utf-8") as fp:
                fp.write(json.dumps(entry, ensure_ascii=False) + "\n")

        logger.info("记录历史 %s success=%s", entry["record_id"], success)
        return entry

    @staticmethod
    def file_hash(path: Path) -> str:
        """计算文件 sha256，用于识别同一源文件的重复转换。"""
        digest = hashlib.sha256()
        with Path(path).open("rb") as fp:
            for chunk in iter(lambda: fp.read(1024 * 1024), b""):
                digest.update(chunk)
        return digest.hexdigest()

    def record_format_result(
        self,
        *,
        task_id: str,
        source_name: str,
        source_hash: str,
        result: "FormatResult",
        parent_id: str | None = None,
    ) -> dict[str, Any]:
        """把一次排版结果写入历史。

        同步与异步两条路径都走这里，避免历史记录出现"只有异步有记录"的缺口。
        """
        output_path = Path(result.output_path) if result.output_path else None
        return self.record(
            task_id=task_id,
            source_name=source_name,
            source_hash=source_hash,
            output_name=output_path.name if output_path else None,
            output_path=str(output_path) if output_path else None,
            success=result.success,
            message=result.message,
            engine=result.engine,
            duration_ms=result.duration_ms,
            steps=[step.model_dump() for step in result.steps],
            warnings=result.warnings,
            parent_id=parent_id,
        )

    # ------------------------------------------------------------------
    # 查询
    # ------------------------------------------------------------------
    def list(self, limit: int = 20, offset: int = 0) -> list[dict[str, Any]]:
        records = self._read_all()
        records.reverse()
        return records[offset : offset + limit]

    def count(self) -> int:
        return len(self._read_all())

    def get(self, record_id: str) -> dict[str, Any]:
        for record in self._read_all():
            if record.get("record_id") == record_id:
                return record
        raise NotFoundError(f"历史记录不存在: {record_id}")

    def find_by_hash(self, source_hash: str) -> list[dict[str, Any]]:
        """查同源文件的全部转换记录，按时间正序。"""
        return [
            record
            for record in self._read_all()
            if record.get("source_hash") == source_hash
        ]

    def _read_all(self) -> list[dict[str, Any]]:
        if not self._path.exists():
            return []

        records: list[dict[str, Any]] = []
        with self._lock:
            with self._path.open("r", encoding="utf-8") as fp:
                for line_no, line in enumerate(fp, start=1):
                    line = line.strip()
                    if not line:
                        continue
                    try:
                        records.append(json.loads(line))
                    except json.JSONDecodeError:
                        logger.warning("历史记录第 %d 行损坏，已跳过", line_no)
        return records
