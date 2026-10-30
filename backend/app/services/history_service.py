"""转换历史与文档转换树。

以 JSONL 追加存储，不引入数据库。
每个成功文档是一个 node（写入 docx 的 node_id）；每次转换是一条边。
失败转换没有产出文件，在树上表现为无后代的叶子。

识别顺序：先读文档里的 node_id，没有再按哈希（ingest_hash / stored_hash）。
node_id 一致但哈希不一致，视为文档被修改，仍按原身份挂树。
文件永久保留，不自动清理；人工删了就标记「有记录、文件已删」。
"""

from __future__ import annotations

import hashlib
import json
import shutil
import threading
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import TYPE_CHECKING, Any

from app.core.config import get_settings
from app.core.constants import lineage_config
from app.core.errors import NotFoundError
from app.core.logging import get_logger
from app.domain.ooxml.lineage import LineageStamp

if TYPE_CHECKING:
    from app.schemas.document import FormatResult

logger = get_logger(__name__)


class HistoryService:
    """转换历史的写入、识别与转换树查询。"""

    def __init__(self, root: Path | None = None) -> None:
        config = lineage_config()
        self.root = root or get_settings().history_dir
        self.root.mkdir(parents=True, exist_ok=True)
        self._files = self.root / str(config["files_subdir"])
        self._files.mkdir(parents=True, exist_ok=True)
        self._nodes_path = self.root / str(config["nodes_filename"])
        self._path = self.root / str(config["records_filename"])
        self._lock = threading.Lock()
        self._stamp = LineageStamp()
        self._modified_message = str(config["modified_message"])
        self._missing_file = str(config["missing_file_message"])

    # ------------------------------------------------------------------
    # 哈希与身份
    # ------------------------------------------------------------------
    @staticmethod
    def file_hash(path: Path) -> str:
        """计算文件 sha256。"""
        digest = hashlib.sha256()
        with Path(path).open("rb") as fp:
            for chunk in iter(lambda: fp.read(1024 * 1024), b""):
                digest.update(chunk)
        return digest.hexdigest()

    def inspect(self, path: Path, filename: str) -> dict[str, Any]:
        """查看转换树，不创建新节点。系统里没见过就当作首次上传。"""
        digest = self.file_hash(path)
        node_id = self._stamp.read(path)
        node, modified = self._match_node(node_id, digest)
        if node is None:
            return {
                "new_document": True,
                "modified": False,
                "modified_message": "",
                "current_node_id": None,
                "root_id": None,
                "path": [],
                "tree": None,
                "file_name": filename,
            }
        tree = self._tree_from(node["node_id"])
        return {
            "new_document": False,
            "modified": modified,
            "modified_message": self._modified_message if modified else "",
            "current_node_id": node["node_id"],
            "root_id": tree["root_id"] if tree else node["node_id"],
            "path": tree["path"] if tree else [self._public_node(node, current=True)],
            "tree": tree["tree"] if tree else self._public_node(node, current=True),
            "file_name": filename,
        }

    def ensure_source(self, path: Path, filename: str) -> dict[str, Any]:
        """转换前确保源文档在树上。没见过就新建根节点，并保存带身份的副本。"""
        digest = self.file_hash(path)
        stamped_id = self._stamp.read(path)
        node, modified = self._match_node(stamped_id, digest)
        if node is not None:
            return {"node": node, "modified": modified, "created": False}

        node_id = stamped_id or uuid.uuid4().hex
        stored = self._store_copy(path, node_id)
        node = self._append_node(
            node_id=node_id,
            file_name=filename,
            file_path=stored,
            ingest_hash=digest,
            stored_hash=self.file_hash(stored),
            parent_id=None,
            origin="upload",
        )
        logger.info("历史节点已建立 %s origin=upload", node_id)
        return {"node": node, "modified": False, "created": True}

    def register_success(
        self,
        *,
        parent_id: str,
        output_path: Path,
        output_name: str,
        task_id: str,
        result: "FormatResult",
        source_name: str,
        source_hash: str,
    ) -> dict[str, Any]:
        """成功产出登记为父节点下的新子节点，并给结果文件盖上新的 node_id。"""
        child_id = uuid.uuid4().hex
        self._stamp.write(output_path, child_id)
        stored = self._store_copy(output_path, child_id)
        stored_hash = self.file_hash(stored)
        self._append_node(
            node_id=child_id,
            file_name=output_name,
            file_path=stored,
            ingest_hash=stored_hash,
            stored_hash=stored_hash,
            parent_id=parent_id,
            origin="conversion",
        )
        return self.record(
            task_id=task_id,
            source_name=source_name,
            source_hash=source_hash,
            output_name=output_name,
            output_path=str(stored),
            success=True,
            message=result.message,
            engine=result.engine,
            duration_ms=result.duration_ms,
            steps=[step.model_dump() for step in result.steps],
            warnings=result.warnings,
            parent_id=parent_id,
            child_id=child_id,
        )

    def register_failure(
        self,
        *,
        parent_id: str,
        task_id: str,
        source_name: str,
        source_hash: str,
        message: str,
        engine: str = "unknown",
        duration_ms: int = 0,
        warnings: list[str] | None = None,
        steps: list[dict[str, Any]] | None = None,
    ) -> dict[str, Any]:
        """失败不创建子文档，只在父节点下挂一片叶子。"""
        return self.record(
            task_id=task_id,
            source_name=source_name,
            source_hash=source_hash,
            output_name=None,
            output_path=None,
            success=False,
            message=message,
            engine=engine,
            duration_ms=duration_ms,
            steps=steps or [],
            warnings=warnings or [],
            parent_id=parent_id,
            child_id=None,
        )

    def node_file(self, node_id: str) -> tuple[Path, str]:
        """返回节点保存的文件。文件被手工删了就明确说。"""
        node = self.get_node(node_id)
        path = Path(str(node.get("file_path") or ""))
        if not path.is_file():
            raise NotFoundError(self._missing_file)
        return path, str(node.get("file_name") or path.name)

    def get_node(self, node_id: str) -> dict[str, Any]:
        for node in self._read_nodes():
            if node.get("node_id") == node_id:
                return node
        raise NotFoundError(f"历史节点不存在: {node_id}")

    # ------------------------------------------------------------------
    # 写入记录（边）
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
        child_id: str | None = None,
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
            "child_id": child_id,
        }

        with self._lock:
            with self._path.open("a", encoding="utf-8") as fp:
                fp.write(json.dumps(entry, ensure_ascii=False) + "\n")

        logger.info("记录历史 %s success=%s parent=%s child=%s", entry["record_id"], success, parent_id, child_id)
        return entry

    def record_format_result(
        self,
        *,
        task_id: str,
        source_name: str,
        source_hash: str,
        result: "FormatResult",
        parent_id: str | None = None,
    ) -> dict[str, Any]:
        """兼容旧调用：有父节点且成功时走 register_success，否则按原样追加。"""
        output_path = Path(result.output_path) if result.output_path else None
        if parent_id and result.success and output_path and output_path.is_file():
            return self.register_success(
                parent_id=parent_id,
                output_path=output_path,
                output_name=output_path.name,
                task_id=task_id,
                result=result,
                source_name=source_name,
                source_hash=source_hash,
            )
        if parent_id and not result.success:
            return self.register_failure(
                parent_id=parent_id,
                task_id=task_id,
                source_name=source_name,
                source_hash=source_hash,
                message=result.message,
                engine=result.engine,
                duration_ms=result.duration_ms,
                warnings=result.warnings,
                steps=[step.model_dump() for step in result.steps],
            )
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

    # ------------------------------------------------------------------
    # 内部：节点匹配与落盘
    # ------------------------------------------------------------------
    def _match_node(self, node_id: str | None, digest: str) -> tuple[dict[str, Any] | None, bool]:
        if node_id:
            found = self._find_node(node_id)
            if found is not None:
                known = {found.get("ingest_hash"), found.get("stored_hash")}
                return found, digest not in known
        by_hash = self._find_node_by_hash(digest)
        return by_hash, False

    def _find_node(self, node_id: str) -> dict[str, Any] | None:
        for node in self._read_nodes():
            if node.get("node_id") == node_id:
                return node
        return None

    def _find_node_by_hash(self, digest: str) -> dict[str, Any] | None:
        for node in self._read_nodes():
            if digest and digest in {node.get("ingest_hash"), node.get("stored_hash")}:
                return node
        return None

    def _store_copy(self, source: Path, node_id: str) -> str:
        dest = self._files / f"{node_id}.docx"
        shutil.copy2(source, dest)
        if self._stamp.read(dest) != node_id:
            self._stamp.write(dest, node_id)
        return str(dest)

    def _append_node(
        self,
        *,
        node_id: str,
        file_name: str,
        file_path: str,
        ingest_hash: str,
        stored_hash: str,
        parent_id: str | None,
        origin: str,
    ) -> dict[str, Any]:
        node = {
            "node_id": node_id,
            "created_at": datetime.now(timezone.utc).isoformat(),
            "file_name": file_name,
            "file_path": file_path,
            "ingest_hash": ingest_hash,
            "stored_hash": stored_hash,
            "parent_id": parent_id,
            "origin": origin,
        }
        with self._lock:
            with self._nodes_path.open("a", encoding="utf-8") as fp:
                fp.write(json.dumps(node, ensure_ascii=False) + "\n")
        return node

    def _read_nodes(self) -> list[dict[str, Any]]:
        return self._read_jsonl(self._nodes_path)

    def _read_all(self) -> list[dict[str, Any]]:
        return self._read_jsonl(self._path)

    def _read_jsonl(self, path: Path) -> list[dict[str, Any]]:
        if not path.exists():
            return []
        records: list[dict[str, Any]] = []
        with self._lock:
            with path.open("r", encoding="utf-8") as fp:
                for line_no, line in enumerate(fp, start=1):
                    line = line.strip()
                    if not line:
                        continue
                    try:
                        records.append(json.loads(line))
                    except json.JSONDecodeError:
                        logger.warning("历史文件 %s 第 %d 行损坏，已跳过", path.name, line_no)
        return records

    # ------------------------------------------------------------------
    # 内部：组树
    # ------------------------------------------------------------------
    def _tree_from(self, node_id: str) -> dict[str, Any] | None:
        nodes = {item["node_id"]: item for item in self._read_nodes() if item.get("node_id")}
        if node_id not in nodes:
            return None
        root_id = node_id
        seen: set[str] = set()
        while nodes[root_id].get("parent_id") and nodes[root_id]["parent_id"] in nodes:
            if root_id in seen:
                break
            seen.add(root_id)
            root_id = nodes[root_id]["parent_id"]

        children: dict[str, list[dict[str, Any]]] = {}
        for item in nodes.values():
            parent = item.get("parent_id")
            if parent:
                children.setdefault(parent, []).append(item)
        for group in children.values():
            group.sort(key=lambda item: item.get("created_at") or "")

        failures: dict[str, list[dict[str, Any]]] = {}
        for record in self._read_all():
            if record.get("success") or not record.get("parent_id"):
                continue
            failures.setdefault(record["parent_id"], []).append(record)
        for group in failures.values():
            group.sort(key=lambda item: item.get("created_at") or "")

        def build(current_id: str) -> dict[str, Any]:
            node = nodes[current_id]
            kids: list[dict[str, Any]] = []
            for child in children.get(current_id, []):
                kids.append(build(child["node_id"]))
            for record in failures.get(current_id, []):
                kids.append(self._public_failure(record))
            return self._public_node(node, current=(current_id == node_id), children=kids)

        path: list[dict[str, Any]] = []
        cursor = node_id
        chain: list[str] = []
        while cursor in nodes:
            chain.append(cursor)
            parent = nodes[cursor].get("parent_id")
            if not parent or parent not in nodes:
                break
            cursor = parent
        chain.reverse()
        for item_id in chain:
            path.append(self._public_node(nodes[item_id], current=(item_id == node_id)))

        return {"root_id": root_id, "path": path, "tree": build(root_id)}

    def _public_node(
        self,
        node: dict[str, Any],
        *,
        current: bool = False,
        children: list[dict[str, Any]] | None = None,
    ) -> dict[str, Any]:
        path = Path(str(node.get("file_path") or ""))
        available = path.is_file()
        node_id = str(node.get("node_id") or "")
        return {
            "kind": "document",
            "node_id": node_id,
            "record_id": None,
            "name": node.get("file_name") or node_id,
            "created_at": node.get("created_at") or "",
            "success": True,
            "message": "",
            "current": current,
            "file_available": available,
            "download_url": f"/api/v1/history/nodes/{node_id}/file" if node_id else None,
            "missing_file_message": "" if available else self._missing_file,
            "children": children or [],
        }

    def _public_failure(self, record: dict[str, Any]) -> dict[str, Any]:
        return {
            "kind": "failed",
            "node_id": None,
            "record_id": record.get("record_id"),
            "name": "转换失败",
            "created_at": record.get("created_at") or "",
            "success": False,
            "message": record.get("message") or "",
            "current": False,
            "file_available": False,
            "download_url": None,
            "missing_file_message": "",
            "children": [],
        }
