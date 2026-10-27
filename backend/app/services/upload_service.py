"""上传文件落盘服务。

把"校验扩展名 + 限制大小 + 分块写盘"收敛到一处，
避免每个接口各写一遍（原实现中有 3 处重复的 NamedTemporaryFile 逻辑）。
"""

from __future__ import annotations

import uuid
from pathlib import Path

from fastapi import UploadFile

from app.core.config import get_settings
from app.core.errors import ValidationError
from app.core.logging import get_logger

logger = get_logger(__name__)

_CHUNK_SIZE = 1024 * 1024


class UploadService:
    """上传文件的校验与落盘。"""

    def __init__(self, root: Path | None = None) -> None:
        settings = get_settings()
        self.root = root or settings.upload_dir
        self.root.mkdir(parents=True, exist_ok=True)
        self.max_bytes = settings.max_upload_mb * 1024 * 1024
        self.allowed = tuple(settings.allowed_extensions)

    async def save(self, upload: UploadFile, *, prefix: str = "src") -> Path:
        """把上传文件写入 uploads 目录，返回本地路径。

        采用"边读边写 + 字节计数"，避免一次性把大文件读进内存。
        """
        filename = upload.filename or "unnamed.docx"
        suffix = Path(filename).suffix.lower()

        if suffix not in self.allowed:
            raise ValidationError(
                f"仅支持 {', '.join(self.allowed)} 格式，收到 {suffix or '无扩展名'}"
            )

        target = self.root / f"{prefix}_{uuid.uuid4().hex}{suffix}"
        written = 0

        try:
            with target.open("wb") as fp:
                while chunk := await upload.read(_CHUNK_SIZE):
                    written += len(chunk)
                    if written > self.max_bytes:
                        raise ValidationError(
                            f"文件超过大小限制（{self.max_bytes // (1024 * 1024)} MB）"
                        )
                    fp.write(chunk)
        except Exception:
            target.unlink(missing_ok=True)
            raise
        finally:
            await upload.close()

        if written == 0:
            target.unlink(missing_ok=True)
            raise ValidationError("上传文件为空")

        logger.info("上传落盘: %s (%d bytes)", target.name, written)
        return target
