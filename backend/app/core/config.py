"""应用配置。

所有运行期参数集中在此处，通过环境变量前缀 DFS_ 覆盖。
业务格式常量不在此文件，统一放在 config/*.json，见 app/core/constants.py。
"""

from __future__ import annotations

from functools import lru_cache
from pathlib import Path

from pydantic_settings import BaseSettings, SettingsConfigDict

# backend/ 根目录
BACKEND_ROOT = Path(__file__).resolve().parents[2]

CONFIG_DIR = BACKEND_ROOT / "config"
DATA_DIR = BACKEND_ROOT / "data"
LOG_DIR = BACKEND_ROOT / "logs"


class Settings(BaseSettings):
    """全局配置。"""

    model_config = SettingsConfigDict(
        env_prefix="DFS_",
        env_file=str(BACKEND_ROOT / ".env"),
        env_file_encoding="utf-8",
        extra="ignore",
    )

    app_name: str = "高阶智能文档排版系统"
    app_version: str = "1.0.0"
    api_prefix: str = "/api/v1"
    debug: bool = False

    # 目录
    data_dir: Path = DATA_DIR
    config_dir: Path = CONFIG_DIR
    log_dir: Path = LOG_DIR
    template_dir: Path = BACKEND_ROOT / "templates"
    # 种子资源（随代码分发的内置资产，如封面与正文模板），与用户上传模板分开存放
    resources_dir: Path = BACKEND_ROOT / "resources"

    # 上传限制
    max_upload_mb: int = 100
    allowed_extensions: tuple[str, ...] = (".docx", ".doc")

    # 并发：文档排版是 CPU 密集任务，默认并发很低
    worker_max_workers: int = 2
    task_retention_hours: int = 24

    # 跨域：前端开发服务器
    cors_origins: tuple[str, ...] = ("http://localhost:5173", "http://127.0.0.1:5173")

    # 排版引擎：auto | libreoffice | none
    layout_engine: str = "auto"
    libreoffice_binary: str = "soffice"
    libreoffice_timeout_seconds: int = 300

    @property
    def upload_dir(self) -> Path:
        return self.data_dir / "uploads"

    @property
    def output_dir(self) -> Path:
        return self.data_dir / "outputs"

    @property
    def work_dir(self) -> Path:
        return self.data_dir / "work"

    @property
    def history_dir(self) -> Path:
        return self.data_dir / "history"

    def ensure_dirs(self) -> None:
        for path in (
            self.data_dir,
            self.upload_dir,
            self.output_dir,
            self.work_dir,
            self.history_dir,
            self.log_dir,
            self.template_dir,
            self.resources_dir,
        ):
            path.mkdir(parents=True, exist_ok=True)


@lru_cache
def get_settings() -> Settings:
    settings = Settings()
    settings.ensure_dirs()
    return settings
