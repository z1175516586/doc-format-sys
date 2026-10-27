"""共享依赖。"""

from __future__ import annotations

from functools import lru_cache

from app.services.history_service import HistoryService
from app.services.template_service import TemplateService


@lru_cache
def get_template_service() -> TemplateService:
    return TemplateService()


@lru_cache
def get_history_service() -> HistoryService:
    return HistoryService()
