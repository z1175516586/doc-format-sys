"""API v1 路由集合。"""

from fastapi import APIRouter

from app.api.v1 import documents, engines, health, history, tasks, templates

api_router = APIRouter()
api_router.include_router(health.router, tags=["系统"])
api_router.include_router(engines.router, prefix="/engines", tags=["引擎"])
api_router.include_router(documents.router, prefix="/documents", tags=["文档排版"])
api_router.include_router(tasks.router, prefix="/tasks", tags=["任务"])
api_router.include_router(templates.router, prefix="/templates", tags=["模板"])
api_router.include_router(history.router, prefix="/history", tags=["历史记录"])

__all__ = ["api_router"]
