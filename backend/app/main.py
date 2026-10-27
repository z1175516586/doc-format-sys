"""FastAPI 应用入口。

启动方式：
    cd backend
    uvicorn app.main:app --reload --port 8000
"""

from __future__ import annotations

from contextlib import asynccontextmanager

from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from app.api.v1 import api_router
from app.converters.registry import get_converter
from app.core.config import get_settings
from app.core.errors import AppError
from app.core.logging import get_logger, setup_logging
from app.workers.executor import shutdown_executor

logger = get_logger(__name__)


@asynccontextmanager
async def lifespan(app: FastAPI):
    """启动自检：引擎缺失要在启动时暴露，而不是等用户上传后才报错。"""
    settings = get_settings()
    setup_logging()

    capabilities = get_converter().capabilities()
    logger.info("应用启动: %s v%s", settings.app_name, settings.app_version)
    logger.info("排版引擎: %s (available=%s)", capabilities.name, capabilities.available)

    if not capabilities.available:
        logger.warning("未检测到排版引擎，目录页码将需要手动更新")
    for note in capabilities.notes:
        logger.info("引擎说明: %s", note)

    yield

    shutdown_executor()
    logger.info("应用已停止")


def create_app() -> FastAPI:
    settings = get_settings()

    app = FastAPI(
        title=settings.app_name,
        version=settings.app_version,
        description=(
            "基于 OOXML 标准的 Word 文档解析、格式规范化与自动化排版服务。\n\n"
            "设计目标：不依赖 Microsoft Word / COM，可在 Linux 部署。"
        ),
        lifespan=lifespan,
        docs_url="/docs",
        redoc_url="/redoc",
    )

    app.add_middleware(
        CORSMiddleware,
        allow_origins=list(settings.cors_origins),
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
        expose_headers=["Content-Disposition"],
    )

    app.include_router(api_router, prefix=settings.api_prefix)

    @app.exception_handler(AppError)
    async def handle_app_error(request: Request, exc: AppError) -> JSONResponse:
        """把业务异常统一转换为标准结构，前端只需处理一种错误格式。"""
        logger.warning("业务异常 %s: %s", exc.code, exc.message)
        return JSONResponse(status_code=exc.http_status, content=exc.to_dict())

    @app.get("/", include_in_schema=False)
    def root() -> dict:
        return {
            "app": settings.app_name,
            "version": settings.app_version,
            "docs": "/docs",
            "api": settings.api_prefix,
        }

    return app


app = create_app()
