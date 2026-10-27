"""模板相关 Schema。"""

from __future__ import annotations

from enum import Enum

from pydantic import BaseModel, Field


class TemplateKind(str, Enum):
    """模板类型。"""

    CONTENT = "content"
    COVER = "cover"


class TemplateInfo(BaseModel):
    """模板元信息。"""

    template_id: str
    name: str
    kind: TemplateKind = TemplateKind.CONTENT
    filename: str
    size_bytes: int = 0
    created_at: str = ""
    is_default: bool = False
    description: str = ""
    tags: list[str] = Field(default_factory=list)
