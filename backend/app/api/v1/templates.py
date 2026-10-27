"""模板管理接口。"""

from __future__ import annotations

from fastapi import APIRouter, File, Form, UploadFile

from app.api.deps import get_template_service
from app.schemas.template import TemplateInfo

router = APIRouter()


@router.get("", summary="模板列表", response_model=list[TemplateInfo])
def list_templates() -> list[TemplateInfo]:
    return [TemplateInfo(**item) for item in get_template_service().list_templates()]


@router.post("", summary="注册模板", response_model=TemplateInfo)
async def register_template(
    file: UploadFile = File(..., description="模板 docx"),
    name: str = Form(..., description="模板名称"),
    kind: str = Form("content", description="content / cover"),
    description: str = Form(""),
    is_default: bool = Form(False),
) -> TemplateInfo:
    data = await file.read()
    await file.close()

    item = get_template_service().register(
        name=name,
        filename=file.filename or "template.docx",
        data=data,
        kind=kind,
        description=description,
        is_default=is_default,
    )
    return TemplateInfo(**item)


@router.delete("/{template_id}", summary="删除模板")
def delete_template(template_id: str) -> dict:
    get_template_service().delete(template_id)
    return {"deleted": template_id}
