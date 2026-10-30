"""文档排版相关 Schema。"""

from __future__ import annotations

from pydantic import BaseModel, Field

from app.core.constants import layout_rules


class ConvertOptions(BaseModel):
    """排版选项。默认值来自 config/layout_rules.json，不在代码里硬编码。"""

    toc_title: str = Field(
        default_factory=lambda: layout_rules()["toc"]["default_title"],
        description="目录标题",
    )
    header_text: str = Field(default="", description="页眉文本")
    header_left_text: str = Field(default="", description="横版页眉左侧文本")
    header_right_text: str = Field(default="", description="横版页眉右侧文本")
    document_type: int = Field(default=1, ge=1, le=4, description="文档类型：1 初步设计 / 2 可行性研究 / 3 初步设计（代可研）/ 4 可行性研究（代初设）")
    insert_toc: bool = Field(default=True, description="是否插入目录域")
    use_cover_template: bool = Field(default=True, description="是否用标准封面模板替换原封面（否则就地清理原封面）")
    remove_highlights: bool = Field(default=True, description="是否清除高亮与底纹")
    convert_numbering: bool = Field(default=True, description="是否把自动编号固化为文本")
    apply_template: bool = Field(default=True, description="是否应用模板的页眉页脚")


class FormatRequest(BaseModel):
    """内部使用的排版请求（由 API 层组装，不直接暴露）。"""

    source_path: str
    output_path: str
    template_path: str | None = None
    source_name: str = ""
    source_node_id: str | None = None
    options: ConvertOptions = Field(default_factory=ConvertOptions)


class StepReport(BaseModel):
    """单个排版步骤的执行报告。"""

    name: str
    ok: bool
    message: str = ""
    duration_ms: int = 0
    stats: dict[str, object] = Field(default_factory=dict)


class FormatResult(BaseModel):
    """排版结果。"""

    success: bool
    output_path: str | None = None
    steps: list[StepReport] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)
    issues: list[dict[str, object]] = Field(default_factory=list)
    engine: str = "none"
    duration_ms: int = 0
    message: str = ""
