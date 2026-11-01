"""文档排版相关 Schema。"""

from __future__ import annotations

from pydantic import BaseModel, Field, model_validator

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
    update_fields_on_open: bool = Field(
        default_factory=lambda: bool(layout_rules()["toc"]["update_fields_on_open"]),
        description="打开文档时让 Word/WPS 自动更新域（页码由客户端算，服务器端零占用）",
    )
    server_refresh_fields: bool = Field(
        default_factory=lambda: bool(layout_rules()["toc"]["server_refresh_fields"]),
        description="用 LibreOffice 在服务器端预刷新目录页码（占全局引擎，内部加锁串行）",
    )

    @model_validator(mode="after")
    def _single_page_source(self) -> "ConvertOptions":
        """页码只能有一个来源。

        两个开关是同一个决策的两种实现：服务端算完落盘，客户端打开时又会重算一遍，
        同时为真只是白占全局单点的排版引擎，且会让「是否需要告警」的判断失去意义。
        前端已是单选，这里再兜一层，防止别的调用方（脚本、旧客户端）造出矛盾状态。
        """
        if self.update_fields_on_open and self.server_refresh_fields:
            raise ValueError(
                "update_fields_on_open 与 server_refresh_fields 互斥，请只启用其中一个"
            )
        return self


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
