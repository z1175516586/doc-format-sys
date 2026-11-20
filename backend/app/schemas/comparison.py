"""转换前后文档对比的返回模型。"""

from typing import Literal

from pydantic import BaseModel, Field

ChangeKind = Literal["equal", "added", "removed", "modified", "style"]


class TextFragment(BaseModel):
    text: str
    changed: bool = False


class ComparisonSide(BaseModel):
    number: int
    text: list[TextFragment]


class StyleChange(BaseModel):
    label: str
    before: str
    after: str
    range_changed: bool = False


class ComparisonRow(BaseModel):
    kind: ChangeKind
    block_type: Literal["paragraph", "table_row"]
    before: ComparisonSide | None = None
    after: ComparisonSide | None = None
    styles: list[StyleChange] = Field(default_factory=list)


class DocumentComparison(BaseModel):
    source_name: str
    output_name: str
    rows: list[ComparisonRow]
    counts: dict[ChangeKind, int]
