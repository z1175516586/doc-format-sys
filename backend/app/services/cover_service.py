"""封面服务。

单一职责：给定"原始封面包"，产出"标准模板封面包"。

原实现 utils/cover_replace.py::replace_content_in_cover_by_dify 把这条链路
绑死在远端 Dify 服务上，并按 4 文档类型 × 2 日期型 × 3 长度档维护 24 份模板。
重构后：

- **一份封面模板** ``cover_template.docx``（backend/templates/）
- **一份封面配置** ``cover_config.json``（backend/config/）
- 文档类型、日期、人员都改为运行时回填，不再按组合选模板

"取元数据"被抽成可替换的一步：默认用本地规则（离线可用），需要时再接 Dify。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from app.core.config import get_settings
from app.core.constants import cover_config
from app.core.errors import ValidationError
from app.core.logging import get_logger
from app.domain.ooxml.cover import CoverMetadata, CoverMetadataExtractor, CoverProcessor
from app.domain.ooxml.package import DocxPackage

logger = get_logger(__name__)


@dataclass
class CoverBuildResult:
    """封面构建结果。"""

    package: DocxPackage
    template_name: str = ""
    metadata: CoverMetadata = field(default_factory=CoverMetadata)
    replacements: int = 0
    personnel: int = 0
    table_transplanted: bool = False
    title_spacing: int = 0
    warnings: list[str] = field(default_factory=list)


class CoverService:
    """按规范选择封面模板并回填内容。"""

    def __init__(self, template_dir: Path | None = None) -> None:
        settings = get_settings()
        self.template_dir = Path(template_dir or settings.template_dir)
        self.resources_dir = Path(settings.resources_dir)
        self.legacy_dir = Path(settings.legacy_template_dir)
        self.config = cover_config()

    # ------------------------------------------------------------------
    def build(self, source_cover: DocxPackage, document_type: int = 1) -> CoverBuildResult:
        """用唯一封面模板替换原始封面，并回填项目信息。"""
        metadata = CoverMetadataExtractor(source_cover).extract()
        warnings: list[str] = []

        if metadata.is_empty:
            warnings.append("未能从原始封面识别到任何项目信息，封面将保留模板占位内容")

        template_path = self._resolve_template()
        stage = self._stage_for(document_type)

        package = DocxPackage.open(template_path)
        processor = CoverProcessor(package)

        processor.apply_stage(stage["primary"], stage["secondary"])
        replacements = processor.replace_fields(self._field_mapping(metadata))
        personnel = processor.replace_personnel(metadata.personnel)
        transplanted = processor.transplant_table_after_marker(source_cover)
        spacing = processor.apply_title_spacing()

        logger.info(
            "封面已按模板 %s 重建：阶段=%s%s，字段 %d 处，人员 %d 位，表格移植=%s，字间距 %d 段",
            template_path.name,
            stage["primary"],
            stage["secondary"],
            replacements,
            personnel,
            transplanted,
            spacing,
        )

        return CoverBuildResult(
            package=package,
            template_name=template_path.name,
            metadata=metadata,
            replacements=replacements,
            personnel=personnel,
            table_transplanted=transplanted,
            title_spacing=spacing,
            warnings=warnings,
        )

    # ------------------------------------------------------------------
    def _stage_for(self, document_type: int) -> dict[str, str]:
        types = self.config["document_types"]
        key = str(document_type)
        if key not in types:
            raise ValidationError(f"不支持的文档类型: {document_type}")
        return types[key]

    def _field_mapping(self, metadata: CoverMetadata) -> dict[str, str]:
        placeholders = self.config["field_placeholders"]
        mapping: dict[str, str] = {}
        if metadata.project_name:
            mapping[placeholders["project_name"]] = metadata.project_name
        if metadata.date_text:
            mapping[placeholders["date"]] = metadata.date_text
        if metadata.library_number:
            mapping[placeholders["library_number"]] = metadata.library_number
        return mapping

    # ------------------------------------------------------------------
    def _resolve_template(self) -> Path:
        name = self.config["template_name"]
        for directory in (self.resources_dir, self.template_dir, self.legacy_dir):
            candidate = directory / name
            if candidate.exists():
                return candidate
        raise ValidationError(
            f"封面模板缺失: {name}（已查找 {self.resources_dir} / {self.template_dir} / {self.legacy_dir}）"
        )
