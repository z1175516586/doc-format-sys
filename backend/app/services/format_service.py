"""文档排版主流程服务。

对应原 service/converter.py 中 5400 行的 DocumentConverter，
这里按"一个步骤一个方法"重组为可枚举、可计时的流水线。

与原实现的三个本质差异：
1. **不依赖 COM**：目录插入、分节、合并全部用 OOXML 结构操作完成
2. **不做全局状态污染**：原实现替换 sys.stdout 来捕获日志、用 os.getcwd() 存中间文件，
   这里通过 on_progress 回调上报，不碰进程级状态
3. **失败即显式失败**：原实现大量 `return False, "..."` 混在业务里，
   这里每个步骤抛出 AppError，由流水线统一捕获并记录到步骤报告
"""

from __future__ import annotations

import shutil
import time
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

from app.converters.base import DocConverter
from app.converters.registry import get_converter
from app.core.config import get_settings
from app.core.constants import layout_rules
from app.core.errors import AppError, LayoutError, UnsupportedFormatError
from app.core.logging import get_logger
from app.domain.ooxml.document import DocumentMerger, DocumentSplitter
from app.domain.ooxml.header_footer import HeaderFooterProcessor
from app.domain.ooxml.headings import HeadingProcessor
from app.domain.ooxml.highlighting import HighlightProcessor
from app.domain.ooxml.numbering import NumberingProcessor
from app.domain.ooxml.package import DocxPackage
from app.domain.ooxml.pictures import PictureProcessor
from app.domain.ooxml.repair import DocumentRepairer
from app.domain.ooxml.sections import SectionProcessor
from app.domain.ooxml.tables import TableFont, TableProcessor
from app.domain.ooxml.toc import TocProcessor
from app.domain.ooxml.validate import StructuralValidator
from app.schemas.document import ConvertOptions, FormatRequest, FormatResult, StepReport
from app.services.cover_service import CoverService
from app.services.template_service import TemplateService

logger = get_logger(__name__)

ProgressCallback = Callable[[str, int, int, str], None]

# 步骤清单：顺序即执行顺序，名称对前端可见，改动需同步文档
PIPELINE_STEPS: tuple[str, ...] = (
    "open",
    "repair",
    "numbering",
    "orientation",
    "footer",
    "highlight",
    "template",
    "toc_cleanup",
    "headings",
    "split",
    "cover",
    "content",
    "merge",
    "toc_insert",
    "toc_section",
    "picture",
    "finalize",
    # save 必须在 engine 之前：engine 会就地刷新并回写文件，
    # 若 save 在其后再写一次内存包，会把刷新结果覆盖掉。
    "save",
    "engine",
    "validate",
)


@dataclass
class _PipelineState:
    """流水线共享状态。"""

    source_path: Path
    output_path: Path
    options: ConvertOptions
    template_path: Path | None = None
    package: DocxPackage | None = None
    cover: DocxPackage | None = None
    content: DocxPackage | None = None
    toc_anchor: object | None = None
    first_table_row_height: int | None = None
    project_name: str | None = None
    warnings: list[str] = field(default_factory=list)


class DocumentFormatService:
    """文档排版服务的单一实体。"""

    def __init__(
        self,
        converter: DocConverter | None = None,
        template_service: TemplateService | None = None,
    ) -> None:
        self.converter = converter or get_converter()
        self.templates = template_service or TemplateService()

    # ------------------------------------------------------------------
    # 对外入口
    # ------------------------------------------------------------------
    def format(
        self,
        request: FormatRequest,
        on_progress: ProgressCallback | None = None,
    ) -> FormatResult:
        started = time.perf_counter()
        state = _PipelineState(
            source_path=Path(request.source_path),
            output_path=Path(request.output_path),
            options=request.options,
            template_path=Path(request.template_path) if request.template_path else None,
        )

        reports: list[StepReport] = []
        total = len(PIPELINE_STEPS)

        for index, step_name in enumerate(PIPELINE_STEPS, start=1):
            self._emit(on_progress, step_name, index, total, f"开始 {step_name}")
            step_started = time.perf_counter()
            try:
                stats = self._run_step(step_name, state)
                elapsed = int((time.perf_counter() - step_started) * 1000)
                reports.append(
                    StepReport(name=step_name, ok=True, duration_ms=elapsed, stats=stats or {})
                )
                self._emit(on_progress, step_name, index, total, f"{step_name} 完成")
            except AppError as exc:
                elapsed = int((time.perf_counter() - step_started) * 1000)
                reports.append(
                    StepReport(name=step_name, ok=False, message=exc.message, duration_ms=elapsed)
                )
                logger.error("步骤 %s 失败: %s", step_name, exc.message)
                return FormatResult(
                    success=False,
                    steps=reports,
                    warnings=state.warnings,
                    engine=self.converter.capabilities().name,
                    duration_ms=int((time.perf_counter() - started) * 1000),
                    message=f"{step_name} 失败：{exc.message}",
                )
            except Exception as exc:  # 未预期异常也必须转成可读报告
                elapsed = int((time.perf_counter() - step_started) * 1000)
                logger.exception("步骤 %s 出现未预期异常", step_name)
                reports.append(
                    StepReport(
                        name=step_name,
                        ok=False,
                        message=f"未预期异常: {exc}",
                        duration_ms=elapsed,
                    )
                )
                return FormatResult(
                    success=False,
                    steps=reports,
                    warnings=state.warnings,
                    engine=self.converter.capabilities().name,
                    duration_ms=int((time.perf_counter() - started) * 1000),
                    message=f"{step_name} 出现未预期异常",
                )

        return FormatResult(
            success=True,
            output_path=str(state.output_path),
            steps=reports,
            warnings=state.warnings,
            engine=self.converter.capabilities().name,
            duration_ms=int((time.perf_counter() - started) * 1000),
            message="文档排版完成",
        )

    @staticmethod
    def _emit(
        callback: ProgressCallback | None,
        step: str,
        index: int,
        total: int,
        message: str,
    ) -> None:
        if callback is not None:
            try:
                callback(step, index, total, message)
            except Exception:  # 进度回调不得影响主流程
                logger.debug("进度回调异常", exc_info=True)

    # ------------------------------------------------------------------
    # 步骤分发
    # ------------------------------------------------------------------
    def _run_step(self, name: str, state: _PipelineState) -> dict[str, object]:
        handler = getattr(self, f"_step_{name}", None)
        if handler is None:
            raise LayoutError(f"未实现的流水线步骤: {name}")
        return handler(state)

    # ------------------------------------------------------------------
    # 各步骤
    # ------------------------------------------------------------------
    def _step_open(self, state: _PipelineState) -> dict[str, object]:
        if not state.source_path.exists():
            raise LayoutError(f"源文件不存在: {state.source_path.name}")

        suffix = state.source_path.suffix.lower()
        allowed = tuple(get_settings().allowed_extensions)
        if suffix not in allowed:
            raise UnsupportedFormatError(f"仅支持 {', '.join(allowed)}，收到 {suffix}")

        if suffix == ".doc":
            state.source_path = self._convert_doc_to_docx(state.source_path)

        state.package = DocxPackage.open(state.source_path)
        state.output_path.parent.mkdir(parents=True, exist_ok=True)
        return {
            "parts": len(state.package.part_names),
            "sections": len(state.package.section_properties()),
            "paragraphs": len(state.package.paragraphs()),
            "tables": len(state.package.tables()),
        }

    def _convert_doc_to_docx(self, source: Path) -> Path:
        """老 .doc 格式必须经由排版引擎转换，纯库无法处理二进制 doc。"""
        capabilities = self.converter.capabilities()
        if not capabilities.doc_to_docx:
            raise UnsupportedFormatError(
                "当前环境无法转换 .doc，请先另存为 .docx 或配置 LibreOffice 引擎"
            )
        target = source.with_suffix(".docx")
        self.converter.convert(source, target, "docx")
        return target

    def _step_repair(self, state: _PipelineState) -> dict[str, object]:
        return DocumentRepairer(self._pkg(state)).run_all()

    def _step_numbering(self, state: _PipelineState) -> dict[str, object]:
        if not state.options.convert_numbering:
            return {"skipped": True}
        converted = NumberingProcessor(self._pkg(state)).convert()
        return {"paragraphs_converted": converted}

    def _step_orientation(self, state: _PipelineState) -> dict[str, object]:
        processor = SectionProcessor(self._pkg(state))
        fixed = processor.normalize_orientation()
        return {"fixed_sections": fixed, "sequence": processor.orientation_sequence()}

    def _step_footer(self, state: _PipelineState) -> dict[str, object]:
        cleared = HeaderFooterProcessor(self._pkg(state)).clear_all_footers()
        return {"footers_cleared": cleared}

    def _step_highlight(self, state: _PipelineState) -> dict[str, object]:
        if not state.options.remove_highlights:
            return {"skipped": True}
        return HighlightProcessor(self._pkg(state)).remove_all()

    def _step_template(self, state: _PipelineState) -> dict[str, object]:
        if not state.options.apply_template:
            return {"skipped": True}

        template_path = state.template_path or self.templates.default_template_path()
        if template_path is None:
            state.warnings.append("未提供模板且无默认模板，已跳过页眉页脚移植")
            return {"skipped": True, "reason": "no_template"}

        warnings = self.templates.apply_header_footer(self._pkg(state), template_path)
        state.warnings.extend(warnings)
        return {"template": template_path.name, "warnings": len(warnings)}

    def _step_toc_cleanup(self, state: _PipelineState) -> dict[str, object]:
        removed = TocProcessor(self._pkg(state)).remove_toc_blocks()
        return {"toc_blocks_removed": removed}

    def _step_headings(self, state: _PipelineState) -> dict[str, object]:
        """识别并标注标题样式。

        必须在目录清理之后、拆分之前：
        - 之前：旧目录条目形如 "1 总论 …… 1"，会被误判为标题
        - 之前拆分：拆分要靠标题样式定位正文起点
        """
        report = HeadingProcessor(self._pkg(state)).apply_heading_styles()
        if report.detected == 0:
            state.warnings.append(
                "未识别到正文标题：目录可能为空，且分节定位会回落到编号文本推断"
            )
        return {
            "heading_count": report.detected,
            "levels": report.levels,
            "reset_to_normal": report.reset_to_normal,
        }

    def _step_split(self, state: _PipelineState) -> dict[str, object]:
        result = DocumentSplitter(self._pkg(state)).split()
        state.cover = result.cover
        state.content = result.content
        state.project_name = result.project_name
        return {
            "reason": result.split_reason,
            "project_name": result.project_name or "",
        }

    def _step_cover(self, state: _PipelineState) -> dict[str, object]:
        """用标准模板封面整块替换原封面并回填项目信息，或就地清理原封面。

        顺序很关键：模板封面的版式由模板保证，原封面只作为
        元数据与设计人员表的来源，不再参与最终排版。

        当 ``use_cover_template=False``（例如单元测试用合成文档）时，
        跳过模板替换，直接在原封面上做表清理/行高压缩/字体归一。
        """
        source_cover = self._require(state.cover, "封面包")

        if state.options.use_cover_template:
            result = CoverService(self.templates.root).build(
                source_cover,
                document_type=state.options.document_type,
            )
            state.cover = result.package
            state.warnings.extend(result.warnings)
            cover_info: dict[str, object] = {
                "template": result.template_name,
                "project_name": result.metadata.project_name,
                "library_number": result.metadata.library_number,
                "date": result.metadata.date_text,
                "personnel": len(result.metadata.personnel),
                "replacements": result.replacements,
                "table_transplanted": result.table_transplanted,
            }
        else:
            state.cover = source_cover
            cover_info = {"template": "inplace"}

        tables = TableProcessor(state.cover)
        removed = tables.remove_tables_before_marker()
        state.first_table_row_height = tables.shrink_first_table_rows()
        runs = tables.apply_font_to_first_table(TableFont.from_config("table_cover"))

        return {
            **cover_info,
            "tables_removed": removed,
            "font_runs": runs,
            "original_row_height": state.first_table_row_height,
        }

    def _step_content(self, state: _PipelineState) -> dict[str, object]:
        content = self._require(state.content, "正文包")

        pictures = PictureProcessor(content)
        reset = pictures.reset_indent()
        aligned = pictures.apply_alignment()

        tables = TableProcessor(content)
        runs = tables.apply_font_to_all_tables(TableFont.from_config("table_default"))

        return {
            "picture_paragraphs": pictures.count(),
            "indent_reset": reset,
            "pictures_aligned": aligned,
            "table_font_runs": runs,
        }

    def _step_merge(self, state: _PipelineState) -> dict[str, object]:
        cover = self._require(state.cover, "封面包")
        content = self._require(state.content, "正文包")

        report = DocumentMerger(cover).append(content)
        state.package = cover
        state.cover = None
        state.content = None

        if report.warnings:
            state.warnings.extend(report.warnings)

        return {
            "appended_elements": report.appended_body_elements,
            "remapped_relationships": report.remapped_relationships,
            "rewritten_references": report.rewritten_references,
            "merged_styles": report.merged_styles,
            "sections": report.final_sections,
        }

    def _step_toc_insert(self, state: _PipelineState) -> dict[str, object]:
        if not state.options.insert_toc:
            return {"skipped": True}

        processor = TocProcessor(self._pkg(state))
        if processor.has_toc_field():
            state.warnings.append("文档已存在目录域，本次未重复插入")
            return {"skipped": True, "reason": "already_present"}

        inserted = processor.insert_field(title=state.options.toc_title)
        if not inserted:
            raise LayoutError("未能定位正文起始标题，无法插入目录")

        state.toc_anchor = self._find_toc_paragraph(state)
        return {"inserted": True, "title": state.options.toc_title}

    def _step_toc_section(self, state: _PipelineState) -> dict[str, object]:
        """让目录独占一节，便于正文单独编页码。"""
        if state.toc_anchor is None:
            return {"skipped": True}

        processor = SectionProcessor(self._pkg(state))
        inserted = processor.insert_break_after(state.toc_anchor)
        return {"inserted": inserted}

    def _step_picture(self, state: _PipelineState) -> dict[str, object]:
        processor = PictureProcessor(self._pkg(state))
        processor.reset_indent()
        return {"picture_paragraphs": processor.count()}

    def _step_finalize(self, state: _PipelineState) -> dict[str, object]:
        pkg = self._pkg(state)
        tables = TableProcessor(pkg)
        restored = tables.restore_first_table_row_height(state.first_table_row_height)
        removed_keep = DocumentRepairer(pkg).remove_heading_keep_options()

        sections = SectionProcessor(pkg).describe()
        return {
            "row_height_restored": restored,
            "keep_options_removed": removed_keep,
            "sections": len(sections),
            "landscape_sections": sum(1 for info in sections if info.is_landscape),
        }

    def _step_engine(self, state: _PipelineState) -> dict[str, object]:
        """交给排版引擎刷新域。

        引擎不支持时必须显式告警：目录页码需要用户在 Word 中按 F9 更新。
        这正是原实现"静默跳过 COM 不可用"的地方，会造成交付文档目录页码为 0。
        """
        capabilities = self.converter.capabilities()
        if not capabilities.available:
            state.warnings.append(
                "未配置排版引擎：目录页码未刷新，请在 Word/WPS 中右键目录选择「更新域」"
            )
            return {"skipped": True, "reason": "engine_unavailable"}

        if not capabilities.refresh_fields:
            state.warnings.append(
                "当前引擎不支持域刷新：目录页码需手动更新（安装 python-uno 可自动化）"
            )
            return {"skipped": True, "reason": "refresh_unsupported"}

        pkg = self._pkg(state)
        pkg.save(state.output_path)
        refreshed = self.converter.refresh_fields(state.output_path)
        if not refreshed:
            state.warnings.append("域刷新未成功，目录页码可能仍是占位内容")
        return {"refreshed": refreshed}

    def _step_save(self, state: _PipelineState) -> dict[str, object]:
        pkg = self._pkg(state)
        saved = pkg.save(state.output_path)
        return {"path": str(saved), "size_bytes": saved.stat().st_size}

    def _step_validate(self, state: _PipelineState) -> dict[str, object]:
        """输出结构断言：不通过即判定失败，绝不静默交付坏文档。"""
        report = StructuralValidator().validate(state.output_path)
        if not report.ok:
            raise LayoutError("输出结构校验未通过: " + "；".join(report.errors))
        return dict(report.stats)

    # ------------------------------------------------------------------
    # 辅助
    # ------------------------------------------------------------------
    @staticmethod
    def _pkg(state: _PipelineState) -> DocxPackage:
        if state.package is None:
            raise LayoutError("文档尚未加载")
        return state.package

    @staticmethod
    def _require(pkg: DocxPackage | None, label: str) -> DocxPackage:
        if pkg is None:
            raise LayoutError(f"{label}未就绪，流水线顺序异常")
        return pkg

    @staticmethod
    def _find_toc_paragraph(state: _PipelineState):
        """定位刚插入的 TOC 域所在段落，用于后续插入分节符。"""
        pkg = state.package
        if pkg is None:
            return None
        for paragraph in pkg.paragraphs():
            for instr in paragraph.iter():
                if instr.tag.endswith("instrText") and instr.text and "TOC" in instr.text.upper():
                    return paragraph
        return None


def build_output_path(source_name: str, output_dir: Path | None = None) -> Path:
    """生成输出文件名。

    规则沿用原实现：YYMMDDHHMMSS + 原文件名最后一段（若含中文），
    保证输出名可读且不冲突。
    """
    settings = get_settings()
    rules = layout_rules()["output"]

    fmt = rules["filename_datetime_format"]
    allow_chinese = bool(rules["allow_chinese_suffix"])

    stem = Path(source_name).stem
    last_part = stem.split("_")[-1] if stem else stem
    has_chinese = any("\u4e00" <= ch <= "\u9fff" for ch in last_part)

    prefix = time.strftime(fmt)
    suffix = last_part if (has_chinese and allow_chinese) else uuid.uuid4().hex[:6]

    target_dir = output_dir or settings.output_dir
    target_dir.mkdir(parents=True, exist_ok=True)
    return target_dir / f"{prefix}_{suffix}.docx"


def copy_upload(source: Path, target: Path) -> Path:
    """把上传文件落到持久化目录，返回新路径。"""
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, target)
    return target
