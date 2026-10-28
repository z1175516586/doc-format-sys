"""源文档规范校验。

设计原则：宁可不转，不可错转。发现不符合排版规范的地方就明确报错，
并给出可定位的问题清单，绝不静默猜测或降级处理。

标题类问题不重复扫描文档，直接复用 HeadingProcessor 的真实判定结果
（HeadingReport），保证「识别口径」与「校验口径」完全一致，杜绝误报。
"""

from __future__ import annotations

from dataclasses import asdict, dataclass

from app.core.constants import cover_config, validation_rules
from app.domain.ooxml.cover import CoverMetadataExtractor, _compact
from app.domain.ooxml.headings import HeadingReport
from app.domain.ooxml.package import DocxPackage


@dataclass
class ValidationIssue:
    """单条校验问题。severity 为 error 时终止转换。"""

    severity: str
    code: str
    message: str
    location: str = ""
    excerpt: str = ""

    def to_dict(self) -> dict[str, str]:
        return asdict(self)


class DocumentValidator:
    """源文档规范校验的单一实体。

    只做「检查 + 报告」，不修改文档；是否终止由调用方根据 severity 决定。
    """

    def __init__(self, pkg: DocxPackage) -> None:
        self.pkg = pkg
        self.rules = validation_rules()
        self.heading_rules = self.rules["heading"]
        self.type_rules = self.rules["document_type"]
        self.messages = self.rules["messages"]
        self.max_issues = int(self.rules.get("max_issues", 20))
        self.max_excerpt = int(self.heading_rules.get("max_excerpt_chars", 40))

        types = cover_config()["document_types"]
        self._primary = {item["primary"] for item in types.values()}
        self._secondary = {item.get("secondary", "") for item in types.values()}
        self._canonical = {
            _compact(item["primary"] + item.get("secondary", "")): tid
            for tid, item in types.items()
        }
        self._labels = {
            tid: item["primary"] + item.get("secondary", "") for tid, item in types.items()
        }

    # ------------------------------------------------------------------
    # 封面文档类型
    # ------------------------------------------------------------------
    def check_document_type(self, selected_type: str) -> list[ValidationIssue]:
        """封面类型必须是四种规范写法之一，且与用户选择一致。"""
        region = self.rules.get("cover_region", {})
        max_chars = int(region.get("type_line_max_chars", 20))
        texts = CoverMetadataExtractor(self.pkg).cover_paragraph_texts()
        keywords = tuple(self.type_rules.get("stage_keywords", ()))

        issues: list[ValidationIssue] = []
        primary = ""
        secondary = ""
        unknown: list[str] = []

        for text in texts:
            compact = self._compact_stage(_compact(text))
            if not compact or len(compact) > max_chars:
                continue
            if compact in self._canonical:
                found_primary, found_secondary = self._split_combined(compact)
                primary = primary or found_primary
                secondary = secondary or found_secondary
                if found_secondary:
                    break
                continue
            if compact in self._primary:
                primary = primary or compact
                continue
            if compact in self._secondary:
                secondary = secondary or compact
                continue
            if not primary and keywords and compact.startswith(keywords):
                unknown.append(compact)

        if not primary:
            # 一个规范写法都没找到：优先报「写了什么但不规范」，否则报「缺类型」
            if unknown:
                issues.extend(
                    self._issue(
                        self.type_rules.get("severity", "error"),
                        "document_type_unknown",
                        self.messages["document_type_unknown"].format(text=text),
                        location="封面",
                        excerpt=text,
                    )
                    for text in unknown
                )
            else:
                issues.append(
                    self._issue(
                        self.type_rules.get("severity", "error"),
                        "document_type_missing",
                        self.messages["document_type_missing"],
                        location="封面",
                    )
                )
            return self._finish(issues)

        detected_id = self._canonical.get(primary + secondary)
        if detected_id is None:
            # 主类型规范但组合不存在（如「初步设计（代初设）」）
            issues.append(
                self._issue(
                    self.type_rules.get("severity", "error"),
                    "document_type_unknown",
                    self.messages["document_type_unknown"].format(text=primary + secondary),
                    location="封面",
                    excerpt=primary + secondary,
                )
            )
            return self._finish(issues)

        if str(selected_type) != detected_id:
            issues.append(
                self._issue(
                    self.type_rules.get("mismatch_severity", "error"),
                    "document_type_mismatch",
                    self.messages["document_type_mismatch"].format(
                        detected=primary + secondary,
                        selected=self._labels.get(str(selected_type), selected_type),
                    ),
                    location="封面",
                    excerpt=primary + secondary,
                )
            )
        return self._finish(issues)

    # ------------------------------------------------------------------
    # 标题编号（复用 HeadingProcessor 的判定结果）
    # ------------------------------------------------------------------
    def collect_heading_issues(self, report: HeadingReport) -> list[ValidationIssue]:
        """把标题识别中记录的异常转成用户可读的问题清单。"""
        issues: list[ValidationIssue] = []

        for ordinal, text in report.malformed:
            issues.append(
                self._issue(
                    self.heading_rules.get("invalid_severity", "error"),
                    "heading_number_invalid",
                    self.messages["heading_number_invalid"].format(text=self._clip(text)),
                    location=f"第 {ordinal} 段",
                    excerpt=self._clip(text),
                )
            )

        for ordinal, text in report.missing_number:
            issues.append(
                self._issue(
                    self.heading_rules.get("missing_number_severity", "warning"),
                    "heading_missing_number",
                    self.messages["heading_missing_number"].format(text=self._clip(text)),
                    location=f"第 {ordinal} 段",
                    excerpt=self._clip(text),
                )
            )

        for ordinal, number, text in report.glued:
            issues.append(
                self._issue(
                    self.heading_rules.get("glued_severity", "warning"),
                    "heading_number_glued",
                    self.messages["heading_number_glued"].format(
                        text=self._clip(text), number=number
                    ),
                    location=f"第 {ordinal} 段",
                    excerpt=self._clip(text),
                )
            )

        for ordinal, prev, curr, text in report.jumps:
            same_chapter = self._same_chapter(prev, curr)
            severity = self.heading_rules.get(
                "jump_same_chapter_severity" if same_chapter else "jump_cross_chapter_severity",
                "error",
            )
            code = "heading_jump_same_chapter" if same_chapter else "heading_jump_cross_chapter"
            issues.append(
                self._issue(
                    severity,
                    code,
                    self.messages[code].format(prev=prev, curr=curr, text=self._clip(text)),
                    location=f"第 {ordinal} 段",
                    excerpt=self._clip(text),
                )
            )

        return self._finish(issues)

    # ------------------------------------------------------------------
    # 辅助
    # ------------------------------------------------------------------
    def _issue(
        self,
        severity: str,
        code: str,
        message: str,
        *,
        location: str = "",
        excerpt: str = "",
    ) -> ValidationIssue:
        return ValidationIssue(
            severity=severity,
            code=code,
            message=message,
            location=location,
            excerpt=excerpt,
        )

    def _finish(self, issues: list[ValidationIssue]) -> list[ValidationIssue]:
        if not self.rules.get("enabled", True):
            return []
        # 规则被设为 off 时不进入报告，避免给出无法处理的提示
        return [item for item in issues if item.severity != "off"][: self.max_issues]

    def _clip(self, text: str) -> str:
        return text if len(text) <= self.max_excerpt else text[: self.max_excerpt] + "…"

    def _compact_stage(self, compact: str) -> str:
        """去掉「说明书」「报告」这类后缀，再与规范类型比对。"""
        for suffix in self.type_rules.get("allowed_suffixes", ()):
            if len(compact) > len(suffix) and compact.endswith(suffix):
                return compact[: -len(suffix)]
        return compact

    def _split_combined(self, compact: str) -> tuple[str, str]:
        for primary in self._primary:
            if compact.startswith(primary):
                rest = compact[len(primary) :]
                return primary, rest if rest in self._secondary else ""
        return compact, ""

    @staticmethod
    def _same_chapter(prev_number: str, curr_number: str) -> bool:
        return bool(prev_number) and prev_number.split(".")[0] == curr_number.split(".")[0]
