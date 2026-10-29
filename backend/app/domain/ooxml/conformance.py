"""源文档规范校验。

设计原则：宁可不转，不可错转。发现不符合排版规范的地方就明确报错，
并给出可定位的问题清单，绝不静默猜测或降级处理。

标题类问题不重复扫描文档，直接复用 HeadingProcessor 的真实判定结果
（HeadingReport），保证「识别口径」与「校验口径」完全一致，杜绝误报。
"""

from __future__ import annotations

from dataclasses import asdict, dataclass

from app.core.constants import cover_config, validation_rules
from app.domain.ooxml.cover import CoverMetadataExtractor
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
        self.front_rules = self.rules["front_page"]
        self.messages = self.rules["messages"]
        self.max_issues = int(self.rules.get("max_issues", 20))
        self.max_excerpt = int(self.heading_rules.get("max_excerpt_chars", 40))

        config = cover_config()
        self._issuers = "、".join(config["front_page"]["issuer_names"])
        self._labels = {
            tid: item["primary"] + item.get("secondary", "")
            for tid, item in config["document_types"].items()
        }

    # ------------------------------------------------------------------
    # 封面文档类型
    # ------------------------------------------------------------------
    def check_document_type(self, selected_type: str) -> list[ValidationIssue]:
        """以封面第一页的单位名定位文档类型和项目名，类型还要与用户选择一致。

        第二页类型不一致只提示：输出封面按模板重建，两页都以第一页为准。
        """
        layout = CoverMetadataExtractor(self.pkg).locate_front_page()
        front = "封面第一页"

        if not layout.issuer_found:
            return self._finish([
                self._issue(
                    self.front_rules.get("issuer_missing_severity", "error"),
                    "cover_issuer_missing",
                    self.messages["cover_issuer_missing"].format(issuer=self._issuers),
                    location=front,
                )
            ])
        if not layout.type_text:
            return self._finish([
                self._issue(
                    self.front_rules.get("type_invalid_severity", "error"),
                    "front_type_missing",
                    self.messages["front_type_missing"],
                    location=front,
                )
            ])
        if layout.type_id is None:
            return self._finish([
                self._issue(
                    self.front_rules.get("type_invalid_severity", "error"),
                    "front_type_invalid",
                    self.messages["front_type_invalid"].format(text=self._clip(layout.type_text)),
                    location=front,
                    excerpt=self._clip(layout.type_text),
                )
            ])
        if not layout.title_lines:
            return self._finish([
                self._issue(
                    self.front_rules.get("project_missing_severity", "error"),
                    "front_project_missing",
                    self.messages["front_project_missing"].format(text=layout.type_text),
                    location=front,
                    excerpt=layout.type_text,
                )
            ])

        page_one = self._labels[layout.type_id]
        issues: list[ValidationIssue] = []
        if str(selected_type) != layout.type_id:
            issues.append(
                self._issue(
                    self.type_rules.get("mismatch_severity", "error"),
                    "document_type_mismatch",
                    self.messages["document_type_mismatch"].format(
                        detected=page_one,
                        selected=self._labels.get(str(selected_type), selected_type),
                    ),
                    location=front,
                    excerpt=layout.type_text,
                )
            )

        if layout.page_two_text and layout.page_two_id != layout.type_id:
            code = "page_two_type_mismatch" if layout.page_two_id else "page_two_type_invalid"
            issues.append(
                self._issue(
                    self.front_rules.get("page_two_severity", "warning"),
                    code,
                    self.messages[code].format(
                        page_two=self._clip(layout.page_two_text), page_one=page_one
                    ),
                    location="封面第二页",
                    excerpt=self._clip(layout.page_two_text),
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

    @staticmethod
    def _same_chapter(prev_number: str, curr_number: str) -> bool:
        return bool(prev_number) and prev_number.split(".")[0] == curr_number.split(".")[0]
