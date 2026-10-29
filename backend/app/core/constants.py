"""业务格式常量加载器。

规则：所有可调的格式参数必须落在 backend/config/*.json，代码中禁止硬编码。
这里的职责只是"读取 + 合并默认值 + 缓存"，不包含业务逻辑。
"""

from __future__ import annotations

import json
from functools import lru_cache
from pathlib import Path
from typing import Any

from app.core.config import get_settings


def _load_json(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    with path.open("r", encoding="utf-8") as fp:
        return json.load(fp)


def _merge(base: dict, override: dict) -> dict:
    """深合并：以 base 为底，override 覆盖。"""
    result = dict(base)
    for key, value in override.items():
        if key in result and isinstance(result[key], dict) and isinstance(value, dict):
            result[key] = _merge(result[key], value)
        else:
            result[key] = value
    return result


# ---- 默认值：仅作为配置文件缺失时的兜底，正式调整请改 config/*.json ----

DEFAULT_LAYOUT_RULES: dict[str, Any] = {
    "fonts": {
        "table_default": {"east_asia": "宋体", "ascii": "宋体", "size_pt": 14.0},
        "table_cover": {"east_asia": "宋体", "ascii": "宋体", "size_pt": 9.0},
    },
    "content_template": "content_template.docx",
    "imported_styles": [
        "Normal",
        "heading 1",
        "heading 2",
        "heading 3",
        "heading 4",
        "toc 1",
        "toc 2",
        "toc 3",
    ],
    "tables": {"alignment": "center", "line_twips": 240, "line_rule": "auto"},
    "toc": {
        "default_title": "目      录",
        "keywords": ["目录", "目 录", "目      录", "Contents", "TABLE OF CONTENTS", "CONTENTS"],
        "levels": "1-3",
        "fonts": {
            "title_east_asia": "宋体",
            "title_ascii": "宋体",
            "title_bold": True,
            "entry_east_asia": "宋体",
            "entry_ascii": "Times New Roman",
            "size_pt": 14.0,
        },
    },
    "headings": {
        "max_level": 4,
        "style_name_template": "Heading {level}",
        "space_before_pt": 0.0,
        "space_after_pt": 0.0,
        "number_separator": " ",
        "candidate_max_chars": 60,
        "chapter_title_max_chars": 20,
    },
    "output": {
        "filename_datetime_format": "%y%m%d%H%M%S",
        "allow_chinese_suffix": True,
    },
}

DEFAULT_DOCUMENT_SETTINGS: dict[str, Any] = {
    "page_settings": {
        "paper_size": {"width": 21.0, "height": 29.7, "unit": "cm"},
        "margins": {
            "top": 3.1,
            "bottom": 2.8,
            "left": 2.8,
            "right": 2.8,
            "header": 2.4,
            "footer": 2.4,
            "gutter": 0,
            "unit": "cm",
        },
    },
    "conversion_factors": {"cm_to_twips": 567, "inch_to_twips": 1440},
}

DEFAULT_PICTURE_SETTINGS: dict[str, Any] = {
    "picture_format": {
        "alignment": "center",
        "line_spacing": 1.0,
        "before_spacing": 0,
        "after_spacing": 0,
        "keep_with_next": False,
        "keep_lines": False,
        "picture_width": None,
        "picture_height": None,
        "wrap_type": "inline",
        "skip_cover": True,
    }
}

DEFAULT_COVER_CONFIG: dict[str, Any] = {
    "template_name": "cover_template.docx",
    "designers_marker": "设计人员名单",
    "first_table_row_height_cm": 0.5,
    "document_types": {
        "1": {"primary": "初步设计", "secondary": ""},
        "2": {"primary": "可行性研究", "secondary": ""},
        "3": {"primary": "初步设计", "secondary": "（代可研）"},
        "4": {"primary": "可行性研究", "secondary": "（代初设）"},
    },
    "stage_primary_variants": ["初步设计", "可行性研究"],
    "stage_secondary_variants": ["（代可研）", "（代初设）"],
    "field_placeholders": {
        "project_name": "高性能、安全、环保的磷酸亚铁锂电池材料项目",
        "date": "2024年4月",
        "library_number": "10032357ARPD01AK008",
    },
    "front_page": {
        "issuer_names": ["鞍钢集团工程技术有限公司"],
    },
    "metadata": {
        "merge_company_prefix": True,
        "company_suffix_pattern": "(公司|集团|厂|研究院|设计院|中心)$",
        "project_name_suffixes": ["项目", "工程", "改造", "总承包", "总承包工程", "技改", "建设"],
        "title_window": 4,
        "min_title_length": 4,
        "date_year_month_day_pattern": r"(\d{4}年\d{1,2}月\d{1,2}[日号]|\d{4}-\d{1,2}-\d{1,2}|\d{4}/\d{1,2}/\d{1,2})",
        "date_year_month_pattern": r"(\d{4}年\d{1,2}月)",
        "library_number_pattern": r"库号[:：]\s*([A-Za-z0-9\-_]+)",
        "personnel_role_keywords": ["主管领导", "副总工程师", "总设计师", "总经理", "经理", "主任", "工程师"],
        "design_filter_keywords": ["可行性研究", "初步设计", "代初步", "代可研", "代可行"],
    },
    "format": {
        "title_spacing_keywords": ["初步设计", "可行性研究"],
        "character_spacing_pt": 7.0,
        "library_number": {
            "label_east_asia": "黑体",
            "label_ascii": "黑体",
            "value_east_asia": "Times New Roman",
            "value_ascii": "Times New Roman",
            "size_pt": 16.0,
        },
    },
}

DEFAULT_VALIDATION_RULES: dict[str, Any] = {
    "enabled": True,
    "max_issues": 20,
    "front_page": {
        "issuer_missing_severity": "error",
        "type_invalid_severity": "error",
        "project_missing_severity": "error",
        "page_two_severity": "warning",
    },
    "document_type": {
        "mismatch_severity": "error",
        "stage_keywords": ["初步设计", "可行性研究", "可研", "初设", "代初步", "代可研", "代初设"],
        "allowed_suffixes": ["说明书", "报告"],
    },
    "heading": {
        "number_pattern": r"^\d+(?:\.\d+)*$",
        "separator_pattern": "[ \t\u3000]+",
        "number_like_pattern": r"^[0-9０-９.．。]+",
        "max_excerpt_chars": 40,
        "glued_severity": "warning",
        "invalid_severity": "error",
        "missing_number_severity": "warning",
        "jump_same_chapter_severity": "error",
        "jump_cross_chapter_severity": "warning",
    },
    "messages": {
        "cover_issuer_missing": "封面第一页没有找到单位名称「{issuer}」，无法定位文档类型和项目名称，请检查封面第一页。",
        "front_type_missing": "封面第一页单位名称上方没有文字，这里应写文档类型：初步设计、可行性研究、初步设计（代可研）、可行性研究（代初设）。",
        "front_type_invalid": "封面第一页单位名称上方应为文档类型，实际写的是「{text}」。文档类型只能是：初步设计、可行性研究、初步设计（代可研）、可行性研究（代初设）。",
        "front_project_missing": "封面第一页文档类型「{text}」上方没有项目名称，请补齐后再上传。",
        "page_two_type_mismatch": "封面第二页的文档类型「{page_two}」与第一页「{page_one}」不一致，已以第一页为准同步第二页。",
        "page_two_type_invalid": "封面第二页的文档类型「{page_two}」不是规范写法，已以第一页「{page_one}」为准同步第二页。",
        "document_type_mismatch": "封面上的文档类型是「{detected}」，与选择的「{selected}」不一致，请选择与封面一致的文档类型。",
        "heading_number_invalid": "标题「{text}」的编号不规范：编号只能由数字和点组成，必须以数字开头和结尾，数字与点必须交替出现，并与标题之间留一个空格。",
        "heading_number_glued": "标题「{text}」的编号与标题之间缺少空格，规范写法应为「{number} 标题」。",
        "heading_missing_number": "「{text}」被标记为标题样式，但没有可用编号，已按正文处理。",
        "heading_jump_same_chapter": "标题编号跳号：「{prev}」之后直接出现「{curr}」（{text}），转换已终止。",
        "heading_jump_cross_chapter": "标题编号不连续：「{prev}」之后出现「{curr}」（{text}），请确认是否跳号。",
    },
}


@lru_cache
def cover_config() -> dict[str, Any]:
    """封面配置（单一来源）。"""
    settings = get_settings()
    return _merge(DEFAULT_COVER_CONFIG, _load_json(settings.config_dir / "cover_config.json"))


@lru_cache
def layout_rules() -> dict[str, Any]:
    """排版规则常量。"""
    settings = get_settings()
    return _merge(DEFAULT_LAYOUT_RULES, _load_json(settings.config_dir / "layout_rules.json"))


@lru_cache
def validation_rules() -> dict[str, Any]:
    """源文档规范校验规则。"""
    settings = get_settings()
    return _merge(
        DEFAULT_VALIDATION_RULES,
        _load_json(settings.config_dir / "validation_rules.json"),
    )


@lru_cache
def document_settings() -> dict[str, Any]:
    """页面与单位换算设置。"""
    settings = get_settings()
    return _merge(
        DEFAULT_DOCUMENT_SETTINGS,
        _load_json(settings.config_dir / "document_settings.json"),
    )


@lru_cache
def picture_settings() -> dict[str, Any]:
    """图片格式设置。"""
    settings = get_settings()
    merged = _merge(
        DEFAULT_PICTURE_SETTINGS,
        _load_json(settings.config_dir / "picture_settings.json"),
    )
    return merged.get("picture_format", merged)


def cm_to_twips(value: float) -> int:
    """厘米转 twips（OOXML 页面尺寸单位）。"""
    factor = document_settings()["conversion_factors"]["cm_to_twips"]
    return int(round(value * factor))


def pt_to_half_points(value: float) -> int:
    """磅转半磅（OOXML 字号 w:sz 的单位）。"""
    return int(round(value * 2))
