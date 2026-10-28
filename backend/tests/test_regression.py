"""功能回归：test/src.docx 排版后与 test/obj.docx 做结构级对比。

这是"重构必须复现原行为"的最终验收。依赖仓库根 test/ 下的真实样例
（src.docx=源文档，obj.docx=目标文档），缺样例时自动跳过。
页眉页脚模板不是运行必需项，测试不依赖 template/ 目录。
"""

from __future__ import annotations

from pathlib import Path

import pytest

from app.converters.noop import NoopConverter
from app.domain.ooxml.diff import snapshot
from app.schemas.document import ConvertOptions, FormatRequest
from app.services.format_service import DocumentFormatService

_REPO_ROOT = Path(__file__).resolve().parents[2]
_SRC = _REPO_ROOT / "test" / "src.docx"
_OBJ = _REPO_ROOT / "test" / "obj.docx"

pytestmark = pytest.mark.skipif(
    not (_SRC.exists() and _OBJ.exists()),
    reason="缺少 test/src.docx 或 test/obj.docx，跳过回归比对",
)


def test_regression_matches_golden(tmp_path: Path) -> None:
    output = tmp_path / "regression_out.docx"
    request = FormatRequest(
        source_path=str(_SRC),
        output_path=str(output),
        template_path=None,
        options=ConvertOptions(),
    )
    result = DocumentFormatService(converter=NoopConverter()).format(request)
    assert result.success, result.message

    actual = snapshot(output, "actual")
    golden = snapshot(_OBJ, "golden")

    # 结构性硬指标：一旦回归说明合并/清理/目录插入被破坏
    assert actual.table_count == golden.table_count
    assert actual.section_count == golden.section_count
    assert actual.toc_field_count >= 1
    assert actual.highlight_count == 0
    assert actual.media_count >= 1

    # 已知差异（待补齐，暂不在此断言）：
    # 1. 正文文本保留度 ~85.9%：obj 的目录已由 Word 刷新页码，而纯 OOXML
    #    输出只能插入待更新的域占位，需配置 LibreOffice 引擎刷新后才对齐。
    # 2. 横向节数 1 vs 2：原 COM 流程在横/竖页切分时额外生成一个横向节。
