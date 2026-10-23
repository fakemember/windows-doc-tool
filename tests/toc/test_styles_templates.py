"""
tests/toc/test_styles_templates.py
验证 4 种预设 TOC 样式模板生成的 XML 合法性及关键排版属性。
"""

import xml.etree.ElementTree as ET
import pytest

from wdt.toc.styles_templates import HYPERLINK_STYLE_ID, TocStyle, toc_styles_fragment

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


@pytest.mark.parametrize("style", list(TocStyle))
@pytest.mark.parametrize("depth", [1, 2, 3])
def test_styles_fragment_well_formed(style: TocStyle, depth: int):
    """验证所有样式与不同深度生成的 XML 片段合法且可被 ElementTree 正确解析"""
    fragment = toc_styles_fragment(style, depth)
    full_xml = f'<w:styles xmlns:w="{W_NS}">{fragment}</w:styles>'

    root = ET.fromstring(full_xml)
    style_elements = root.findall("w:style", NS)

    # 包含: 1个 TOC Heading + depth个 TOC{level} + 1个 Hyperlink
    expected_count = 1 + depth + 1
    assert len(style_elements) == expected_count

    style_ids = [s.attrib.get(f"{{{W_NS}}}styleId") for s in style_elements]
    assert "TOC" in style_ids
    assert HYPERLINK_STYLE_ID in style_ids
    for i in range(1, depth + 1):
        assert f"TOC{i}" in style_ids


def test_fragment_skips_existing_styles():
    """原文档已有的样式不重复注入，仅补齐缺失项"""
    existing = (
        '<w:style w:type="paragraph" w:styleId="TOC1"/>'
        '<w:style w:type="character" w:styleId="a7"/>'
    )
    fragment = toc_styles_fragment(TocStyle.CLASSIC_DOT, 3, existing_styles_xml=existing)
    root = ET.fromstring(f'<w:styles xmlns:w="{W_NS}">{fragment}</w:styles>')
    style_ids = [s.attrib.get(f"{{{W_NS}}}styleId") for s in root.findall("w:style", NS)]

    assert "TOC1" not in style_ids  # 已存在，不重复
    assert "a7" not in style_ids  # 与 Word 自动分配的短 ID 撞名风险由独立 ID 规避
    assert "TOC" in style_ids
    assert "TOC2" in style_ids
    assert "TOC3" in style_ids
    assert HYPERLINK_STYLE_ID in style_ids


def test_indentation_progression():
    """验证多级目录样式缩进递增"""
    fragment = toc_styles_fragment(TocStyle.CLASSIC_DOT, depth=3)
    root = ET.fromstring(f'<w:styles xmlns:w="{W_NS}">{fragment}</w:styles>')

    # TOC1 无缩进
    s_toc1 = root.find("./w:style[@w:styleId='TOC1']", NS)
    assert s_toc1 is not None
    p_pr1 = s_toc1.find("w:pPr", NS)
    assert p_pr1 is None or p_pr1.find("w:ind", NS) is None

    # TOC2 缩进 420 dxa
    s_toc2 = root.find("./w:style[@w:styleId='TOC2']", NS)
    assert s_toc2 is not None
    ind2 = s_toc2.find("./w:pPr/w:ind", NS)
    assert ind2 is not None
    assert ind2.attrib.get(f"{{{W_NS}}}left") == "420"

    # TOC3 缩进 840 dxa
    s_toc3 = root.find("./w:style[@w:styleId='TOC3']", NS)
    assert s_toc3 is not None
    ind3 = s_toc3.find("./w:pPr/w:ind", NS)
    assert ind3 is not None
    assert ind3.attrib.get(f"{{{W_NS}}}left") == "840"


def test_modern_accent_highlight():
    """验证 Modern Accent 样式的色彩规范"""
    fragment = toc_styles_fragment(TocStyle.MODERN_ACCENT, depth=2)
    root = ET.fromstring(f'<w:styles xmlns:w="{W_NS}">{fragment}</w:styles>')

    # TOC Heading 颜色为 1F4E79
    s_heading = root.find("./w:style[@w:styleId='TOC']", NS)
    color_h = s_heading.find("./w:rPr/w:color", NS)
    assert color_h is not None
    assert color_h.attrib.get(f"{{{W_NS}}}val") == "1F4E79"

    # TOC1 样式带加粗和 1F4E79
    s_toc1 = root.find("./w:style[@w:styleId='TOC1']", NS)
    b_toc1 = s_toc1.find("./w:rPr/w:b", NS)
    color_toc1 = s_toc1.find("./w:rPr/w:color", NS)
    assert b_toc1 is not None
    assert color_toc1 is not None
    assert color_toc1.attrib.get(f"{{{W_NS}}}val") == "1F4E79"


def test_regulatory_listing_fonts():
    """验证 Regulatory 样式的字体与大标题颜色 2F5496"""
    fragment = toc_styles_fragment(TocStyle.REGULATORY_LISTING, depth=1)
    root = ET.fromstring(f'<w:styles xmlns:w="{W_NS}">{fragment}</w:styles>')

    s_heading = root.find("./w:style[@w:styleId='TOC']", NS)
    color_h = s_heading.find("./w:rPr/w:color", NS)
    assert color_h is not None
    assert color_h.attrib.get(f"{{{W_NS}}}val") == "2F5496"
