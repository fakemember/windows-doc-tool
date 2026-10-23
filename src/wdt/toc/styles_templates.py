"""
TOC 样式模板库：提供 4 套预设目录样式的 XML 模板片段生成能力。
"""

from __future__ import annotations

from enum import Enum


# Hyperlink 字符样式 ID：刻意避开 Word 自动分配的 a0~a9 短 ID，
# 防止与原文档已有样式撞 ID 导致 Word 报“XML 非法/内容有问题”错误
HYPERLINK_STYLE_ID = "WDTocHyperlink"


class TocStyle(Enum):
    CLASSIC_DOT = "classic"          # 传统点导线样式（宋体/Times，五号字）
    MODERN_ACCENT = "modern"         # 现代商务（微软雅黑/Calibri，一级标题加粗 #1F4E79）
    REGULATORY_LISTING = "regulatory"# 临床 Listing 规范样式（10.5pt 等线/Times）
    MINIMAL_CLEAN = "minimal"        # 极简无导线，空白隔离，页码弱化


def _toc_heading_style(style: TocStyle) -> str:
    """生成 TOC Heading 大标题段落样式"""
    if style == TocStyle.CLASSIC_DOT:
        rfonts = '<w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
        color = '<w:color w:val="000000"/>'
        bold = '<w:b/><w:bCs/>'
    elif style == TocStyle.MODERN_ACCENT:
        rfonts = '<w:rFonts w:ascii="Calibri" w:eastAsia="微软雅黑" w:hAnsi="Calibri" w:cs="Calibri"/>'
        color = '<w:color w:val="1F4E79"/>'
        bold = '<w:b/><w:bCs/>'
    elif style == TocStyle.REGULATORY_LISTING:
        rfonts = '<w:rFonts w:ascii="等线 Light" w:eastAsia="等线 Light" w:hAnsi="等线 Light"/>'
        color = '<w:color w:val="2F5496"/>'
        bold = '<w:b w:val="0"/><w:bCs w:val="0"/>'
    else:  # MINIMAL_CLEAN
        rfonts = '<w:rFonts w:ascii="Arial" w:eastAsia="微软雅黑" w:hAnsi="Arial" w:cs="Arial"/>'
        color = '<w:color w:val="333333"/>'
        bold = '<w:b/><w:bCs/>'

    return (
        f'<w:style w:type="paragraph" w:styleId="TOC">'
        f'<w:name w:val="TOC Heading"/>'
        f'<w:uiPriority w:val="39"/>'
        f'<w:unhideWhenUsed/>'
        f'<w:qFormat/>'
        f'<w:pPr>'
        f'<w:widowControl/>'
        f'<w:spacing w:before="240" w:after="120" w:line="259" w:lineRule="auto"/>'
        f'<w:outlineLvl w:val="9"/>'
        f'</w:pPr>'
        f'<w:rPr>'
        f'{rfonts}'
        f'{bold}'
        f'{color}'
        f'<w:sz w:val="32"/><w:szCs w:val="32"/>'
        f'</w:rPr>'
        f'</w:style>'
    )


def _toc_level_style(style: TocStyle, level: int) -> str:
    """生成 TOC1..TOC9 单个级别段落样式"""
    # 缩进计算
    if level == 1:
        ind_xml = ""
    else:
        left_chars = 200 * (level - 1)
        left_dxa = 420 * (level - 1)
        ind_xml = f'<w:ind w:leftChars="{left_chars}" w:left="{left_dxa}"/>'

    p_pr = f"<w:pPr>{ind_xml}</w:pPr>" if ind_xml else ""

    # 字体与格式参数
    if style == TocStyle.CLASSIC_DOT:
        rfonts = '<w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
        sz = "21"
        r_pr = f'<w:rPr>{rfonts}<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
    elif style == TocStyle.MODERN_ACCENT:
        rfonts = '<w:rFonts w:ascii="Calibri" w:eastAsia="微软雅黑" w:hAnsi="Calibri" w:cs="Calibri"/>'
        sz = "22" if level == 1 else "20"
        bold = "<w:b/><w:bCs/>" if level == 1 else ""
        color = '<w:color w:val="1F4E79"/>' if level == 1 else '<w:color w:val="262626"/>'
        r_pr = f'<w:rPr>{rfonts}{bold}{color}<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
    elif style == TocStyle.REGULATORY_LISTING:
        rfonts = '<w:rFonts w:ascii="Times New Roman" w:eastAsia="等线" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
        sz = "21"
        r_pr = f'<w:rPr>{rfonts}<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
    else:  # MINIMAL_CLEAN
        rfonts = '<w:rFonts w:ascii="Arial" w:eastAsia="微软雅黑" w:hAnsi="Arial" w:cs="Arial"/>'
        sz = "20"
        color = '<w:color w:val="333333"/>'
        r_pr = f'<w:rPr>{rfonts}{color}<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'

    return (
        f'<w:style w:type="paragraph" w:styleId="TOC{level}">'
        f'<w:name w:val="toc {level}"/>'
        f'<w:autoRedefine/>'
        f'<w:uiPriority w:val="39"/>'
        f'<w:unhideWhenUsed/>'
        f'{p_pr}'
        f'{r_pr}'
        f'</w:style>'
    )


def _hyperlink_style() -> str:
    """生成 Hyperlink 字符样式"""
    return (
        f'<w:style w:type="character" w:styleId="{HYPERLINK_STYLE_ID}">'
        '<w:name w:val="Hyperlink"/>'
        '<w:uiPriority w:val="99"/>'
        '<w:unhideWhenUsed/>'
        '<w:rPr>'
        '<w:color w:val="0563C1"/>'
        '<w:u w:val="single"/>'
        '</w:rPr>'
        '</w:style>'
    )


def toc_styles_fragment(
    style: TocStyle, depth: int, existing_styles_xml: str = ""
) -> str:
    """生成可注入 styles.xml 的 XML 片段字符串，仅包含原文档缺失的样式。

    按样式 ID 逐个比对，已存在的样式不重复注入：
    1. TOC Heading 段落样式 (styleId="TOC")
    2. TOC1..TOC{depth} 段落样式
    3. Hyperlink 字符样式 (styleId=HYPERLINK_STYLE_ID)
    """
    safe_depth = max(1, min(depth, 9))
    parts: list[str] = []

    if 'w:styleId="TOC"' not in existing_styles_xml:
        parts.append(_toc_heading_style(style))

    for level in range(1, safe_depth + 1):
        if f'w:styleId="TOC{level}"' not in existing_styles_xml:
            parts.append(_toc_level_style(style, level))

    if f'w:styleId="{HYPERLINK_STYLE_ID}"' not in existing_styles_xml:
        parts.append(_hyperlink_style())

    return "".join(parts)
