"""
TOC 样式模板库：提供 4 套预设目录样式的 XML 模板片段生成能力。

4 种经典/申报/现代排版规范：
- REGULATORY_LISTING: 临床申报紧凑风 (宋体/Times, 10.5pt, 密排点导线, 满足 FDA/NMPA eCTD)
- MODERN_ACCENT: 现代商务海蓝风 (微软雅黑/Calibri, 11pt, 一级海蓝加粗, 疏排导线)
- CLASSIC_DOT: 经典学术风 (宋体/Times, 10.5pt, 传统细点导线, 优雅页码)
- MINIMAL_CLEAN: 极简纯净风 (微软雅黑/Arial, 10pt, 无点导线, 灰阶隔离)
"""

from __future__ import annotations

from enum import Enum


HYPERLINK_STYLE_ID = "WDTocHyperlink"


class TocStyle(Enum):
    REGULATORY_LISTING = "regulatory"  # 临床申报规范 (宋体/Times，五号/10.5pt，紧凑密集点导线)
    MODERN_ACCENT = "modern"           # 现代商务 (微软雅黑/Calibri，一级标题加粗 #1F4E79，高级灰导线)
    CLASSIC_DOT = "classic"            # 传统点导线样式 (宋体/Times，五号字，典雅居中大标题)
    MINIMAL_CLEAN = "minimal"          # 极简纯净 (无导线，依靠对齐与低饱和度灰阶隔离)


def _toc_heading_style(style: TocStyle) -> str:
    """生成 TOC Heading 大标题段落样式"""
    if style == TocStyle.REGULATORY_LISTING:
        rfonts = '<w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
        color = '<w:color w:val="000000"/>'
        bold = '<w:b/><w:bCs/>'
        sz = "22"  # 11pt / 小四
        spacing = '<w:spacing w:before="180" w:after="120" w:line="240" w:lineRule="auto"/>'
        jc = '<w:jc w:val="center"/>'
    elif style == TocStyle.MODERN_ACCENT:
        rfonts = '<w:rFonts w:ascii="Calibri" w:eastAsia="微软雅黑" w:hAnsi="Calibri" w:cs="Calibri"/>'
        color = '<w:color w:val="1F4E79"/>'
        bold = '<w:b/><w:bCs/>'
        sz = "30"  # 15pt / 小二
        spacing = '<w:spacing w:before="240" w:after="160" w:line="280" w:lineRule="auto"/>'
        jc = '<w:jc w:val="left"/>'
    elif style == TocStyle.CLASSIC_DOT:
        rfonts = '<w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
        color = '<w:color w:val="000000"/>'
        bold = '<w:b/><w:bCs/>'
        sz = "28"  # 14pt / 三号
        spacing = '<w:spacing w:before="240" w:after="140" w:line="260" w:lineRule="auto"/>'
        jc = '<w:jc w:val="center"/>'
    else:  # MINIMAL_CLEAN
        rfonts = '<w:rFonts w:ascii="Arial" w:eastAsia="微软雅黑" w:hAnsi="Arial" w:cs="Arial"/>'
        color = '<w:color w:val="333333"/>'
        bold = '<w:b/><w:bCs/>'
        sz = "26"  # 13pt
        spacing = '<w:spacing w:before="200" w:after="120" w:line="260" w:lineRule="auto"/>'
        jc = '<w:jc w:val="left"/>'

    return (
        f'<w:style w:type="paragraph" w:styleId="TOC">'
        f'<w:name w:val="TOC Heading"/>'
        f'<w:uiPriority w:val="39"/>'
        f'<w:unhideWhenUsed/>'
        f'<w:qFormat/>'
        f'<w:pPr>'
        f'<w:widowControl/>{jc}{spacing}'
        f'<w:outlineLvl w:val="9"/>'
        f'</w:pPr>'
        f'<w:rPr>'
        f'{rfonts}{bold}{color}'
        f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
        f'</w:rPr>'
        f'</w:style>'
    )


def _toc_level_style(style: TocStyle, level: int) -> str:
    """生成 TOC1..TOC9 单个级别段落样式"""
    # 层级缩进 (TOC1 为 0，TOC2 为 420 dxa ~ 2字符，TOC3 为 840 dxa)
    if level == 1:
        ind_xml = ""
    else:
        left_chars = 200 * (level - 1)
        left_dxa = 420 * (level - 1)
        ind_xml = f'<w:ind w:leftChars="{left_chars}" w:left="{left_dxa}"/>'

    # 行距配置
    if style == TocStyle.REGULATORY_LISTING:
        spacing_xml = '<w:spacing w:before="30" w:after="30" w:line="240" w:lineRule="auto"/>'
    elif style == TocStyle.MODERN_ACCENT:
        spacing_xml = '<w:spacing w:before="80" w:after="40" w:line="280" w:lineRule="auto"/>' if level == 1 else '<w:spacing w:before="20" w:after="20" w:line="260" w:lineRule="auto"/>'
    elif style == TocStyle.CLASSIC_DOT:
        spacing_xml = '<w:spacing w:before="60" w:after="40" w:line="260" w:lineRule="auto"/>' if level == 1 else '<w:spacing w:before="20" w:after="20" w:line="240" w:lineRule="auto"/>'
    else:  # MINIMAL_CLEAN
        spacing_xml = '<w:spacing w:before="50" w:after="30" w:line="270" w:lineRule="auto"/>'

    p_pr = f"<w:pPr>{ind_xml}{spacing_xml}</w:pPr>" if (ind_xml or spacing_xml) else ""

    # 字体与格式参数
    if style == TocStyle.REGULATORY_LISTING:
        rfonts = '<w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
        sz = "21" if level == 1 else "20"  # 10.5pt (五号)
        bold = "<w:b/><w:bCs/>" if level == 1 else ""
        color = '<w:color w:val="000000"/>'
        r_pr = f'<w:rPr>{rfonts}{bold}{color}<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
    elif style == TocStyle.MODERN_ACCENT:
        rfonts = '<w:rFonts w:ascii="Calibri" w:eastAsia="微软雅黑" w:hAnsi="Calibri" w:cs="Calibri"/>'
        sz = "22" if level == 1 else "20"
        bold = "<w:b/><w:bCs/>" if level == 1 else ""
        color = '<w:color w:val="1F4E79"/>' if level == 1 else '<w:color w:val="262626"/>'
        r_pr = f'<w:rPr>{rfonts}{bold}{color}<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
    elif style == TocStyle.CLASSIC_DOT:
        rfonts = '<w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
        sz = "21"
        bold = "<w:b/><w:bCs/>" if level == 1 else ""
        color = '<w:color w:val="111827"/>'
        r_pr = f'<w:rPr>{rfonts}{bold}{color}<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
    else:  # MINIMAL_CLEAN
        rfonts = '<w:rFonts w:ascii="Arial" w:eastAsia="微软雅黑" w:hAnsi="Arial" w:cs="Arial"/>'
        sz = "20"
        bold = "<w:b/><w:bCs/>" if level == 1 else ""
        color = '<w:color w:val="334155"/>'
        r_pr = f'<w:rPr>{rfonts}{bold}{color}<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'

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
        '<w:u w:val="none"/>'
        '</w:rPr>'
        '</w:style>'
    )


def toc_styles_fragment(
    style: TocStyle, depth: int, existing_styles_xml: str = ""
) -> str:
    """生成可注入 styles.xml 的 XML 片段字符串，仅包含原文档缺失的样式。"""
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
