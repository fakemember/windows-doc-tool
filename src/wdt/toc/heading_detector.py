"""
标题检测器：基于底层 XML 解析，实现 pStyle + outlineLvl 双通道标题识别与分级。
"""

from __future__ import annotations

import html
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


@dataclass
class Heading:
    level: int          # 1..9
    p_start: int        # document.xml 中 <w:p 起始字节偏移
    p_end: int          # </w:p> 结束偏移（不含）
    text: str           # 段落内全部 <w:t> 文本拼接


_P_TAG_RE = re.compile(rb"<w:p\b.*?</w:p>", re.DOTALL)
_PSTYLE_RE = re.compile(rb'<w:pStyle\b[^>]*?w:val="([^"]+)"', re.DOTALL)
_OUTLINELVL_RE = re.compile(rb'<w:outlineLvl\b[^>]*?w:val="([0-8])"', re.DOTALL)
_T_TAG_RE = re.compile(rb'<w:t\b[^>]*?>(.*?)</w:t>', re.DOTALL)


def _normalize_name(name: str) -> str:
    """去空格并转小写"""
    return re.sub(r"\s+", "", name).lower()


def _is_toc_style(name: str, style_id: str) -> bool:
    """判断是否为目录相关样式（排除目录自身）"""
    norm_name = _normalize_name(name)
    norm_id = _normalize_name(style_id)

    toc_prefixes = ("toc", "目录", "tocheading")
    for prefix in toc_prefixes:
        if norm_name.startswith(prefix) or norm_id.startswith(prefix):
            return True
    return False


def parse_styles_mapping(styles_xml: bytes) -> tuple[dict[str, int], set[str]]:
    """解析 styles.xml，构建 styleId(小写) -> level(1..9) 映射以及被排除的 toc styleId 集合。"""
    style_id_to_level: dict[str, int] = {}
    excluded_style_ids: set[str] = set()

    if not styles_xml:
        return style_id_to_level, excluded_style_ids

    try:
        root = ET.fromstring(styles_xml)
    except ET.ParseError:
        return style_id_to_level, excluded_style_ids

    # 遍历所有 style 元素
    for style_elem in root.findall("w:style", NS):
        style_type = style_elem.attrib.get(f"{{{W_NS}}}type")
        if style_type is not None and style_type != "paragraph":
            continue

        raw_style_id = style_elem.attrib.get(f"{{{W_NS}}}styleId", "")
        style_id_lower = raw_style_id.lower()

        name_elem = style_elem.find("w:name", NS)
        name_val = name_elem.attrib.get(f"{{{W_NS}}}val", "") if name_elem is not None else ""

        # 检查是否为目录样式
        if _is_toc_style(name_val, raw_style_id):
            if style_id_lower:
                excluded_style_ids.add(style_id_lower)
            continue

        level: int | None = None

        # 检查 name 是否命中 heading 1~9 或 标题 1~9
        norm_name = _normalize_name(name_val)
        m_heading = re.match(r"^(?:heading|标题)([1-9])$", norm_name)
        if m_heading:
            level = int(m_heading.group(1))
        else:
            # 检查 pPr 内的 outlineLvl
            p_pr = style_elem.find("w:pPr", NS)
            if p_pr is not None:
                outline_elem = p_pr.find("w:outlineLvl", NS)
                if outline_elem is not None:
                    val = outline_elem.attrib.get(f"{{{W_NS}}}val")
                    if val is not None and val.isdigit():
                        lvl_num = int(val)
                        if 0 <= lvl_num <= 8:
                            level = lvl_num + 1

        if level is not None and style_id_lower:
            style_id_to_level[style_id_lower] = level

    return style_id_to_level, excluded_style_ids


def detect_headings(doc_bytes: bytes, styles_xml: bytes) -> list[Heading]:
    """从 docx 的 document.xml 字节流和 styles.xml 字节流中检测全部标题。

    1. 解析 styles.xml 建立 styleId -> level 映射；
    2. 正则流式提取 document.xml 中的 <w:p>...</w:p>；
    3. 结合 pStyle 与 outlineLvl 双通道判定级别，排除目录自身；
    4. 提取段落内全部 <w:t> 文本拼接并返回。
    """
    style_id_to_level, excluded_style_ids = parse_styles_mapping(styles_xml)
    headings: list[Heading] = []

    for match in _P_TAG_RE.finditer(doc_bytes):
        p_bytes = match.group(0)
        p_start = match.start()
        p_end = match.end()

        # 检查段落中的 pStyle
        pstyle_match = _PSTYLE_RE.search(p_bytes)
        level: int | None = None
        is_excluded = False

        if pstyle_match:
            raw_pstyle = pstyle_match.group(1).decode("utf-8", errors="replace")
            pstyle_lower = raw_pstyle.lower()
            if pstyle_lower in excluded_style_ids or _is_toc_style("", raw_pstyle):
                is_excluded = True
            elif pstyle_lower in style_id_to_level:
                level = style_id_to_level[pstyle_lower]

        if is_excluded:
            continue

        # 如果没有通过 pStyle 确定级别，尝试读取段落自身的 outlineLvl
        if level is None:
            outline_match = _OUTLINELVL_RE.search(p_bytes)
            if outline_match:
                outline_val = int(outline_match.group(1))
                level = outline_val + 1

        if level is not None:
            # 提取段落中的全部 <w:t>
            text_parts: list[str] = []
            for t_match in _T_TAG_RE.finditer(p_bytes):
                raw_t = t_match.group(1).decode("utf-8", errors="replace")
                text_parts.append(html.unescape(raw_t))
            p_text = "".join(text_parts)

            headings.append(
                Heading(
                    level=level,
                    p_start=p_start,
                    p_end=p_end,
                    text=p_text,
                )
            )

    headings.sort(key=lambda h: h.p_start)
    return headings
