r"""
旧目录剥离器：字节级清理 Word 生成的旧目录结构：
1. sdt 内容控件块（含 Table of Contents 或 TOC \o）
2. 跨段落 TOC \o 字段段落
3. SAS / 临床试验模板纯文本占位符目录（Table of Contents + From Microsoft Word 占位节）
4. 前置标题段与孤立换页符吸收
"""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET

_HEADING_TEXT_NORMALIZED = {"目录", "tableofcontents", "contents"}
_HEADING_STYLES_LOWER = {"toc heading", "tocheading", "目录标题"}


def _extract_text_and_style_from_p(p_bytes: bytes) -> tuple[str, str]:
    """从段落字节流中提取纯文本（去除空白转小写）与段落样式 styleId。"""
    try:
        root = ET.fromstring(p_bytes)
    except Exception:
        t_matches = re.findall(r"<w:t\b[^>]*>(.*?)</w:t>", p_bytes.decode("utf-8", errors="replace"))
        text = "".join(t_matches).strip().replace(" ", "").lower()
        m_style = re.search(r'<w:pStyle\b[^>]*?w:val="([^"]+)"', p_bytes.decode("utf-8", errors="replace"))
        style_val = m_style.group(1).strip() if m_style else ""
        return text, style_val

    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    texts = [node.text or "" for node in root.findall(".//w:t", ns)]
    full_text = "".join(texts).strip().replace(" ", "").lower()

    pstyle = root.find(".//w:pPr/w:pStyle", ns)
    style_val = ""
    if pstyle is not None:
        style_val = pstyle.attrib.get(f"{{{ns['w']}}}val", "")

    return full_text, style_val


def _get_toc_heading_style_ids(styles_xml: str | bytes) -> set[str]:
    """从 styles.xml 中解析名称属于 TOC Heading 的 styleId 集合。"""
    matching_ids: set[str] = set()
    if not styles_xml:
        return matching_ids

    xml_text = styles_xml.decode("utf-8", errors="replace") if isinstance(styles_xml, bytes) else styles_xml
    try:
        root = ET.fromstring(xml_text)
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        for style in root.findall(".//w:style", ns):
            style_id = style.attrib.get(f"{{{ns['w']}}}styleId", "")
            name_node = style.find("./w:name", ns)
            name_val = name_node.attrib.get(f"{{{ns['w']}}}val", "") if name_node is not None else ""
            if name_val.strip().lower() in _HEADING_STYLES_LOWER or style_id.strip().lower() in _HEADING_STYLES_LOWER:
                matching_ids.add(style_id)
                matching_ids.add(style_id.lower())
    except Exception:
        pass

    return matching_ids


def _rfind_p_open(data: bytes, end_pos: int) -> int:
    """向前查找真正的段落起始标签 <w:p> 或 <w:p ...>（严格排除 <w:pPr>、<w:pStyle> 等）。"""
    pos = end_pos
    while True:
        idx = data.rfind(b"<w:p", 0, pos)
        if idx == -1:
            return -1
        next_byte = data[idx + 4 : idx + 5]
        if next_byte in (b">", b" ", b"\t", b"\r", b"\n"):
            return idx
        pos = idx


def _find_p_open(data: bytes, start_pos: int) -> int:
    """向后查找真正的段落起始标签 <w:p> 或 <w:p ...>（严格排除 <w:pPr>、<w:pStyle> 等）。"""
    pos = start_pos
    while True:
        idx = data.find(b"<w:p", pos)
        if idx == -1:
            return -1
        next_byte = data[idx + 4 : idx + 5]
        if next_byte in (b">", b" ", b"\t", b"\r", b"\n"):
            return idx
        pos = idx + 4


def _is_isolated_page_break_paragraph(p_bytes: bytes) -> bool:
    """判断段落是否为仅含硬分页符 <w:br w:type="page"/> 的孤立空段落。"""
    if b"<w:sectPr" in p_bytes:
        return False

    if b'<w:br w:type="page"/>' not in p_bytes and b'<w:br w:type="page" />' not in p_bytes:
        return False

    m_texts = re.findall(rb"<w:t\b[^>]*>(.*?)</w:t>", p_bytes)
    all_text = b"".join(m_texts).strip()
    return len(all_text) == 0


def strip_existing_toc(doc_bytes: bytes, styles_xml: str | bytes = "") -> bytes:
    """字节级剥离 Word 生成的旧目录结构及 SAS / 模板目录占位符节。

    :param doc_bytes: word/document.xml 原始字节流
    :param styles_xml: word/styles.xml 原始文本或字节
    :return: 剥离旧目录后的 document.xml 字节流；无旧目录则原样返回
    """
    if not doc_bytes:
        return doc_bytes

    toc_style_ids = _get_toc_heading_style_ids(styles_xml)
    ranges_to_remove: list[tuple[int, int]] = []

    # -------------------------------------------------------------------------
    # 1. 扫描 SAS / 临床试验模板纯文本占位符目录（如 Table of Contents + From Microsoft Word）
    # -------------------------------------------------------------------------
    m_ph = doc_bytes.find(b"From Microsoft Word")
    if m_ph == -1:
        m_ph = doc_bytes.find(b"index and content")

    if m_ph != -1:
        # 向前寻找 Table of Contents 标题段或占位符起始段
        p_start = _rfind_p_open(doc_bytes, m_ph)
        m_toc = doc_bytes.rfind(b"Table of Contents", 0, m_ph)
        if m_toc != -1:
            p_toc = _rfind_p_open(doc_bytes, m_toc)
            if p_toc != -1:
                p_start = p_toc

        # 向后寻找该占位符节结束的分节符 </w:sectPr>
        sect_end = doc_bytes.find(b"</w:sectPr>", m_ph)
        if sect_end != -1:
            p_end = doc_bytes.find(b"</w:p>", sect_end)
            if p_end != -1:
                ranges_to_remove.append((p_start, p_end + 6))
        else:
            # 若无分节符，寻找紧随其后的硬换页符段落闭合处
            br_page = doc_bytes.find(b'w:type="page"', m_ph)
            if br_page != -1:
                p_end = doc_bytes.find(b"</w:p>", br_page)
                if p_end != -1:
                    ranges_to_remove.append((p_start, p_end + 6))

    # -------------------------------------------------------------------------
    # 2. 扫描 <w:sdt> 块
    # -------------------------------------------------------------------------
    pos = 0
    sdt_start_tag = b"<w:sdt"
    while True:
        sdt_start = doc_bytes.find(sdt_start_tag, pos)
        if sdt_start == -1:
            break

        next_byte = doc_bytes[sdt_start + 6 : sdt_start + 7]
        if next_byte not in (b">", b" ", b"\t", b"\r", b"\n", b"/"):
            pos = sdt_start + 6
            continue

        depth = 0
        sdt_end = -1
        tag_pattern = re.compile(rb"<\s*(/)?\s*w:sdt(\b[^>]*)>", re.DOTALL)
        for m in tag_pattern.finditer(doc_bytes, sdt_start):
            is_close = bool(m.group(1))
            tag_content = m.group(0)
            if tag_content.endswith(b"/>"):
                continue
            if is_close:
                depth -= 1
                if depth == 0:
                    sdt_end = m.end()
                    break
            else:
                depth += 1

        if sdt_end != -1:
            sdt_content = doc_bytes[sdt_start:sdt_end]
            is_toc_sdt = (
                re.search(rb'docPartGallery\b[^>]*?w:val="Table of Contents"', sdt_content, re.IGNORECASE) is not None
                or re.search(rb'TOC\s+\\o', sdt_content) is not None
            )
            if is_toc_sdt:
                ranges_to_remove.append((sdt_start, sdt_end))
                pos = sdt_end
                continue
            pos = sdt_start + 6
        else:
            pos = sdt_start + 6

    # -------------------------------------------------------------------------
    # 3. 扫描字段段落形式 (Field Paragraphs TOC \o)
    # -------------------------------------------------------------------------
    toc_field_iter = re.finditer(rb'<w:instrText\b[^>]*>\s*TOC\s+\\o', doc_bytes)
    for fm in toc_field_iter:
        instr_pos = fm.start()
        if any(r_start <= instr_pos < r_end for r_start, r_end in ranges_to_remove):
            continue

        p_start = _rfind_p_open(doc_bytes, instr_pos)
        if p_start == -1:
            continue

        fld_pattern = re.compile(rb'<w:fldChar\b[^>]*?w:fldCharType="([^"]+)"')
        fld_depth = 0
        field_end_pos = -1

        for m_fld in fld_pattern.finditer(doc_bytes, p_start):
            ftype = m_fld.group(1)
            if ftype == b"begin":
                fld_depth += 1
            elif ftype == b"end":
                fld_depth -= 1
                if fld_depth <= 0:
                    end_p = doc_bytes.find(b"</w:p>", m_fld.end())
                    field_end_pos = end_p + 6 if end_p != -1 else m_fld.end()
                    break

        if field_end_pos != -1:
            ranges_to_remove.append((p_start, field_end_pos))

    if not ranges_to_remove:
        return doc_bytes

    # 合并重叠或相邻的待删区间
    ranges_to_remove.sort(key=lambda r: r[0])
    merged_ranges: list[list[int]] = []
    for r_start, r_end in ranges_to_remove:
        if not merged_ranges:
            merged_ranges.append([r_start, r_end])
        else:
            prev = merged_ranges[-1]
            if r_start <= prev[1]:
                prev[1] = max(prev[1], r_end)
            else:
                merged_ranges.append([r_start, r_end])

    # -------------------------------------------------------------------------
    # 4. 检查紧邻前面的前导标题段与紧随其后的孤立换页段
    # -------------------------------------------------------------------------
    extended_ranges: list[tuple[int, int]] = []

    for r_start, r_end in merged_ranges:
        new_start = r_start
        new_end = r_end

        # (a) 检查紧邻前面的段落是否为“目录”标题段
        prev_p_end = doc_bytes.rfind(b"</w:p>", 0, r_start)
        if prev_p_end != -1:
            gap = doc_bytes[prev_p_end + 6 : r_start].strip()
            if not gap:
                prev_p_start = _rfind_p_open(doc_bytes, prev_p_end)
                if prev_p_start != -1:
                    prev_p_bytes = doc_bytes[prev_p_start : prev_p_end + 6]
                    text, style_val = _extract_text_and_style_from_p(prev_p_bytes)
                    is_title = (
                        text in _HEADING_TEXT_NORMALIZED
                        or style_val.lower() in _HEADING_STYLES_LOWER
                        or style_val in toc_style_ids
                        or style_val.lower() in toc_style_ids
                    )
                    if is_title:
                        new_start = prev_p_start

        # (b) 检查紧随其后的段落是否为孤立换页符段落
        next_p_start = _find_p_open(doc_bytes, r_end)
        if next_p_start != -1:
            gap = doc_bytes[r_end:next_p_start].strip()
            if not gap:
                next_p_end = doc_bytes.find(b"</w:p>", next_p_start)
                if next_p_end != -1:
                    next_p_bytes = doc_bytes[next_p_start : next_p_end + 6]
                    if _is_isolated_page_break_paragraph(next_p_bytes):
                        new_end = next_p_end + 6

        extended_ranges.append((new_start, new_end))

    # -------------------------------------------------------------------------
    # 5. 执行切片剔除
    # -------------------------------------------------------------------------
    chunks: list[bytes] = []
    curr = 0
    for r_start, r_end in extended_ranges:
        if r_start > curr:
            chunks.append(doc_bytes[curr:r_start])
        curr = max(curr, r_end)
    chunks.append(doc_bytes[curr:])

    return b"".join(chunks)
