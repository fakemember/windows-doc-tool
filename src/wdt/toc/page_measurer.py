"""
实测页码测量与回写器：基于隐藏锚点从导出的临时 PDF 中逐页精确测量各标题的物理页码，
并以原子方式安全回写 DOCX 缓存域，清理锚点与哨兵标记。
"""

from __future__ import annotations

import logging
import re
import zipfile
from pathlib import Path
from typing import Sequence

from pypdf import PdfReader

from wdt.toc.zip_rewriter import rewrite_zip

logger = logging.getLogger(__name__)

# LO 渲染时对页眉/页脚区域的固定预留（默认页脚间距 0.2cm×2），
# 会使其每页正文高度比 Word 少约 226 twips，导致实测页码系统性漂移。
_LO_RESERVED_TWIPS = 226


def build_measurement_copy(docx_path: Path, dst_path: Path) -> None:
    """生成用于实测页码的临时测量副本。

    LibreOffice 把页眉/页脚区域计入正文之外的额外预留，其分页比 Word 更"矮"，
    直接以 LO 渲染测量会得到在 Word 中系统性偏大的页码。副本中剥除全部页眉/页脚
    引用并清零其距离，同时将各节上边距缩减 226 twips，使 LO 的正文高度与 Word
    完全一致——实测页码即 Word 页码。副本仅用于临时 PDF 测量，交付文件不受影响。
    """
    with zipfile.ZipFile(docx_path, "r") as z:
        if "word/document.xml" not in z.namelist():
            raise ValueError("Invalid docx: missing word/document.xml")
        doc_xml = z.read("word/document.xml").decode("utf-8")

    doc_xml = re.sub(r"<w:footerReference\b[^>]*/>", "", doc_xml)
    doc_xml = re.sub(r"<w:headerReference\b[^>]*/>", "", doc_xml)
    doc_xml = re.sub(r'w:footer="[^"]*"', 'w:footer="0"', doc_xml)
    doc_xml = re.sub(r'w:header="[^"]*"', 'w:header="0"', doc_xml)

    def _shrink_top(m: re.Match[str]) -> str:
        tag = m.group(0)
        m_top = re.search(r'w:top="(\d+)"', tag)
        if m_top:
            new_top = max(0, int(m_top.group(1)) - _LO_RESERVED_TWIPS)
            tag = tag[: m_top.start()] + f'w:top="{new_top}"' + tag[m_top.end() :]
        return tag

    doc_xml = re.sub(r"<w:pgMar\b[^>]*/>", _shrink_top, doc_xml)

    with zipfile.ZipFile(docx_path, "r") as zin, zipfile.ZipFile(
        dst_path, "w", zipfile.ZIP_DEFLATED
    ) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = doc_xml.encode("utf-8")
            zout.writestr(item, data)


def build_heading_tag(index: int) -> str:
    """生成标题段落内的隐藏锚点标记名"""
    return f"WDTTAG{index:04d}"


def build_heading_anchor_xml(index: int) -> str:
    """生成嵌入标题段落的 1pt 白字微型隐藏锚点 run XML"""
    tag = build_heading_tag(index)
    return f'<w:r><w:rPr><w:sz w:val="2"/><w:color w:val="FFFFFF"/></w:rPr><w:t>{tag}</w:t></w:r>'


def build_sentinel_xml(sentinel_id: str) -> str:
    """生成正文首页前哨标记段落 XML（1pt 白字）"""
    return (
        f'<w:p>'
        f'<w:pPr><w:rPr><w:sz w:val="2"/><w:color w:val="FFFFFF"/></w:rPr></w:pPr>'
        f'<w:r><w:rPr><w:sz w:val="2"/><w:color w:val="FFFFFF"/></w:rPr>'
        f'<w:t>WDTSENTINEL{sentinel_id}</w:t></w:r>'
        f'</w:p>'
    )


def measure_pages_from_pdf(
    pdf_path: Path,
    heading_count: int,
    sentinel_id: str = "",
    page_offset: int = 0,
) -> list[int] | None:
    """从临时导出的 PDF 中通过隐藏锚点逐页实测标题物理页码。

    :param pdf_path: 临时 PDF 文件路径
    :param heading_count: 标题条目总数
    :param sentinel_id: 正文起始哨兵标记 ID
    :param page_offset: 额外页码偏移量
    :return: 各标题计算后的显示页码列表，若测量失败（任一锚点缺失）返回 None
    """
    if not pdf_path.is_file():
        logger.warning("实测页码失败：PDF 文件不存在: %s", pdf_path)
        return None

    try:
        reader = PdfReader(str(pdf_path))
        total_pdf_pages = len(reader.pages)
    except Exception as e:
        logger.warning("实测页码失败：读取 PDF 异常: %s", e)
        return None

    sentinel_tag = f"WDTSENTINEL{sentinel_id}" if sentinel_id else "WDTSENTINEL"
    sentinel_page: int | None = None
    tag_map: dict[int, int] = {}  # index -> physical 1-based page

    tag_pattern = re.compile(r"WDTTAG(\d{4})")

    for p_idx in range(total_pdf_pages):
        page_num = p_idx + 1
        try:
            text = reader.pages[p_idx].extract_text() or ""
        except Exception:
            text = ""

        # 检查哨兵标记
        if sentinel_page is None and sentinel_tag in text:
            sentinel_page = page_num

        # 匹配标题锚点（取首次出现页）
        for m in tag_pattern.finditer(text):
            idx = int(m.group(1))
            if idx not in tag_map and 0 <= idx < heading_count:
                tag_map[idx] = page_num

        # 性能优化：一旦哨兵与全部标题锚点均已找齐，立即提前退出，避免扫描剩余数千页
        if sentinel_page is not None and len(tag_map) >= heading_count:
            break

    # 校验哨兵与所有标题锚点
    if sentinel_page is None:
        logger.warning("实测页码告警：未在 PDF 中找到哨兵标记 %s，可能受字体或排版裁剪影响", sentinel_tag)
        return None

    k = sentinel_page - 1
    missing_indices = [i for i in range(heading_count) if i not in tag_map]
    if missing_indices:
        logger.warning("实测页码告警：未在 PDF 中找到标题锚点索引: %s", missing_indices)
        return None

    # 计算显示页码：物理页 - k + page_offset（保底 >= 1）
    display_pages: list[int] = []
    for i in range(heading_count):
        raw_page = tag_map[i] - k + page_offset
        display_pages.append(max(1, raw_page))

    return display_pages


def patch_docx_measured_pages(
    docx_path: Path,
    measured_pages: Sequence[int],
    bookmark_names: Sequence[str],
    sentinel_id: str = "",
    seven_zip: Path | None = None,
) -> None:
    """将实测页码严格回写至 docx 的 PAGEREF 缓存，并彻底清理隐藏锚点与哨兵标记。

    :param docx_path: 目标 docx 文件路径
    :param measured_pages: 测得的实测页码列表
    :param bookmark_names: 对应的书签名列表
    :param sentinel_id: 哨兵标记 ID
    :param seven_zip: 7-zip 路径（可选）
    """
    docx_path = Path(docx_path).resolve()
    with zipfile.ZipFile(docx_path, "r") as z:
        if "word/document.xml" not in z.namelist():
            raise ValueError("Invalid docx: missing word/document.xml")
        doc_xml = z.read("word/document.xml").decode("utf-8")

    # 1. 严格按书签名替换目录条目中的页码（支持静态超链接与 PAGEREF 域两种形式）
    search_start = 0
    for bm_name, page_num in zip(bookmark_names, measured_pages):
        anchor_marker = f'w:anchor="{bm_name}"'
        pos = doc_xml.find(anchor_marker, search_start)
        if pos == -1:
            # 尝试从头再找一次
            pos = doc_xml.find(anchor_marker)

        if pos != -1:
            end_pos = doc_xml.find("</w:hyperlink>", pos)
            if end_pos != -1:
                end_pos += len("</w:hyperlink>")
                snippet = doc_xml[pos:end_pos]
                # 匹配静态超链接中 tab 后的纯文本页码 <w:tab/>...<w:t>0</w:t>
                m_tab = re.search(r'(<w:tab/>[\s\S]*?<w:t\b[^>]*>)(\d*)(</w:t>)', snippet)
                if m_tab:
                    new_snippet = (
                        snippet[: m_tab.start(2)] + str(page_num) + snippet[m_tab.end(2) :]
                    )
                    doc_xml = doc_xml[:pos] + new_snippet + doc_xml[end_pos:]
                    search_start = pos + len(new_snippet)
                    continue

                # 匹配超链接内部的 PAGEREF 域缓存
                m_page = re.search(
                    r'(PAGEREF\s+[\s\S]*?<w:fldChar\b[^>]*?w:fldCharType="separate"[\s\S]*?<w:t[^>]*>)(\d+)(</w:t>)',
                    snippet,
                )
                if m_page:
                    new_snippet = (
                        snippet[: m_page.start(2)] + str(page_num) + snippet[m_page.end(2) :]
                    )
                    doc_xml = doc_xml[:pos] + new_snippet + doc_xml[end_pos:]
                    search_start = pos + len(new_snippet)
                    continue

        # 回退全局 PAGEREF 域结构匹配（兼容未包裹在 hyperlink 中的旧形式）
        bm_escaped = re.escape(bm_name)
        pattern = re.compile(
            rf'(PAGEREF\s+{bm_escaped}\b[\s\S]*?<w:fldChar\b[^>]*?w:fldCharType="separate"[\s\S]*?<w:t[^>]*>)(\d+)(</w:t>[\s\S]*?<w:fldChar\b[^>]*?w:fldCharType="end")',
            re.DOTALL,
        )
        doc_xml = pattern.sub(rf"\g<1>{page_num}\g<3>", doc_xml, count=1)

    # 2. 清理正文标题中的隐藏锚点 run
    # 匹配由 build_heading_anchor_xml 精准生成的微型锚点 run（无昂贵 lookahead 回溯）
    anchor_pattern = re.compile(
        r'<w:r\b[^>]*><w:rPr><w:sz w:val="2"/><w:color w:val="FFFFFF"/></w:rPr><w:t>WDTTAG\d{4}</w:t></w:r>'
    )
    doc_xml = anchor_pattern.sub("", doc_xml)

    # 3. 清理哨兵标记段落
    # 匹配包含 WDTSENTINEL 的整段 <w:p>...</w:p>
    sentinel_pattern = re.compile(
        r'<w:p\b[^>]*>(?:(?!</w:p>).)*?WDTSENTINEL[a-zA-Z0-9_-]*(?:(?!</w:p>).)*?</w:p>',
        re.DOTALL,
    )
    doc_xml = sentinel_pattern.sub("", doc_xml)

    # 4. 回写 docx 压缩包
    rewrite_zip(
        docx_path,
        {"word/document.xml": doc_xml.encode("utf-8")},
        seven_zip=seven_zip,
    )
