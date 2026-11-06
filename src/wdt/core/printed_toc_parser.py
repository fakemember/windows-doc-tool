"""打印版目录（Printed Table of Content）解析模块。

部分上游工具（如 SAS ClinStReport）导出的 PDF 不携带任何书签大纲（/Outlines），
目录仅以打印文字形式出现在文首若干页。本模块从这些页面解析
"清单16.2.1.1 筛选失败受试者清单 ...... 3" 形式的条目，重建规整的
TocNode 列表，供 PDF 拆分聚类使用，并识别目录页本身以便从分卷中剔除。
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

from wdt.core.pdf_splitter import TocNode

if TYPE_CHECKING:
    from pypdf import PdfReader

# 目录条目行：标题 ....页码（点导线 + 行尾页码）
_ENTRY_RE = re.compile(r"^(?P<title>.+?)\s*\.{3,}\s*(?P<page>\d{1,6})\s*$")
# 条目标题中的业务编号（如 16.2.1.1）
_CODE_RE = re.compile(r"(\d+(?:\.\d+)*)")
# 页眉/页脚噪声：页码条、SAS 输出头、日期时间戳
_NOISE_RE = re.compile(
    r"page\s*\d+\s*of\s*\d+|clinstreport|\d{2}[a-z]{3}\d{4}[: ]\d{2}:\d{2}", re.IGNORECASE
)
_TOC_MARKER_RE = re.compile(r"table\s*of\s+contents?", re.IGNORECASE)

# 至少解析出该数量的条目才认为页面是打印版目录（避免误判正文页）
_MIN_TOC_ENTRIES = 3
# 至多扫描文首页数
_MAX_SCAN_PAGES = 10


def _iter_entry_lines(text: str) -> list[tuple[str, int]]:
    """从页面文本中提取 (标题, 页码) 条目，剔除页眉页脚噪声。"""
    entries: list[tuple[str, int]] = []
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line or _NOISE_RE.search(line):
            continue
        m = _ENTRY_RE.match(line)
        if m is None:
            continue
        title = re.sub(r"\s+", " ", m.group("title")).strip(" .")
        if not title:
            continue
        entries.append((title, int(m.group("page"))))
    return entries


def extract_printed_toc(
    reader: "PdfReader", max_scan_pages: int = _MAX_SCAN_PAGES
) -> tuple[list[TocNode], set[int]]:
    """从文首打印版目录页重建大纲。

    返回 (扁平 TocNode 列表, 目录页物理页码集合)。
    无有效目录时返回 ([], set())。
    """
    total_pages = len(reader.pages)
    if total_pages == 0:
        return [], set()

    toc_pages: set[int] = set()
    raw_entries: list[tuple[str, int]] = []
    for idx in range(min(max_scan_pages, total_pages)):
        try:
            text = reader.pages[idx].extract_text() or ""
        except Exception:
            break
        # 目录页判定：含 Table of Content 标题，或该行式条目足够多
        is_marker_page = bool(_TOC_MARKER_RE.search(text))
        entries = _iter_entry_lines(text)
        if is_marker_page:
            toc_pages.add(idx)
            raw_entries.extend(entries)
            continue
        # 无标题但条目很多的页：仅当其紧邻已确认目录页时才并入（多页目录的后续页）
        if entries and toc_pages and idx == max(toc_pages) + 1 and len(entries) >= _MIN_TOC_ENTRIES:
            toc_pages.add(idx)
            raw_entries.extend(entries)
            continue
        # 既无标记也无条目（或目录已结束）：停止扫描
        break

    if len(raw_entries) < _MIN_TOC_ENTRIES:
        return [], set()

    # 过滤并校准页码：1-based 打印页码 -> 0-based 物理页码，要求严格递增
    nodes: list[TocNode] = []
    last_page = 0
    for title, printed_page in raw_entries:
        page_index = printed_page - 1
        if not (0 <= page_index < total_pages):
            continue
        if nodes and page_index <= last_page:
            continue
        m = _CODE_RE.search(title)
        level = m.group(1).count(".") + 1 if m else 1
        nodes.append(TocNode(title=title, page_index=page_index, level=level))
        last_page = page_index

    if len(nodes) < _MIN_TOC_ENTRIES:
        return [], set()

    # 层级归一化：最小层级提升为 1
    min_level = min(n.level for n in nodes)
    if min_level > 1:
        for n in nodes:
            n.level -= min_level - 1

    return nodes, toc_pages
