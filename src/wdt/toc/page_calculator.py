"""
页码计算器：基于 bisect 二分查找高效计算各标题在 Word 中的物理页码。
支持分节重启（STATIC_RESTART）、页数偏移（STATIC_OFFSET）、原位替换（STATIC_IN_PLACE）与动态域（DYNAMIC）四种模式。
"""

from __future__ import annotations

import bisect
from enum import Enum
from typing import Sequence


class PageMode(Enum):
    STATIC_RESTART = "static-restart"      # 分节页码重启（默认）
    STATIC_OFFSET = "static-offset"        # 估算目录页数偏移补偿
    STATIC_IN_PLACE = "static-in-place"    # 仅 REPLACE_PLACEHOLDER 用，移植基座公式
    DYNAMIC = "dynamic"                    # 动态域，页码由 Word 打开时刷新


def find_all_bytes(sub: bytes, b: bytes) -> list[int]:
    """快速定位字节串中所有子串的起始位置"""
    res: list[int] = []
    pos = 0
    while True:
        pos = b.find(sub, pos)
        if pos == -1:
            break
        res.append(pos)
        pos += len(sub)
    return res


def compute_pages(
    doc_bytes: bytes,
    heading_positions: Sequence[int],
    insert_pos: int,
    mode: PageMode,
    page_offset: int = 0,
) -> list[int]:
    """计算各标题对应的物理页码。

    :param doc_bytes: document.xml 的原始字节流
    :param heading_positions: 各标题段落的起始字节偏移列表
    :param insert_pos: 目录插入位置（字节偏移）
    :param mode: 页码计算模式
    :param page_offset: 页码偏移增量（仅用于 STATIC_OFFSET 模式）
    :return: 对应的页码列表（DYNAMIC 模式返回空列表）
    """
    if mode == PageMode.DYNAMIC:
        return []

    # 扫描全文中所有的硬分页符与分节符起始位置
    br_positions = find_all_bytes(b'<w:br w:type="page"/>', doc_bytes)
    sect_positions = find_all_bytes(b'<w:sectPr', doc_bytes)

    pages: list[int] = []

    if mode == PageMode.STATIC_IN_PLACE:
        # 移植基座原公式：
        # 占位符原位替换模式下，目录物理占用占位符所在区域，页码由全文字节流统计：
        # page = 1 + 2 + br_count + (sect_count - 1)
        for pos in heading_positions:
            br_count = bisect.bisect_left(br_positions, pos)
            sect_count = bisect.bisect_left(sect_positions, pos)
            page_num = 1 + 2 + br_count + (sect_count - 1)
            pages.append(page_num)
        return pages

    # STATIC_RESTART 与 STATIC_OFFSET 模式：
    # 核心公式推导：
    # 1. 目录插入点 insert_pos 处会注入分节符（目录节结），正文节从此处开启并将页码重启为 1 (pgNumType start=1)。
    # 2. insert_pos 之前的所有分页符与分节符属于封面/目录节，完全不计入正文页码。
    # 3. 统计严格落在 (insert_pos, pos) 区间内的硬分页符数量 br_after 与分节符数量 sect_after。
    #    正文首个分节之后的每个分节均带来一次页面换节，故正文内每个 sectPr 产生一次新页。
    # 4. 正文首段页码基础值为 1，故 page = 1 + br_after + sect_after。
    # 5. OFFSET 模式再追加 page_offset（补偿目录实际物理占页数）。
    br_insert_idx = bisect.bisect_right(br_positions, insert_pos)
    sect_insert_idx = bisect.bisect_right(sect_positions, insert_pos)

    for pos in heading_positions:
        if pos <= insert_pos:
            # 防御性处理：位于目录插入点之前的异常标题直接赋为第 1 页
            pages.append(1)
            continue

        br_pos_idx = bisect.bisect_left(br_positions, pos)
        sect_pos_idx = bisect.bisect_left(sect_positions, pos)

        br_count_after = max(0, br_pos_idx - br_insert_idx)
        sect_count_after = max(0, sect_pos_idx - sect_insert_idx)

        base_page = 1 + br_count_after + sect_count_after
        if mode == PageMode.STATIC_OFFSET:
            base_page += page_offset

        pages.append(base_page)

    return pages
