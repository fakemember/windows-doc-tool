"""
PDF 拆分与独立目录重构模块。

核心功能：
1. 智能解析 PDF 原生大纲 / 书签树（Outlines），提取业务编码与层级关系；
2. 历史目录/占位符页面自动识别与剥离（如旧版 "Table of Contents" 占位节）；
3. 自适应层级聚类装箱算法 (Tree Bin-Packing)：
   - 保持业务大类完整性，大类超标时平滑下钻至子分类，绝不在表格内部断开；
4. 纯 Python 矢量生成合规中文 PDF 目录页（采用 GB1-UCS-2/UniGB-UTF16-UCS-2 标准 Type0/CIDFont，杜绝中文问号与乱码）；
5. 卷内原生大纲书签树重建（剔除非本卷书签，重定向当前卷相对页码）。
"""

from __future__ import annotations

import io
import math
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Sequence

from pypdf import PdfReader, PdfWriter
from pypdf.generic import ArrayObject, Destination, DictionaryObject, FloatObject, NameObject, NumberObject

from wdt.core.toc_pdf_builder import create_toc_pages_batch


@dataclass
class TocNode:
    title: str
    page_index: int  # 0-based PDF 全局物理页码
    level: int = 1  # 1-based 层级 (1=顶级)
    children: list[TocNode] = field(default_factory=list)


@dataclass
class SplitVolume:
    volume_index: int  # 1-based
    title: str
    page_start: int  # 0-based 包含
    page_end: int  # 0-based 包含
    root_node: TocNode
    inner_nodes: list[TocNode] = field(default_factory=list)


def extract_outline_tree(reader: PdfReader) -> list[TocNode]:
    """递归提取 PDF 书签大纲并转为规整的 TocNode 树。"""
    try:
        outline = reader.outline
    except Exception:
        return []

    if not outline:
        return []

    def _parse_item(item, current_level: int) -> TocNode | None:
        if isinstance(item, Destination):
            try:
                page_idx = reader.get_destination_page_number(item)
                if page_idx is not None and page_idx >= 0:
                    return TocNode(title=str(item.title).strip(), page_index=page_idx, level=current_level)
            except Exception:
                pass
        return None

    def _traverse(items, current_level: int) -> list[TocNode]:
        nodes: list[TocNode] = []
        i = 0
        while i < len(items):
            entry = items[i]
            if isinstance(entry, list):
                if nodes:
                    sub_nodes = _traverse(entry, current_level + 1)
                    nodes[-1].children.extend(sub_nodes)
                i += 1
            else:
                node = _parse_item(entry, current_level)
                if node:
                    nodes.append(node)
                i += 1
        return nodes

    return _traverse(outline, 1)


def flatten_nodes(nodes: Sequence[TocNode]) -> list[TocNode]:
    """平铺大纲节点（按遍历顺序）。"""
    res: list[TocNode] = []
    for n in nodes:
        res.append(n)
        if n.children:
            res.extend(flatten_nodes(n.children))
    return res


def _clean_node_title(title: str) -> str:
    """清理并规范化大纲节点标题。"""
    return re.sub(r"[\r\n\t]+", " ", title).strip()


def detect_placeholder_pages(reader: PdfReader, max_check: int = 5) -> set[int]:
    """识别源 PDF 前部废弃的占位符/残留目录页。
    
    典型特征：临床 SAS/RTF 转出的 "Table of Contents" / "From Microsoft Word" 提示文本。
    这类页面应被剔除，避免作为正文第一页或残留于分卷中。
    """
    placeholder_pages = set()
    num_to_check = min(max_check, len(reader.pages))
    
    for idx in range(num_to_check):
        try:
            txt = reader.pages[idx].extract_text() or ""
            norm = re.sub(r"\s+", "", txt).lower()
            if "tableofcontents" in norm and ("frommicrosoftword" in norm or "insert->index" in norm or "indexandcontent" in norm):
                placeholder_pages.add(idx)
        except Exception:
            pass
            
    return placeholder_pages


def plan_volume_splits(
    reader: PdfReader,
    max_pages_per_volume: int = 150,
) -> list[SplitVolume]:
    """自适应大纲聚类装箱算法 (Tree Bin-Packing)。

    核心机制：
    1. 自动探查并跳过前部的残留/废弃目录占位符页；
    2. 提取所有大纲节点并分析层级关系；
    3. 临床 TFL 场景智能适配：若所有节点均处于 Level 1，按 14.1 / 14.2 等业务编号前缀推导一级分类；
    4. 自适应决策：大类未超标则整类聚合；超标则平滑在次级大纲（如 14.2.1, 14.2.9）边界切分，
       绝不从表格正文/续表中途截断；
    5. 兜底逻辑：无大纲时按 max_pages_per_volume 机械均分。
    """
    total_pages = len(reader.pages)
    if total_pages == 0:
        return []

    placeholders = detect_placeholder_pages(reader)
    first_valid_page = 0
    while first_valid_page in placeholders and first_valid_page < total_pages:
        first_valid_page += 1

    outline_tree = extract_outline_tree(reader)
    if not outline_tree:
        chunk_size = max_pages_per_volume
        content_pages = [p for p in range(total_pages) if p not in placeholders]
        if not content_pages:
            content_pages = list(range(total_pages))
        num_chunks = math.ceil(len(content_pages) / chunk_size)
        volumes = []
        for i in range(num_chunks):
            chunk = content_pages[i * chunk_size : (i + 1) * chunk_size]
            dummy = TocNode(title=f"Section {i + 1}", page_index=chunk[0], level=1)
            volumes.append(
                SplitVolume(
                    volume_index=i + 1,
                    title=dummy.title,
                    page_start=chunk[0],
                    page_end=chunk[-1],
                    root_node=dummy,
                    inner_nodes=[dummy],
                )
            )
        return volumes

    # 分析各节点前缀
    parsed_items = []
    for n in outline_tree:
        clean_t = _clean_node_title(n.title)
        m = re.search(r"(\d+(?:\.\d+)*)", clean_t)
        code = m.group(1) if m else ""
        parts = code.split(".") if code else []
        l1 = ".".join(parts[:2]) if len(parts) >= 2 else (parts[0] if parts else "")
        parsed_items.append({"node": n, "code": code, "l1": l1})

    unique_l1s = set(p["l1"] for p in parsed_items if p["l1"])
    if len(unique_l1s) >= 2 and all(n.level == 1 for n in outline_tree):
        seen_l1 = set()
        top_nodes = []
        for p in parsed_items:
            if p["l1"] and p["l1"] not in seen_l1:
                seen_l1.add(p["l1"])
                top_nodes.append(p["node"])
    else:
        top_nodes = [node for node in outline_tree if node.level == 1]
        if not top_nodes:
            flat = flatten_nodes(outline_tree)
            min_lvl = min(n.level for n in flat)
            top_nodes = [n for n in flat if n.level == min_lvl]

    valid_tops: list[TocNode] = []
    for n in top_nodes:
        if 0 <= n.page_index < total_pages:
            if not valid_tops or n.page_index > valid_tops[-1].page_index:
                valid_tops.append(n)

    has_oversized_top = False
    for idx, top in enumerate(valid_tops):
        p_start = first_valid_page if idx == 0 else top.page_index
        p_end = (valid_tops[idx + 1].page_index - 1) if (idx + 1 < len(valid_tops)) else (total_pages - 1)
        if (p_end - p_start + 1) > max_pages_per_volume:
            has_oversized_top = True
            break

    # 模式 A: 顶级章节皆适中，直接顶级切分
    if not has_oversized_top and len(valid_tops) > 1:
        volumes = []
        for idx, top in enumerate(valid_tops):
            vol_idx = idx + 1
            p_start = first_valid_page if idx == 0 else top.page_index
            p_end = (valid_tops[idx + 1].page_index - 1) if (idx + 1 < len(valid_tops)) else (total_pages - 1)
            all_inner = [n for n in flatten_nodes([top]) if p_start <= n.page_index <= p_end]
            if not all_inner:
                all_inner = [top]
            clean_title = _clean_node_title(top.title)
            volumes.append(
                SplitVolume(
                    volume_index=vol_idx,
                    title=clean_title or f"Part {vol_idx}",
                    page_start=p_start,
                    page_end=p_end,
                    root_node=top,
                    inner_nodes=all_inner,
                )
            )
        return volumes

    # 模式 B: 自适应下钻聚类装箱
    flat_nodes = flatten_nodes(outline_tree)
    valid_nodes: list[TocNode] = []
    for n in flat_nodes:
        if 0 <= n.page_index < total_pages:
            if not valid_nodes or n.page_index >= valid_nodes[-1].page_index:
                valid_nodes.append(n)

    units: list[dict] = []
    for idx, node in enumerate(valid_nodes):
        p_start = first_valid_page if idx == 0 else node.page_index
        p_end = (valid_nodes[idx + 1].page_index - 1) if (idx + 1 < len(valid_nodes)) else (total_pages - 1)
        if p_end < p_start:
            p_end = p_start

        m = re.search(r"(\d+(?:\.\d+)*)", node.title)
        num_code = m.group(1) if m else ""
        parts = num_code.split(".") if num_code else []
        l1_prefix = ".".join(parts[:2]) if len(parts) >= 2 else (parts[0] if parts else "Section")
        l2_prefix = ".".join(parts[:3]) if len(parts) >= 3 else l1_prefix

        units.append({
            "node": node,
            "p_start": p_start,
            "p_end": p_end,
            "pages": p_end - p_start + 1,
            "l1": l1_prefix,
            "l2": l2_prefix,
        })

    volumes = []
    current_units = []
    current_pages = 0

    def _flush_volume():
        nonlocal current_units, current_pages
        if not current_units:
            return
        vol_idx = len(volumes) + 1
        vol_start = current_units[0]["p_start"]
        vol_end = current_units[-1]["p_end"]
        
        u_l1s = list(dict.fromkeys(u["l1"] for u in current_units))
        if len(u_l1s) == 1:
            vol_title = f"{u_l1s[0]} ({_clean_node_title(current_units[0]['node'].title)})"
        else:
            vol_title = f"{u_l1s[0]} - {u_l1s[-1]}"

        inner = [u["node"] for u in current_units]
        volumes.append(
            SplitVolume(
                volume_index=vol_idx,
                title=vol_title,
                page_start=vol_start,
                page_end=vol_end,
                root_node=current_units[0]["node"],
                inner_nodes=inner,
            )
        )
        current_units = []
        current_pages = 0

    for idx, u in enumerate(units):
        u_pages = u["pages"]
        if current_units:
            prev_u = current_units[-1]
            same_l1 = (u["l1"] == prev_u["l1"])
            same_l2 = (u["l2"] == prev_u["l2"])

            overflow = (current_pages + u_pages > max_pages_per_volume)
            l1_boundary_cut = (not same_l1 and current_pages >= 30)
            
            if (overflow and not same_l2) or l1_boundary_cut or (overflow and current_pages >= max_pages_per_volume * 0.8):
                _flush_volume()

        current_units.append(u)
        current_pages += u_pages

    _flush_volume()
    return volumes


def _text_to_pdf_hex(text: str) -> str:
    """将包含中英文的字符串转为 PDF Hex 编码 (<feff...>)，适配 Type0 / UCS-2 CIDFont。"""
    return text.encode("utf-16-be").hex()


def generate_toc_pdf_page(
    volume: SplitVolume,
    all_volumes: list[SplitVolume],
    page_width: float = 595.276,  # A4 width in pt
    page_height: float = 841.890,  # A4 height in pt
) -> bytes:
    """纯 Python 生成支持合规中文显示的矢量 PDF 独立目录页。
    
    内置标准 CIDFontType0 + UniGB-UTF16-UCS-2 字体描述符，
    支持在任何标准 PDF 查看器（Acrobat, Edge, Chrome, Preview）中完美渲染中文字符。
    """
    total_vols = len(all_volumes)
    vol_title = f"Volume {volume.volume_index} of {total_vols}: {volume.title}"
    if len(vol_title) > 60:
        vol_title = vol_title[:57] + "..."

    items_to_draw: list[tuple[str, int, int]] = []
    seen_prefixes = set()
    for node in volume.inner_nodes:
        # 正文页码从 1 起始
        rel_page = (node.page_index - volume.page_start) + 1
        clean_name = _clean_node_title(node.title)
        
        m = re.search(r"(\d+(?:\.\d+)*)", clean_name)
        code = m.group(1) if m else clean_name[:15]
        if len(volume.inner_nodes) > 25:
            prefix_key = ".".join(code.split(".")[:3])
            if prefix_key in seen_prefixes and len(items_to_draw) >= 20:
                continue
            seen_prefixes.add(prefix_key)

        items_to_draw.append((clean_name, rel_page, node.level))

    content_lines: list[str] = []

    # 1. 顶部居中精简大标题：目  录
    title_text = "目    录"
    content_lines.append("BT")
    content_lines.append("/F_CN 16 Tf")
    content_lines.append("0.12 0.25 0.45 rg")
    # A4 居中 (270 pt 附近)
    content_lines.append(f"{(page_width / 2.0) - 35:.2f} {page_height - 65:.2f} Td")
    content_lines.append(f"<{_text_to_pdf_hex(title_text)}> Tj")
    content_lines.append("ET")

    content_lines.append("0.8 0.85 0.9 RG 0.75 w")
    content_lines.append(f"40 {page_height - 78:.2f} m {page_width - 40:.2f} {page_height - 78:.2f} l S")

    # 2. 渲染目录条目
    y = page_height - 105
    max_items = 25
    line_height = 22.0
    right_margin = page_width - 40

    for idx, (title, rel_p, lvl) in enumerate(items_to_draw[:max_items]):
        indent = 40 + (min(3, lvl) - 1) * 14
        p_str = str(rel_p)

        color = "0.1 0.1 0.1 rg" if lvl == 1 else "0.25 0.25 0.25 rg"
        font_size = "9.5" if lvl == 1 else "8.5"

        # 标题长度安全截断
        max_title_len = 50 - (lvl - 1) * 4
        display_title = title if len(title) <= max_title_len else title[: max_title_len - 3] + "..."

        # 条目标题文本
        content_lines.append("BT")
        content_lines.append(f"/F_CN {font_size} Tf")
        content_lines.append(color)
        content_lines.append(f"{indent:.2f} {y:.2f} Td")
        content_lines.append(f"<{_text_to_pdf_hex(display_title)}> Tj")
        content_lines.append("ET")

        # 页码 (靠右对齐)
        page_indent = right_margin - (len(p_str) * 7.5)
        content_lines.append("BT")
        content_lines.append(f"/F_CN {font_size} Tf")
        content_lines.append(color)
        content_lines.append(f"{page_indent:.2f} {y:.2f} Td")
        content_lines.append(f"<{_text_to_pdf_hex(p_str)}> Tj")
        content_lines.append("ET")

        # 点导线 (Dot Leaders)
        # 中英文字符宽度预估：ASCII 约 5.5pt，中文约 9pt
        est_len = sum(9.0 if ord(c) > 127 else 5.2 for c in display_title)
        dot_start = indent + est_len + 8
        dot_end = page_indent - 8
        if dot_end > dot_start + 18:
            num_dots = int((dot_end - dot_start) / 4.8)
            dots = ". " * (num_dots // 2)
            content_lines.append("BT")
            content_lines.append("/F_CN 8 Tf")
            content_lines.append("0.6 0.6 0.6 rg")
            content_lines.append(f"{dot_start:.2f} {y:.2f} Td")
            content_lines.append(f"<{_text_to_pdf_hex(dots)}> Tj")
            content_lines.append("ET")

        y -= line_height

    if len(volume.inner_nodes) > len(items_to_draw[:max_items]):
        remaining = len(volume.inner_nodes) - len(items_to_draw[:max_items])
        omitted_text = f"[+ 尚有 {remaining} 个子表未在概览页展开...]"
        content_lines.append("BT")
        content_lines.append("/F_CN 8 Tf")
        content_lines.append("0.5 0.5 0.5 rg")
        content_lines.append(f"40 {y - 4:.2f} Td")
        content_lines.append(f"<{_text_to_pdf_hex(omitted_text)}> Tj")
        content_lines.append("ET")

    footer_text = f"Part {volume.volume_index} | 原文档对应页码 P.{volume.page_start + 1} - P.{volume.page_end + 1}"
    content_lines.append("BT")
    content_lines.append("/F_CN 8 Tf")
    content_lines.append("0.55 0.55 0.55 rg")
    content_lines.append(f"40 32 Td")
    content_lines.append(f"<{_text_to_pdf_hex(footer_text)}> Tj")
    content_lines.append("ET")
    content_lines.append("Q")

    stream_content = "\n".join(content_lines).encode("latin-1")

    # 构建标准的 Type0 CIDFont 中文字体 PDF 页面（兼容 STSong-Light / SimSun）
    mediabox_str = f"/MediaBox [0 0 {page_width:.2f} {page_height:.2f}] ".encode("ascii")
    length_str = f"<< /Length {len(stream_content)} >>\n".encode("ascii")
    pdf_obj = (
        b"%PDF-1.4\n"
        b"1 0 obj\n"
        b"<< /Type /Catalog /Pages 2 0 R >>\n"
        b"endobj\n"
        b"2 0 obj\n"
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>\n"
        b"endobj\n"
        b"3 0 obj\n"
        b"<< /Type /Page /Parent 2 0 R "
        + mediabox_str
        + b"/Contents 4 0 R /Resources << /Font << "
        b"/F_CN << "
        b"/Type /Font /Subtype /Type0 "
        b"/BaseFont /STSong-Light "
        b"/Encoding /UniGB-UTF16-UCS-2 "
        b"/DescendantFonts [5 0 R] "
        b">> >> >> >>\n"
        b"endobj\n"
        b"5 0 obj\n"
        b"<< /Type /Font /Subtype /CIDFontType0 /BaseFont /STSong-Light "
        b"/CIDSystemInfo << /Registry (Adobe) /Ordering (GB1) /Supplement 4 >> "
        b">> endobj\n"
        b"4 0 obj\n"
        + length_str
        + b"stream\n"
        + stream_content
        + b"\nendstream\n"
        b"endobj\n"
        b"xref\n"
        b"0 6\n"
        b"0000000000 65535 f \n"
        b"0000000009 00000 n \n"
        b"0000000058 00000 n \n"
        b"0000000115 00000 n \n"
        b"0000000450 00000 n \n"
        b"0000000300 00000 n \n"
        b"trailer\n"
        b"<< /Size 6 /Root 1 0 R >>\n"
        b"startxref\n"
        b"550\n"
        b"%%EOF\n"
    )

    tmp_reader = PdfReader(io.BytesIO(pdf_obj))
    tmp_writer = PdfWriter()
    tmp_writer.add_page(tmp_reader.pages[0])
    out_buf = io.BytesIO()
    tmp_writer.write(out_buf)
    return out_buf.getvalue()


@dataclass
class SplitResult:
    volume_index: int
    title: str
    output_path: Path
    page_count: int
    original_pages: tuple[int, int]
    file_size_bytes: int


def split_pdf_by_outline(
    src_pdf: Path,
    out_dir: Path | None = None,
    with_visual_toc: bool = True,
    max_pages: int = 150,
) -> list[SplitResult]:
    """主函数：按大纲与智能层级聚类装箱将超大 PDF 拆分为多个独立分卷。"""
    src_pdf = Path(src_pdf).resolve()
    if not src_pdf.exists():
        raise FileNotFoundError(f"PDF 文件不存在: {src_pdf}")

    if out_dir is None:
        target_dir = src_pdf.parent / f"{src_pdf.stem}_volumes"
    else:
        target_dir = Path(out_dir).resolve()

    target_dir.mkdir(parents=True, exist_ok=True)

    reader = PdfReader(str(src_pdf))
    total_pages = len(reader.pages)
    if total_pages == 0:
        return []

    # 自适应聚类规划分卷
    volumes = plan_volume_splits(reader, max_pages_per_volume=max_pages)
    if not volumes:
        return []

    placeholders = detect_placeholder_pages(reader)

    # 预先批量构建各分卷的高保真 Word COM 中文目录页
    toc_pdf_bytes_list: list[bytes] = []
    if with_visual_toc:
        toc_specs = []
        for vol in volumes:
            items_to_draw = []
            seen_prefixes = set()
            for node in vol.inner_nodes:
                rel_p = (node.page_index - vol.page_start) + 1
                clean_name = _clean_node_title(node.title)
                m = re.search(r"(\d+(?:\.\d+)*)", clean_name)
                code = m.group(1) if m else clean_name[:15]
                if len(vol.inner_nodes) > 25:
                    prefix_key = ".".join(code.split(".")[:3])
                    if prefix_key in seen_prefixes and len(items_to_draw) >= 20:
                        continue
                    seen_prefixes.add(prefix_key)
                items_to_draw.append((clean_name, rel_p, node.level))

            overview_text = f"MULTI-VOLUME COMPILATION: {len(volumes)} VOLUMES TOTAL"
            v_title = f"Volume {vol.volume_index} of {len(volumes)}: {vol.title}"
            footer_txt = f"Part {vol.volume_index} | 原文档对应页码 P.{vol.page_start + 1} - P.{vol.page_end + 1}"
            toc_specs.append({
                "vol_title": v_title,
                "overview_text": overview_text,
                "items": items_to_draw,
                "footer_text": footer_txt,
            })
        toc_pdf_bytes_list = create_toc_pages_batch(toc_specs)

    results: list[SplitResult] = []
    digits = len(str(len(volumes)))

    for idx, vol in enumerate(volumes):
        writer = PdfWriter()
        toc_page_offset = 0

        # 获取首个目标页的尺寸以适配目录页
        first_page = reader.pages[vol.page_start]
        page_box = first_page.mediabox
        p_width = float(page_box.width)
        p_height = float(page_box.height)

        if with_visual_toc:
            toc_bytes = toc_pdf_bytes_list[idx] if (idx < len(toc_pdf_bytes_list) and toc_pdf_bytes_list[idx]) else None
            if not toc_bytes:
                toc_bytes = generate_toc_pdf_page(vol, volumes, page_width=p_width, page_height=p_height)

            toc_reader = PdfReader(io.BytesIO(toc_bytes))
            writer.add_page(toc_reader.pages[0])
            toc_page_offset = 1

            writer.add_outline_item(
                title=f"Table of Contents ({vol.title})",
                page_number=0,
            )

        # 添加该卷正文页面（自动跳过残留的占位符页）
        valid_written_pages = []
        for page_idx in range(vol.page_start, vol.page_end + 1):
            if page_idx in placeholders:
                continue
            writer.add_page(reader.pages[page_idx])
            valid_written_pages.append(page_idx)

        # 写入落入该卷的所有书签（相对物理页码校准）
        page_index_map = {orig_p: (i + toc_page_offset) for i, orig_p in enumerate(valid_written_pages)}
        for node in vol.inner_nodes:
            if node.page_index in page_index_map:
                writer.add_outline_item(
                    title=node.title,
                    page_number=page_index_map[node.page_index],
                )

        # 为第 0 页可视目录页添加真实点击交互超链接 (Link Annotations)
        if with_visual_toc and toc_page_offset == 1:
            drawn_items = toc_specs[idx]["items"][:24] if toc_specs else []
            # Word 渲染起始 y 坐标约为 717.1 pt，逐行递减 27.0 pt；内置 Python fallback 渲染起始为 height - 105，行高 22 pt
            used_word = bool(toc_pdf_bytes_list and idx < len(toc_pdf_bytes_list) and toc_pdf_bytes_list[idx])
            start_y = 717.1 if used_word else (p_height - 105.0)
            line_step = 27.0 if used_word else 22.0

            for item_idx, (it_title, rel_p, lvl) in enumerate(drawn_items):
                # rel_p 为 1-based 的正文页码，对应的物理页为 target_page_idx = rel_p
                target_page_idx = rel_p
                if target_page_idx < len(writer.pages):
                    target_page_ref = writer.pages[target_page_idx].indirect_reference
                    center_y = start_y - (item_idx * line_step)
                    y_bottom = center_y - 7.0
                    y_top = center_y + 13.0

                    link_annot = DictionaryObject({
                        NameObject("/Type"): NameObject("/Annot"),
                        NameObject("/Subtype"): NameObject("/Link"),
                        NameObject("/Rect"): ArrayObject([
                            FloatObject(40.0),
                            FloatObject(y_bottom),
                            FloatObject(p_width - 40.0),
                            FloatObject(y_top),
                        ]),
                        NameObject("/Border"): ArrayObject([
                            NumberObject(0),
                            NumberObject(0),
                            NumberObject(0),
                        ]),
                        NameObject("/Dest"): ArrayObject([
                            target_page_ref,
                            NameObject("/Fit"),
                        ]),
                    })
                    writer.add_annotation(page_number=0, annotation=link_annot)

        safe_title = re.sub(r'[\\/*?:"<>|]', "_", vol.title).strip()
        safe_title = safe_title[:40].strip()
        out_name = f"{src_pdf.stem}_Part{vol.volume_index:0{digits}d}_{safe_title}.pdf"
        dst_path = target_dir / out_name

        with open(dst_path, "wb") as f:
            writer.write(f)

        actual_page_count = len(valid_written_pages) + toc_page_offset
        results.append(
            SplitResult(
                volume_index=vol.volume_index,
                title=vol.title,
                output_path=dst_path,
                page_count=actual_page_count,
                original_pages=(vol.page_start + 1, vol.page_end + 1),
                file_size_bytes=dst_path.stat().st_size,
            )
        )

    return results
