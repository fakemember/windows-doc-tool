"""
DOCX 字节流注入器：TOC 目录注入引擎主入口。
纯 Python 字节流切片、书签全局避让、分节页码重启、4 样式注入及原子写回。
"""

from __future__ import annotations

import html
import re
import shutil
import time
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
import zipfile

from wdt.toc.heading_detector import Heading, detect_headings
from wdt.toc.page_calculator import PageMode, compute_pages
from wdt.toc.styles_templates import HYPERLINK_STYLE_ID, TocStyle, toc_styles_fragment
from wdt.toc.zip_rewriter import rewrite_zip


class Position(Enum):
    DOCUMENT_START = "document-start"
    BEFORE_FIRST_HEADING = "before-first-heading"
    REPLACE_PLACEHOLDER = "replace-placeholder"   # 移植基座占位符区域替换
    BOOKMARK = "bookmark"                          # 需配 bookmark_name


@dataclass
class TocConfig:
    style: TocStyle = TocStyle.CLASSIC_DOT
    depth: int = 1                    # 1..3，超出忽略
    title: str = "Table of Contents"
    position: Position = Position.DOCUMENT_START
    bookmark_name: str | None = None
    page_mode: PageMode = PageMode.STATIC_RESTART
    page_offset: int = 0
    backup: bool = True


@dataclass
class TocReport:
    entries: int
    pages: list[int]
    backup_path: Path | None
    elapsed: float


class NoHeadingError(ValueError):
    """未在文档中检测到符合层级条件的标题"""
    pass


def _generate_bookmark_prefix(doc_bytes: bytes) -> str:
    """生成不与文档既有书签前缀冲突的新前缀"""
    existing_names = re.findall(rb'<w:bookmarkStart\b[^>]*?w:name="([^"]+)"', doc_bytes)
    names_str = {n.decode("utf-8", errors="replace") for n in existing_names}

    prefix = "_WDToc"
    while any(name.startswith(prefix) for name in names_str):
        prefix += "X"
    return prefix


def _find_max_bookmark_id(doc_bytes: bytes) -> int:
    """获取文档现有书签的最大 id"""
    existing_ids = [
        int(m)
        for m in re.findall(rb'<w:bookmarkStart\b[^>]*?w:id="(\d+)"', doc_bytes)
    ]
    return max(existing_ids) if existing_ids else -1


def _build_toc_entry_xml(
    index: int,
    total: int,
    heading: Heading,
    bm_name: str,
    page_num: int | None,
    style: TocStyle,
    depth: int,
) -> str:
    """生成单条 TOC 目录项段落 XML"""
    level = heading.level
    escaped_title = html.escape(heading.text)

    # 导线设置
    leader = "none" if style == TocStyle.MINIMAL_CLEAN else "dot"
    tab_pr = f'<w:tabs><w:tab w:val="right" w:leader="{leader}" w:pos="9305"/></w:tabs>'

    # 标题字体与颜色
    if style == TocStyle.MODERN_ACCENT:
        if level == 1:
            title_rpr = (
                '<w:rPr>'
                f'<w:rStyle w:val="{HYPERLINK_STYLE_ID}"/>'
                '<w:rFonts w:ascii="Calibri" w:eastAsia="微软雅黑" w:hAnsi="Calibri"/>'
                '<w:b/><w:color w:val="1F4E79"/><w:noProof/>'
                '</w:rPr>'
            )
        else:
            title_rpr = (
                '<w:rPr>'
                f'<w:rStyle w:val="{HYPERLINK_STYLE_ID}"/>'
                '<w:rFonts w:ascii="Calibri" w:eastAsia="微软雅黑" w:hAnsi="Calibri"/>'
                '<w:color w:val="262626"/><w:noProof/>'
                '</w:rPr>'
            )
    elif style == TocStyle.REGULATORY_LISTING:
        title_rpr = (
            '<w:rPr>'
            f'<w:rStyle w:val="{HYPERLINK_STYLE_ID}"/>'
            '<w:rFonts w:ascii="Times New Roman" w:eastAsia="等线" w:hAnsi="Times New Roman"/>'
            '<w:noProof/>'
            '</w:rPr>'
        )
    elif style == TocStyle.CLASSIC_DOT:
        title_rpr = (
            '<w:rPr>'
            f'<w:rStyle w:val="{HYPERLINK_STYLE_ID}"/>'
            '<w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体" w:hAnsi="Times New Roman"/>'
            '<w:noProof/>'
            '</w:rPr>'
        )
    else:  # MINIMAL_CLEAN
        title_rpr = (
            '<w:rPr>'
            f'<w:rStyle w:val="{HYPERLINK_STYLE_ID}"/>'
            '<w:rFonts w:ascii="Arial" w:eastAsia="微软雅黑" w:hAnsi="Arial"/>'
            '<w:color w:val="333333"/><w:noProof/>'
            '</w:rPr>'
        )

    # 页码字体样式（极简模式弱化）
    if style == TocStyle.MINIMAL_CLEAN:
        page_rpr = '<w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/><w:color w:val="7F7F7F"/><w:noProof/></w:rPr>'
    else:
        page_rpr = '<w:rPr><w:noProof/></w:rPr>'

    page_str = str(page_num) if page_num is not None else ""

    hyperlink_block = (
        f'<w:hyperlink w:anchor="{bm_name}" w:history="1">'
        f'<w:r>{title_rpr}<w:t xml:space="preserve">{escaped_title}</w:t></w:r>'
        f'<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:tab/></w:r>'
        f'<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>'
        f'<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:instrText xml:space="preserve"> PAGEREF {bm_name} </w:instrText></w:r>'
        f'<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:instrText>\\</w:instrText></w:r>'
        f'<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:instrText xml:space="preserve">h </w:instrText></w:r>'
        f'<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>'
        f'<w:r>{page_rpr}<w:t>{page_str}</w:t></w:r>'
        f'<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>'
        f'</w:hyperlink>'
    )

    if index == 0:
        # 首个条目包装 TOC 域开始标记
        p_xml = (
            f'<w:p>'
            f'<w:pPr><w:pStyle w:val="TOC{level}"/>{tab_pr}<w:rPr><w:noProof/></w:rPr></w:pPr>'
            f'<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            f'<w:r><w:instrText xml:space="preserve"> TOC \\o "1-{depth}" \\h \\z \\u </w:instrText></w:r>'
            f'<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            f'{hyperlink_block}'
            f'</w:p>'
        )
    else:
        p_xml = (
            f'<w:p>'
            f'<w:pPr><w:pStyle w:val="TOC{level}"/>{tab_pr}<w:rPr><w:noProof/></w:rPr></w:pPr>'
            f'{hyperlink_block}'
            f'</w:p>'
        )

    return p_xml


def _build_dynamic_toc_xml(depth: int) -> str:
    """生成动态字段模式下的 TOC 域单段落"""
    return (
        f'<w:p>'
        f'<w:pPr><w:pStyle w:val="TOC1"/></w:pPr>'
        f'<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        f'<w:r><w:instrText xml:space="preserve"> TOC \\o "1-{depth}" \\h \\z \\u </w:instrText></w:r>'
        f'<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        f'<w:r><w:t>Table of Contents</w:t></w:r>'
        f'<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        f'</w:p>'
    )


def inject_toc(
    docx_path: Path,
    config: TocConfig,
    seven_zip: Path | None = None,
) -> TocReport:
    """主入口：向 docx 文档中注入格式化 TOC 目录。"""
    t0 = time.time()
    docx_path = Path(docx_path).resolve()
    if not docx_path.is_file():
        raise FileNotFoundError(f"文件不存在: {docx_path}")

    # 1. 读取 docx 压缩包中的关键 XML 条目
    try:
        with zipfile.ZipFile(docx_path, "r") as z:
            namelist = z.namelist()
            if "word/document.xml" not in namelist or "word/styles.xml" not in namelist:
                raise ValueError("Invalid docx: missing word/document.xml or word/styles.xml")

            doc_bytes = z.read("word/document.xml")
            styles_xml = z.read("word/styles.xml").decode("utf-8")
            settings_xml = (
                z.read("word/settings.xml").decode("utf-8")
                if "word/settings.xml" in namelist
                else '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
            )
    except zipfile.BadZipFile as e:
        raise ValueError(f"Invalid docx: bad zip archive: {e}") from e

    # 2. 标题检测与过滤
    styles_bytes = styles_xml.encode("utf-8")
    all_headings = detect_headings(doc_bytes, styles_bytes)
    safe_depth = max(1, min(config.depth, 3))
    valid_headings = [h for h in all_headings if h.level <= safe_depth]

    if not valid_headings:
        raise NoHeadingError(f"未在文档中检测到符合深度 <= {safe_depth} 的标题")

    # 3. 书签避让规划
    bm_prefix = _generate_bookmark_prefix(doc_bytes)
    start_bm_id = _find_max_bookmark_id(doc_bytes) + 1

    bm_assignments: list[tuple[int, str]] = []
    for i in range(len(valid_headings)):
        bm_id = start_bm_id + i
        bm_name = f"{bm_prefix}{i + 1:06d}"
        bm_assignments.append((bm_id, bm_name))

    # 4. 插入位置定位
    insert_pos: int
    replace_start: int
    replace_end: int
    has_own_title = True

    if config.position == Position.DOCUMENT_START:
        m_body = re.search(rb"<w:body\b[^>]*>", doc_bytes)
        if not m_body:
            raise ValueError("Invalid document.xml: <w:body> not found")
        insert_pos = m_body.end()
        replace_start = insert_pos
        replace_end = insert_pos

    elif config.position == Position.BEFORE_FIRST_HEADING:
        insert_pos = valid_headings[0].p_start
        replace_start = insert_pos
        replace_end = insert_pos

    elif config.position == Position.REPLACE_PLACEHOLDER:
        m_toc = doc_bytes.find(b"Table of Contents")
        m_ph = doc_bytes.find(b"From Microsoft Word")
        if m_toc == -1 or m_ph == -1:
            raise ValueError(
                "未在文档中找到公司模板默认目录占位符 ('Table of Contents' 或 'From Microsoft Word')"
            )
        p_toc_title_end = doc_bytes.find(b"</w:p>", m_toc) + 6
        p_empty_center_end = doc_bytes.find(b"</w:p>", p_toc_title_end) + 6
        p_ph_end = doc_bytes.find(b"</w:p>", m_ph) + 6

        replace_start = p_empty_center_end
        replace_end = p_ph_end
        insert_pos = replace_start
        has_own_title = False  # 原模板已有 Table of Contents 标题段落

    elif config.position == Position.BOOKMARK:
        if not config.bookmark_name:
            raise ValueError("Position.BOOKMARK requires config.bookmark_name")
        bm_pattern = re.compile(
            rb'<w:bookmarkStart\b[^>]*?w:name="'
            + re.escape(config.bookmark_name.encode("utf-8"))
            + rb'"',
            re.DOTALL,
        )
        bm_match = bm_pattern.search(doc_bytes)
        if not bm_match:
            raise ValueError(f"Bookmark '{config.bookmark_name}' not found")
        bm_p_end = doc_bytes.find(b"</w:p>", bm_match.start())
        if bm_p_end == -1:
            raise ValueError(f"Bookmark '{config.bookmark_name}' paragraph end not found")
        insert_pos = bm_p_end + 6
        replace_start = insert_pos
        replace_end = insert_pos
    else:
        raise ValueError(f"Unknown position: {config.position}")

    # 5. 计算物理页码
    pages = compute_pages(
        doc_bytes=doc_bytes,
        heading_positions=[h.p_start for h in valid_headings],
        insert_pos=insert_pos,
        mode=config.page_mode,
        page_offset=config.page_offset,
    )

    # 6. 生成 TOC 目录主体 XML
    toc_xml_parts: list[str] = []

    # 目录标题段落
    if has_own_title and config.title:
        escaped_title = html.escape(config.title)
        toc_xml_parts.append(
            f'<w:p><w:pPr><w:pStyle w:val="TOC"/></w:pPr>'
            f'<w:r><w:t>{escaped_title}</w:t></w:r></w:p>'
        )

    # 目录项内容
    if config.page_mode == PageMode.DYNAMIC:
        toc_xml_parts.append(_build_dynamic_toc_xml(safe_depth))
    else:
        total_items = len(valid_headings)
        for i, heading in enumerate(valid_headings):
            _, bm_name = bm_assignments[i]
            p_num = pages[i] if i < len(pages) else 1
            entry_xml = _build_toc_entry_xml(
                index=i,
                total=total_items,
                heading=heading,
                bm_name=bm_name,
                page_num=p_num,
                style=config.style,
                depth=safe_depth,
            )
            toc_xml_parts.append(entry_xml)

        # 追加闭合域标记段落
        p_close_field = (
            '<w:p><w:pPr><w:rPr><w:noProof/></w:rPr></w:pPr>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
        )
        toc_xml_parts.append(p_close_field)

    # 7. STATIC_RESTART 分节处理
    # 提取现有最后一个 sectPr 作为分节模板
    sect_matches = list(re.finditer(rb"<w:sectPr\b.*?</w:sectPr>", doc_bytes, re.DOTALL))
    if sect_matches:
        raw_last_sect = sect_matches[-1].group(0).decode("utf-8", errors="replace")
        # 清除已有的 pgNumType
        sect_template = re.sub(r"<w:pgNumType\b[^>]*/>", "", raw_last_sect)
    else:
        sect_template = (
            '<w:sectPr>'
            '<w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>'
            '</w:sectPr>'
        )

    # 目录节后追加独立分节符
    if config.page_mode == PageMode.STATIC_RESTART:
        toc_sect_p = f"<w:p><w:pPr>{sect_template}</w:pPr></w:p>"
        toc_xml_parts.append(toc_sect_p)

    toc_replacement_bytes = "".join(toc_xml_parts).encode("utf-8")

    # 8. 正文标题书签注入与正文 sectPr 改造规划
    replacements_list: list[tuple[int, int, bytes]] = []

    # 目录插入点
    replacements_list.append((replace_start, replace_end, toc_replacement_bytes))

    # 各标题段落书签注入
    for i, heading in enumerate(valid_headings):
        bm_id, bm_name = bm_assignments[i]
        orig_p_bytes = doc_bytes[heading.p_start : heading.p_end]

        bm_start_bytes = f'<w:bookmarkStart w:id="{bm_id}" w:name="{bm_name}"/>'.encode("utf-8")
        bm_end_bytes = f'<w:bookmarkEnd w:id="{bm_id}"/>'.encode("utf-8")

        if b"</w:pPr>" in orig_p_bytes:
            new_p_bytes = orig_p_bytes.replace(b"</w:pPr>", b"</w:pPr>" + bm_start_bytes, 1)
        else:
            m_tag = re.match(rb"<w:p\b[^>]*>", orig_p_bytes)
            assert m_tag is not None
            pos_end_tag = m_tag.end()
            new_p_bytes = orig_p_bytes[:pos_end_tag] + bm_start_bytes + orig_p_bytes[pos_end_tag:]

        new_p_bytes = new_p_bytes[:-6] + bm_end_bytes + b"</w:p>"
        replacements_list.append((heading.p_start, heading.p_end, new_p_bytes))

    # 正文首个分节符添加/修改 pgNumType start="1"
    if config.page_mode == PageMode.STATIC_RESTART and sect_matches:
        # 定位 insert_pos 之后的第一个 sectPr（若无则取最后一个）
        target_sect = None
        for s in sect_matches:
            if s.start() >= insert_pos:
                target_sect = s
                break
        if target_sect is None:
            target_sect = sect_matches[-1]

        sect_bytes = target_sect.group(0)
        sect_str = sect_bytes.decode("utf-8", errors="replace")

        if "<w:pgNumType" in sect_str:
            new_sect_str = re.sub(
                r'<w:pgNumType\b([^>]*?)/>',
                r'<w:pgNumType\1 w:start="1"/>',
                sect_str,
            )
            # 若已有 w:start 则替换其值
            new_sect_str = re.sub(r'w:start="[^"]*"', 'w:start="1"', new_sect_str)
        else:
            # 插入在 <w:sectPr 开标签后
            m_open = re.match(r"<w:sectPr\b[^>]*>", sect_str)
            assert m_open is not None
            idx = m_open.end()
            new_sect_str = sect_str[:idx] + '<w:pgNumType w:start="1"/>' + sect_str[idx:]

        replacements_list.append(
            (target_sect.start(), target_sect.end(), new_sect_str.encode("utf-8"))
        )

    # 9. 排序并组装新 document.xml
    replacements_list.sort(key=lambda item: (item[0], item[1]))

    chunks: list[bytes] = []
    curr_pos = 0
    for r_start, r_end, r_bytes in replacements_list:
        if r_start < curr_pos:
            # 兼容重叠区间（如目录插入在标题段落起始处）
            continue
        chunks.append(doc_bytes[curr_pos:r_start])
        chunks.append(r_bytes)
        curr_pos = r_end
    chunks.append(doc_bytes[curr_pos:])
    new_doc_bytes = b"".join(chunks)

    # 10. 处理 styles.xml 注入（按样式 ID 逐个补缺，不重复注入已有样式）
    style_frag = toc_styles_fragment(
        config.style, safe_depth, existing_styles_xml=styles_xml
    )
    if style_frag:
        new_styles_xml = styles_xml.replace("</w:styles>", style_frag + "</w:styles>")
    else:
        new_styles_xml = styles_xml

    # 11. 处理 settings.xml updateFields
    if config.page_mode == PageMode.DYNAMIC:
        if "<w:updateFields" in settings_xml:
            new_settings_xml = re.sub(
                r'<w:updateFields\b[^>]*?/>',
                '<w:updateFields w:val="true"/>',
                settings_xml,
            )
        elif "</w:settings>" in settings_xml:
            new_settings_xml = settings_xml.replace(
                "</w:settings>", '<w:updateFields w:val="true"/></w:settings>'
            )
        else:
            new_settings_xml = re.sub(
                r'<w:settings\b([^>]*?)\s*/>',
                r'<w:settings\1><w:updateFields w:val="true"/></w:settings>',
                settings_xml,
            )
    else:
        new_settings_xml = re.sub(r'<w:updateFields\b[^>]*?/>', "", settings_xml)
        new_settings_xml = new_settings_xml.replace('<w:updateFields w:val="true"/>', '')
        new_settings_xml = new_settings_xml.replace('<w:updateFields w:val="1"/>', '')
        new_settings_xml = new_settings_xml.replace('<w:updateFields/>', '')

    # 12. 备份原文件
    backup_path = None
    if config.backup:
        potential_backup = docx_path.with_name(docx_path.name + ".bak")
        if not potential_backup.exists():
            shutil.copy2(docx_path, potential_backup)
            backup_path = potential_backup

    # 13. 原子写回
    rewrite_zip(
        docx_path=docx_path,
        replacements={
            "word/document.xml": new_doc_bytes,
            "word/styles.xml": new_styles_xml.encode("utf-8"),
            "word/settings.xml": new_settings_xml.encode("utf-8"),
        },
        seven_zip=seven_zip,
    )

    elapsed = time.time() - t0
    return TocReport(
        entries=len(valid_headings),
        pages=pages,
        backup_path=backup_path,
        elapsed=elapsed,
    )
