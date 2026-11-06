"""
DOCX 字节流注入器：TOC 目录注入引擎主入口。
纯 Python 字节流切片、旧目录自动剥离、书签全局避让、分节页码重启/实测锚点、4 样式注入及原子写回。
"""

from __future__ import annotations

import html
import math
import re
import shutil
import time
import uuid
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
import zipfile

from wdt.toc.existing_toc_stripper import strip_existing_toc
from wdt.toc.heading_detector import Heading, detect_headings
from wdt.toc.page_calculator import PageMode, compute_pages
from wdt.toc.page_measurer import build_heading_anchor_xml, build_sentinel_xml
from wdt.toc.styles_templates import HYPERLINK_STYLE_ID, TocStyle, toc_styles_fragment
from wdt.toc.zip_rewriter import rewrite_zip

# 默认目录每页约容纳 38 行（基准版心高 11187 twips = SAS landscape 正文节 11907 - 上 432 - 下 288，
# Times New Roman 9-10pt、紧凑行距，已实测验证；其余版式由 _toc_lines_per_page 按版心高等比换算）
# 真实场景中，如遇多行标题折行或宽行距，可由调用方传入 explicit page_offset 进行覆写
TOC_LINES_PER_PAGE_DEFAULT = 38


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
    page_mode: PageMode = PageMode.STATIC_CONTINUOUS
    page_offset: int = 0
    backup: bool = True


@dataclass
class TocReport:
    entries: int
    pages: list[int]
    backup_path: Path | None
    elapsed: float
    bookmark_names: list[str] = field(default_factory=list)
    heading_texts: list[str] = field(default_factory=list)
    sentinel: str = ""


@dataclass
class BodyFont:
    ascii: str | None = None
    hAnsi: str | None = None
    eastAsia: str | None = None


def detect_body_font(styles_xml: str | bytes) -> BodyFont | None:
    """从 styles.xml 探测正文字体（优先级：docDefaults -> Normal 样式）。"""
    if not styles_xml:
        return None
    xml_str = styles_xml.decode("utf-8", errors="replace") if isinstance(styles_xml, bytes) else styles_xml
    try:
        root = ET.fromstring(xml_str)
    except Exception:
        return None

    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

    # 1. docDefaults -> rPrDefault -> rPr -> rFonts
    rfonts_def = root.find(".//w:docDefaults/w:rPrDefault/w:rPr/w:rFonts", ns)
    if rfonts_def is not None:
        ascii_f = rfonts_def.attrib.get(f"{{{ns['w']}}}ascii")
        hansi_f = rfonts_def.attrib.get(f"{{{ns['w']}}}hAnsi")
        eastasia_f = rfonts_def.attrib.get(f"{{{ns['w']}}}eastAsia")
        if ascii_f or hansi_f or eastasia_f:
            return BodyFont(ascii=ascii_f, hAnsi=hansi_f, eastAsia=eastasia_f)

    # 2. Normal 样式 (styleId="Normal" 或 w:default="1" 且 w:type="paragraph")
    for style in root.findall(".//w:style", ns):
        stype = style.attrib.get(f"{{{ns['w']}}}type")
        sid = style.attrib.get(f"{{{ns['w']}}}styleId", "")
        is_default = style.attrib.get(f"{{{ns['w']}}}default") == "1"
        if stype == "paragraph" and (sid == "Normal" or is_default):
            rfonts_norm = style.find("./w:rPr/w:rFonts", ns)
            if rfonts_norm is not None:
                ascii_f = rfonts_norm.attrib.get(f"{{{ns['w']}}}ascii")
                hansi_f = rfonts_norm.attrib.get(f"{{{ns['w']}}}hAnsi")
                eastasia_f = rfonts_norm.attrib.get(f"{{{ns['w']}}}eastAsia")
                if ascii_f or hansi_f or eastasia_f:
                    return BodyFont(ascii=ascii_f, hAnsi=hansi_f, eastAsia=eastasia_f)

    return None


_RFONTS_TAGS = re.compile(r"<w:rFonts\b[^>]*/>")

def _dominant_font_attrs(doc_bytes: bytes) -> dict[str, str]:
    """统计正文 run 中 rFonts 各字体属性的频次众数。

    LO 转换文档的 docDefaults 常与正文实际字体不一致（正文由 run 级直接格式决定），
    以正文众数字体为准可使目录与正文视觉统一。
    """
    xml = doc_bytes.decode("utf-8", errors="ignore")
    counts: dict[str, dict[str, int]] = {"ascii": {}, "hAnsi": {}, "eastAsia": {}}
    for tag in _RFONTS_TAGS.findall(xml):
        for attr, key in (("w:ascii", "ascii"), ("w:hAnsi", "hAnsi"), ("w:eastAsia", "eastAsia")):
            m = re.search(attr + r'="([^"]+)"', tag)
            if m:
                v = m.group(1)
                counts[key][v] = counts[key].get(v, 0) + 1
    result: dict[str, str] = {}
    for key, c in counts.items():
        if c:
            top, n = max(c.items(), key=lambda kv: kv[1])
            if n >= 10:
                result[key] = top
    return result


class NoHeadingError(ValueError):
    """未在文档中检测到符合层级条件的标题"""
    pass


# 跨页续表头标记：SAS Listing 每页重复输出带 "(Continued)" / "（续）" 的表头
_CONTINUED_HEADING_RE = re.compile(r"(?i)\(\s*continued\s*\)|（\s*续\s*）|\(\s*续\s*\)")


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


def _parse_sect_geometry(sect_xml: str) -> tuple[int, int, int]:
    """解析 sectPr 中的页面宽度与左右页边距（dxa）。返回 (width, left_margin, right_margin)。"""
    width = 11906
    left = 1440
    right = 1440
    m_w = re.search(r'<w:pgSz\b[^>]*?w:w="(\d+)"', sect_xml)
    if m_w:
        width = int(m_w.group(1))
    m_left = re.search(r'<w:pgMar\b[^>]*?w:left="(\d+)"', sect_xml)
    if m_left:
        left = int(m_left.group(1))
    m_right = re.search(r'<w:pgMar\b[^>]*?w:right="(\d+)"', sect_xml)
    if m_right:
        right = int(m_right.group(1))
    return width, left, right


def _format_rfonts(default_ascii: str, default_eastasia: str, default_hansi: str, body_font: BodyFont | None) -> str:
    """结合正文探测字体生成 rFonts 标签"""
    ascii_f = body_font.ascii if (body_font and body_font.ascii) else default_ascii
    eastasia_f = body_font.eastAsia if (body_font and body_font.eastAsia) else default_eastasia
    hansi_f = body_font.hAnsi if (body_font and body_font.hAnsi) else default_hansi
    return f'<w:rFonts w:ascii="{ascii_f}" w:eastAsia="{eastasia_f}" w:hAnsi="{hansi_f}"/>'


def _build_toc_entry_xml(
    index: int,
    total: int,
    heading: Heading,
    bm_name: str,
    page_num: int | None,
    style: TocStyle,
    depth: int,
    tab_pos: int = 9026,
    body_font: BodyFont | None = None,
    is_dynamic: bool = False,
) -> str:
    """生成单条 TOC 目录项段落 XML"""
    level = heading.level
    escaped_title = html.escape(heading.text)

    # 导线设置
    leader = "none" if style == TocStyle.MINIMAL_CLEAN else "dot"
    tab_pr = f'<w:tabs><w:tab w:val="right" w:leader="{leader}" w:pos="{tab_pos}"/></w:tabs>'

    # 标题字体与颜色
    if style == TocStyle.MODERN_ACCENT:
        fonts_xml = _format_rfonts("Calibri", "微软雅黑", "Calibri", body_font)
        if level == 1:
            title_rpr = (
                '<w:rPr>'
                f'<w:rStyle w:val="{HYPERLINK_STYLE_ID}"/>'
                f'{fonts_xml}'
                '<w:b/><w:color w:val="1F4E79"/><w:noProof/>'
                '</w:rPr>'
            )
        else:
            title_rpr = (
                '<w:rPr>'
                f'<w:rStyle w:val="{HYPERLINK_STYLE_ID}"/>'
                f'{fonts_xml}'
                '<w:color w:val="262626"/><w:noProof/>'
                '</w:rPr>'
            )
    elif style == TocStyle.REGULATORY_LISTING:
        fonts_xml = _format_rfonts("Times New Roman", "宋体", "Times New Roman", body_font)
        bold_tag = "<w:b/><w:bCs/>" if level == 1 else ""
        title_rpr = (
            '<w:rPr>'
            f'<w:rStyle w:val="{HYPERLINK_STYLE_ID}"/>'
            f'{fonts_xml}'
            f'{bold_tag}'
            '<w:color w:val="000000"/><w:noProof/>'
            '</w:rPr>'
        )
    elif style == TocStyle.CLASSIC_DOT:
        fonts_xml = _format_rfonts("Times New Roman", "宋体", "Times New Roman", body_font)
        bold_tag = "<w:b/><w:bCs/>" if level == 1 else ""
        title_rpr = (
            '<w:rPr>'
            f'<w:rStyle w:val="{HYPERLINK_STYLE_ID}"/>'
            f'{fonts_xml}'
            f'{bold_tag}'
            '<w:color w:val="111827"/><w:noProof/>'
            '</w:rPr>'
        )
    else:  # MINIMAL_CLEAN
        fonts_xml = _format_rfonts("Arial", "微软雅黑", "Arial", body_font)
        title_rpr = (
            '<w:rPr>'
            f'<w:rStyle w:val="{HYPERLINK_STYLE_ID}"/>'
            f'{fonts_xml}'
            '<w:color w:val="333333"/><w:noProof/>'
            '</w:rPr>'
        )

    # 页码字体样式（支持 body_font 覆盖）
    page_font_frag = _format_rfonts("Times New Roman", "宋体", "Times New Roman", body_font) if body_font else ""
    if style == TocStyle.MINIMAL_CLEAN:
        page_rpr = f'<w:rPr>{page_font_frag}<w:sz w:val="18"/><w:szCs w:val="18"/><w:color w:val="7F7F7F"/><w:noProof/></w:rPr>'
    else:
        page_rpr = f'<w:rPr>{page_font_frag}<w:noProof/></w:rPr>'

    page_str = str(page_num) if page_num is not None else ""

    # 关键排版设计：
    # 在静态实测模式下，直接生成确凿的纯文本页码，保留超链接跳转，
    # 彻底杜绝嵌套 PAGEREF 域在 LibreOffice 无头导出 PDF 时被强制重新求值清空的问题！
    if not is_dynamic:
        hyperlink_block = (
            f'<w:hyperlink w:anchor="{bm_name}" w:history="1">'
            f'<w:r>{title_rpr}<w:t xml:space="preserve">{escaped_title}</w:t></w:r>'
            f'<w:r><w:tab/></w:r>'
            f'<w:r>{page_rpr}<w:t>{page_str}</w:t></w:r>'
            f'</w:hyperlink>'
        )
    else:
        # 仅在动态域模式下输出 PAGEREF
        hyperlink_block = (
            f'<w:hyperlink w:anchor="{bm_name}" w:history="1">'
            f'<w:r>{title_rpr}<w:t xml:space="preserve">{escaped_title}</w:t></w:r>'
            f'<w:r><w:tab/></w:r>'
            f'<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            f'<w:r><w:instrText xml:space="preserve"> PAGEREF {bm_name} </w:instrText></w:r>'
            f'<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            f'<w:r>{page_rpr}<w:t>{page_str}</w:t></w:r>'
            f'<w:r><w:fldChar w:fldCharType="end"/></w:r>'
            f'</w:hyperlink>'
        )

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


# 模板目录标题页标题文本特征（公司模板预留的 portrait 目录占位页）
_TOC_TITLE_TEXT = ("Table of Content", "目录")


def _find_leading_toc_title_page(
    doc_bytes: bytes, body_end: int, first_heading_pos: int
) -> tuple[int, int, bytes, bytes] | None:
    """在文首首个标题之前查找独占一个 portrait 节的目录标题占位页。

    返回 (标题段起始, 占位页结束, sectPr 字节, 去除 sectPr 后的标题段字节)；
    占位页结束涵盖标题段之后仅含分节符的空段。
    未找到时返回 None。
    """
    region = doc_bytes[body_end:first_heading_pos]
    paragraphs = list(re.finditer(rb"<w:p\b[^>]*>.*?</w:p>", region, re.DOTALL))
    for idx, m in enumerate(paragraphs):
        seg = m.group(0)
        text = seg.decode("utf-8", errors="replace")
        if not any(t in text for t in _TOC_TITLE_TEXT):
            continue
        sm = re.search(rb"<w:sectPr\b.*?</w:sectPr>", seg, re.DOTALL)
        page_end = m.end()
        fixed = seg
        if sm is None:
            # 分节符不在标题段内：向后查找标题页收尾的纯分节符空段
            # （SAS 模板布局：标题段 + 独立的分节空段）
            for m2 in paragraphs[idx + 1 : idx + 4]:
                seg2 = m2.group(0)
                inner_texts = re.findall(rb"<w:t\b[^>]*>([^<]*)</w:t>", seg2)
                if any(t.strip() for t in inner_texts):
                    break  # 中间出现有文字的段落，不属于标题页
                sm2 = re.search(rb"<w:sectPr\b.*?</w:sectPr>", seg2, re.DOTALL)
                if sm2 is not None:
                    sm = sm2
                    page_end = m2.end()
                    break
            if sm is None:
                continue
        else:
            fixed = seg[: sm.start()] + seg[sm.end() :]
        pg = re.search(r"<w:pgSz\b[^>]*/>", sm.group(0).decode("utf-8", errors="replace"))
        if pg is not None:
            attrs = pg.group(0)
            wm = re.search(r'w:w="(\d+)"', attrs)
            hm = re.search(r'w:h="(\d+)"', attrs)
            if wm and hm:
                w, h = int(wm.group(1)), int(hm.group(1))
                landscape = 'w:orient="landscape"' in attrs or w > h
                if landscape:
                    continue  # 只复用 portrait 的模板标题页
        return (body_end + m.start(), body_end + page_end, sm.group(0), fixed)
    return None


def _toc_lines_per_page(sect_template: str) -> int:
    """按目录节版心高度等比换算每页可容纳的目录行数。

    基准：TOC_LINES_PER_PAGE_DEFAULT 行 ↔ 版心高 11187 twips
    （SAS landscape 正文节：11907 - 上 432 - 下 288，已实测验证）。
    """
    pg = re.search(r"<w:pgSz\b[^>]*/>", sect_template)
    mar = re.search(r"<w:pgMar\b[^>]*/>", sect_template)

    def _attr(tag: str, name: str) -> int | None:
        m = re.search(name + r'="(\d+)"', tag)
        return int(m.group(1)) if m else None

    if pg is None or mar is None:
        return TOC_LINES_PER_PAGE_DEFAULT
    attrs = pg.group(0)
    w = _attr(attrs, "w:w")
    h = _attr(attrs, "w:h")
    top = _attr(mar.group(0), "w:top") or 1440
    bottom = _attr(mar.group(0), "w:bottom") or 1440
    if w is None or h is None:
        return TOC_LINES_PER_PAGE_DEFAULT
    landscape = 'w:orient="landscape"' in attrs or w > h
    eff_h = min(w, h) if landscape else max(w, h)
    text_h = max(4000, eff_h - top - bottom)
    return max(15, round(TOC_LINES_PER_PAGE_DEFAULT * text_h / 11187))


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

    # 2. 先行剥离既有旧目录（含 sdt 块、字段段落与 SAS 模板占位节），后续计算基于干净字节流
    doc_bytes = strip_existing_toc(doc_bytes, styles_xml)

    # 3. 探测正文字体（docDefaults/Normal 样式），并以正文 run 众数字体修正
    body_font = detect_body_font(styles_xml)
    if body_font is not None:
        dom = _dominant_font_attrs(doc_bytes)
        if dom:
            body_font = BodyFont(
                ascii=dom.get("ascii", body_font.ascii),
                hAnsi=dom.get("hAnsi", body_font.hAnsi),
                eastAsia=dom.get("eastAsia", body_font.eastAsia),
            )

    # 4. 标题检测与过滤
    styles_bytes = styles_xml.encode("utf-8")
    all_headings = detect_headings(doc_bytes, styles_bytes)
    safe_depth = max(1, min(config.depth, 3))
    # 过滤超过目标深度的标题及跨页续表头（如 SAS Listing 每页重复的 "(Continued)"）
    valid_headings = [
        h
        for h in all_headings
        if h.level <= safe_depth and not _CONTINUED_HEADING_RE.search(h.text)
    ]

    if not valid_headings:
        raise NoHeadingError(f"未在文档中检测到符合深度 <= {safe_depth} 的标题")

    # 5. 书签避让规划
    bm_prefix = _generate_bookmark_prefix(doc_bytes)
    start_bm_id = _find_max_bookmark_id(doc_bytes) + 1

    bm_assignments: list[tuple[int, str]] = []
    for i in range(len(valid_headings)):
        bm_id = start_bm_id + i
        bm_name = f"{bm_prefix}{i + 1:06d}"
        bm_assignments.append((bm_id, bm_name))

    # 6. 插入位置定位
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
        has_own_title = False

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

    # 6b. 模板目录标题页复用检测：文首存在 portrait 的目录占位页时，
    # 将目录条目并入该节（标题页不再孤立为空白页，目录自然继承 portrait 版式）
    reuse_title: tuple[int, int, bytes, bytes] | None = None
    first_heading_pos = valid_headings[0].p_start
    if config.position in (Position.DOCUMENT_START, Position.BEFORE_FIRST_HEADING):
        m_body_start = re.search(rb"<w:body\b[^>]*>", doc_bytes)
        assert m_body_start is not None
        reuse_title = _find_leading_toc_title_page(
            doc_bytes, m_body_start.end(), first_heading_pos
        )
    if reuse_title is not None:
        _p_start, _p_end, _sect_span, _fixed = reuse_title
        insert_pos = _p_end
        replace_start = _p_start
        replace_end = _p_end
        has_own_title = False

    # 7. 分节模板与右制表位绝对计算（取正文首个标题之后首个 sectPr，无则取 insert_pos 之后或最后一个）
    sect_matches = list(re.finditer(rb"<w:sectPr\b.*?</w:sectPr>", doc_bytes, re.DOTALL))
    target_sect = None
    if sect_matches:
        for s in sect_matches:
            if s.start() >= first_heading_pos:
                target_sect = s
                break
        if target_sect is None:
            for s in sect_matches:
                if s.start() >= insert_pos:
                    target_sect = s
                    break
        if target_sect is None:
            target_sect = sect_matches[-1]

    if reuse_title is not None:
        # 复用模板标题页：目录节版式 = 原标题页所在节的版式（portrait，含页脚）
        sect_template = re.sub(
            r"<w:pgNumType\b[^>]*/>", "", reuse_title[2].decode("utf-8", errors="replace")
        )
    elif target_sect is not None:
        raw_sect = target_sect.group(0).decode("utf-8", errors="replace")
        sect_template = re.sub(r"<w:pgNumType\b[^>]*/>", "", raw_sect)
    else:
        sect_template = (
            '<w:sectPr>'
            '<w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>'
            '</w:sectPr>'
        )

    pg_w, pg_left, pg_right = _parse_sect_geometry(sect_template)
    # 版心宽度，保留 120 dxa (~6pt) 安全边界以防极端字体触顶换行
    right_tab_pos = max(1000, pg_w - pg_left - pg_right - 120)

    # 8. 计算页码（STATIC_MEASURED 模式先置占位 0，同时准备 break-count 兜底页码）
    # 在 STATIC_CONTINUOUS 模式下，预估目录物理占用页数（默认每页约 38 行，最少 1 页）
    calc_offset = config.page_offset
    if config.page_mode == PageMode.STATIC_CONTINUOUS:
        if calc_offset == 0:
            # 目录行数 = 标题项数 + (1 if has_own_title and config.title else 0)
            toc_lines = len(valid_headings) + (1 if (has_own_title and config.title) else 0)
            calc_offset = max(1, math.ceil(toc_lines / _toc_lines_per_page(sect_template)))
        else:
            import logging
            logging.getLogger("wdt.toc").info(
                "STATIC_CONTINUOUS 模式下检测到用户指定 page_offset=%d，将跳过按行数自动估算", calc_offset
            )

    break_count_pages = compute_pages(
        doc_bytes=doc_bytes,
        heading_positions=[(h.p_start, h.p_end) for h in valid_headings],
        insert_pos=insert_pos,
        mode=config.page_mode if config.page_mode != PageMode.STATIC_MEASURED else PageMode.STATIC_RESTART,
        page_offset=calc_offset,
    )
    if config.page_mode == PageMode.STATIC_MEASURED:
        pages = [0] * len(valid_headings)
    else:
        pages = break_count_pages

    # 9. 生成 TOC 目录主体 XML
    toc_xml_parts: list[str] = []

    # 目录标题段落
    if has_own_title and config.title:
        escaped_title = html.escape(config.title)
        title_rfonts = _format_rfonts("Times New Roman", "宋体", "Times New Roman", body_font) if body_font else ""
        title_rpr_str = f"<w:rPr>{title_rfonts}</w:rPr>" if title_rfonts else ""
        toc_xml_parts.append(
            f'<w:p><w:pPr><w:pStyle w:val="TOC"/></w:pPr>'
            f'<w:r>{title_rpr_str}<w:t>{escaped_title}</w:t></w:r></w:p>'
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
                tab_pos=right_tab_pos,
                body_font=body_font,
                is_dynamic=False,
            )
            toc_xml_parts.append(entry_xml)

    # 10. 分节与前哨标记处理
    # STATIC_CONTINUOUS、STATIC_RESTART、STATIC_MEASURED 与 DYNAMIC 均在目录节后注入独立分节符，
    # 隔离目录与正文版式
    if config.page_mode in (PageMode.STATIC_CONTINUOUS, PageMode.STATIC_RESTART, PageMode.STATIC_MEASURED, PageMode.DYNAMIC):
        toc_sect_p = f"<w:p><w:pPr>{sect_template}</w:pPr></w:p>"
        toc_xml_parts.append(toc_sect_p)

    sentinel_uuid = uuid.uuid4().hex[:8]
    if config.page_mode == PageMode.STATIC_MEASURED:
        sentinel_p = build_sentinel_xml(sentinel_uuid)
        toc_xml_parts.append(sentinel_p)

    toc_replacement_bytes = "".join(toc_xml_parts).encode("utf-8")
    if reuse_title is not None:
        # 复用模板标题页：标题段（已剥除 sectPr）置前，目录条目与收尾分节空段落紧随其后
        toc_replacement_bytes = reuse_title[3] + toc_replacement_bytes

    # 11. 正文标题书签/隐藏锚点注入与正文 sectPr 改造规划
    replacements_list: list[tuple[int, int, bytes]] = []

    # 目录插入点
    replacements_list.append((replace_start, replace_end, toc_replacement_bytes))

    # 各标题段落书签及隐藏锚点注入
    for i, heading in enumerate(valid_headings):
        bm_id, bm_name = bm_assignments[i]
        orig_p_bytes = doc_bytes[heading.p_start : heading.p_end]

        bm_start_bytes = f'<w:bookmarkStart w:id="{bm_id}" w:name="{bm_name}"/>'.encode("utf-8")
        if config.page_mode == PageMode.STATIC_MEASURED:
            anchor_xml = build_heading_anchor_xml(i)
            bm_start_bytes = bm_start_bytes + anchor_xml.encode("utf-8")

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
    if config.page_mode in (PageMode.STATIC_RESTART, PageMode.STATIC_MEASURED, PageMode.DYNAMIC) and sect_matches:
        sect_bytes = target_sect.group(0)
        sect_str = sect_bytes.decode("utf-8", errors="replace")

        if "<w:pgNumType" in sect_str:
            new_sect_str = re.sub(
                r'<w:pgNumType\b([^>]*?)/>',
                r'<w:pgNumType\1 w:start="1"/>',
                sect_str,
            )
            new_sect_str = re.sub(r'w:start="[^"]*"', 'w:start="1"', new_sect_str)
        else:
            m_open = re.match(r"<w:sectPr\b[^>]*>", sect_str)
            assert m_open is not None
            idx = m_open.end()
            new_sect_str = sect_str[:idx] + '<w:pgNumType w:start="1"/>' + sect_str[idx:]

        replacements_list.append(
            (target_sect.start(), target_sect.end(), new_sect_str.encode("utf-8"))
        )
    elif config.page_mode == PageMode.STATIC_CONTINUOUS and sect_matches:
        # CONTINUOUS 模式下，正文节绝不能带 w:start（否则会重新从 1 开始编号造成末页 299/302）
        # 若源文档或之前处理残留了 w:pgNumType / w:start，必须予以彻底剥除
        sect_bytes = target_sect.group(0)
        sect_str = sect_bytes.decode("utf-8", errors="replace")
        if "<w:pgNumType" in sect_str:
            # 移除 w:start 属性；若 pgNumType 只包含 start 则连同标签一起移除
            new_sect_str = re.sub(r'\s*w:start="[^"]*"', '', sect_str)
            new_sect_str = re.sub(r'<w:pgNumType\s*/>', '', new_sect_str)
            if new_sect_str != sect_str:
                replacements_list.append(
                    (target_sect.start(), target_sect.end(), new_sect_str.encode("utf-8"))
                )

    # 12. 排序并组装新 document.xml
    replacements_list.sort(key=lambda item: (item[0], item[1]))

    chunks: list[bytes] = []
    curr_pos = 0
    for r_start, r_end, r_bytes in replacements_list:
        if r_start < curr_pos:
            continue
        chunks.append(doc_bytes[curr_pos:r_start])
        chunks.append(r_bytes)
        curr_pos = r_end
    chunks.append(doc_bytes[curr_pos:])
    new_doc_bytes = b"".join(chunks)

    # 13. 处理 styles.xml 注入
    style_frag = toc_styles_fragment(
        config.style, safe_depth, existing_styles_xml=styles_xml
    )
    if style_frag:
        new_styles_xml = styles_xml.replace("</w:styles>", style_frag + "</w:styles>")
    else:
        new_styles_xml = styles_xml

    # 14. 处理 settings.xml updateFields
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

    # 15. 备份原文件
    backup_path = None
    if config.backup:
        potential_backup = docx_path.with_name(docx_path.name + ".bak")
        if not potential_backup.exists():
            shutil.copy2(docx_path, potential_backup)
            backup_path = potential_backup

    # 16. 原子写回
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
        pages=break_count_pages if config.page_mode == PageMode.STATIC_MEASURED else pages,
        backup_path=backup_path,
        elapsed=elapsed,
        bookmark_names=[bm_name for _, bm_name in bm_assignments],
        heading_texts=[h.text for h in valid_headings],
        sentinel=sentinel_uuid if config.page_mode == PageMode.STATIC_MEASURED else "",
    )
