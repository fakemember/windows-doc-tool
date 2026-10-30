"""
LibreOffice 转换产物规范化器。

修正 LO 转换 DOC/RTF -> DOCX 时引入的三类保真度损失：

1. 段落标记字号丢失：LO 导出为空 <w:pPr><w:rPr></w:rPr>，段落标记回落到
   docDefaults 字号（常见 12pt），把 10pt 正文的行高撑大、分页膨胀。
   修复：段落标记字号补齐为该段落 run 的最大字号。
2. 空白填充变形：LO 将单元格内连续空格转换为 "空格+U+2006（六分空）" 交替
   序列，字符量膨胀，且在固定列宽内可能引发意外换行。
   修复：折叠回普通空格。
3. CJK 字体替换：LO 用本机替代字体（如 Noto Serif SC）重写源文档缺失字体名，
   其行高盒与 Word 的替代结果不一致，导致 Word/LO 双端分页分歧、实测页码失准。
   修复：换回 Windows 标准中文字体，使双端渲染度量一致。
"""

from __future__ import annotations

import logging
import re
import zipfile
from pathlib import Path

from wdt.toc.zip_rewriter import rewrite_zip

logger = logging.getLogger(__name__)

_P_BLOCK = re.compile(r"<w:p(?:\s[^>]*)?>.*?</w:p>", re.DOTALL)
_RUN_SZ = re.compile(r'<w:sz w:val="(\d+)"/>')
_PPR_PREFIX = re.compile(r"<w:p(?:\s[^>]*)?>(<w:pPr>.*?</w:pPr>)?", re.DOTALL)
_MARK_RPR = re.compile(r"<w:rPr\b[^>]*?(?:/>|>(.*?)</w:rPr>)", re.DOTALL)

# rPr 子元素 OOXML 顺序中位于 sz 之后的元素（sz 需插入其前）
_POST_SZ_TAGS = (
    "<w:szCs ",
    "<w:highlight",
    "<w:u ",
    "<w:effect",
    "<w:bdr",
    "<w:shd",
    "<w:fitText",
    "<w:vertAlign",
    "<w:rtl",
    "<w:cs ",
    "<w:em ",
    "<w:lang",
    "<w:eastAsianLayout",
    "<w:specVanish",
    "<w:oMath",
)

# LO 在 CJK 环境常见的替代字体 -> Windows 标准中文字体。
# 统一映射到宋体：与 Word 对缺失 CJK 字体的回退度量一致，
# 保证 Word/LO 双端渲染（含实测页码）结果相同。
_CJK_FONT_SWAP = {
    "Noto Serif SC": "宋体",
    "Noto Serif CJK SC": "宋体",
    "Noto Sans SC": "宋体",
    "Noto Sans CJK SC": "宋体",
}


def _fix_paragraph_mark(p_block: str, counter: list[int]) -> str:
    """将段落标记 rPr 的字号补齐为该段落 run 的最大字号。"""
    open_end = p_block.index(">") + 1  # 保留 <w:p ...> 开标签
    m = _PPR_PREFIX.match(p_block)
    if m is None:
        return p_block
    ppr = m.group(1)
    rest = p_block[m.end() :]
    szs = _RUN_SZ.findall(rest)
    if not szs:
        return p_block
    mark_sz = max(int(s) for s in szs)
    sz_val = f'<w:sz w:val="{mark_sz}"/>'
    szcs_val = f'<w:szCs w:val="{mark_sz}"/>'

    if ppr is None:
        insert = f"<w:pPr><w:rPr>{sz_val}{szcs_val}</w:rPr></w:pPr>"
        counter[0] += 1
        return p_block[:open_end] + insert + p_block[open_end:]

    m_mark = _MARK_RPR.search(ppr)
    if m_mark is None:
        new_ppr = (
            ppr[: -len("</w:pPr>")]
            + f"<w:rPr>{sz_val}{szcs_val}</w:rPr></w:pPr>"
        )
    else:
        mark_inner = m_mark.group(1)
        if mark_inner is not None and "<w:sz " in mark_inner:
            return p_block  # 段落标记已有字号，保持原样
        add_sz = mark_inner is None or "<w:sz " not in mark_inner
        add_szcs = mark_inner is None or "<w:szCs " not in mark_inner
        sz_xml = (sz_val if add_sz else "") + (szcs_val if add_szcs else "")
        if not sz_xml:
            return p_block
        if not mark_inner:
            new_rpr = f"<w:rPr>{sz_xml}</w:rPr>"
        else:
            # 有子元素时按 OOXML 顺序插在首个 sz 后元素之前
            cut = len(mark_inner)
            for tag in _POST_SZ_TAGS:
                idx = mark_inner.find(tag)
                if idx != -1:
                    cut = min(cut, idx)
            new_rpr = (
                f"<w:rPr>{mark_inner[:cut]}{sz_xml}{mark_inner[cut:]}</w:rPr>"
            )
        new_ppr = ppr[: m_mark.start()] + new_rpr + ppr[m_mark.end() :]

    counter[0] += 1
    return p_block[:open_end] + new_ppr + p_block[m.end() :]


def normalize_docx(docx_path: Path, seven_zip: Path | None = None) -> dict[str, int]:
    """就地规范化 docx（主要面向 LibreOffice 转换产物）。"""
    docx_path = Path(docx_path)

    with zipfile.ZipFile(docx_path) as z:
        names = set(z.namelist())
        doc_xml = z.read("word/document.xml").decode("utf-8")
        styles_xml = z.read("word/styles.xml").decode("utf-8")
        font_table: str | None = (
            z.read("word/fontTable.xml").decode("utf-8")
            if "word/fontTable.xml" in names
            else None
        )

    stats = {"paragraph_mark": 0, "thin_space": 0, "font_swap": 0}

    # 2) 空白填充折叠："空格+六分空" 对 -> 单个空格
    prev = -1
    while prev != doc_xml.count("\u2006"):
        prev = doc_xml.count("\u2006")
        doc_xml, n = re.subn(r" ?\u2006 ?", " ", doc_xml)
        stats["thin_space"] += n

    # 1) 段落标记字号补齐
    counter = [0]
    doc_xml = _P_BLOCK.sub(lambda m: _fix_paragraph_mark(m.group(0), counter), doc_xml)
    stats["paragraph_mark"] = counter[0]

    # 3) CJK 字体名恢复
    replacements: dict[str, bytes] = {}
    parts = [("word/document.xml", doc_xml), ("word/styles.xml", styles_xml)]
    if font_table is not None:
        parts.append(("word/fontTable.xml", font_table))
    for name, text in parts:
        for old, new in _CJK_FONT_SWAP.items():
            if old in text:
                n = text.count(old)
                text = text.replace(old, new)
                stats["font_swap"] += n
        replacements[name] = text.encode("utf-8")

    rewrite_zip(docx_path=docx_path, replacements=replacements, seven_zip=seven_zip)
    logger.info("DOCX 规范化完成 %s: %s", docx_path.name, stats)
    return stats
