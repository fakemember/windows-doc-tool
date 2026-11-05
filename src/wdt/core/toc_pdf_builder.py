"""
用于动态生成独立分卷目录页的 DOCX/PDF 渲染模块。
利用 Word COM / OOXML 引擎生成 100% 具备本地矢量字体子集嵌入、支持中文、带专业 Tab Leader 点导线的目录页。
"""

from __future__ import annotations

import html
import os
import tempfile
import time
import zipfile
from pathlib import Path
from typing import Any, Mapping, Sequence

from wdt.core.wincom_worker import WordSession, is_wincom_supported
from wdt.toc.styles_templates import TocStyle


def _build_minimal_docx(
    path: Path,
    paragraphs: Sequence[Mapping[str, Any]],
    styles_xml: str,
) -> Path:
    """构建用于输出目录页的最简合法 DOCX 文件。"""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    p_elements: list[str] = []
    for p in paragraphs:
        p_pr_children: list[str] = []
        if "pstyle" in p:
            p_pr_children.append(f'<w:pStyle w:val="{html.escape(str(p["pstyle"]))}"/>')
        if "outline" in p:
            p_pr_children.append(f'<w:outlineLvl w:val="{int(p["outline"])}"/>')
        if "tabs" in p:
            p_pr_children.append(str(p["tabs"]))

        p_pr_xml = f"<w:pPr>{''.join(p_pr_children)}</w:pPr>" if p_pr_children else ""

        runs = p.get("runs")
        if runs is not None:
            r_parts: list[str] = []
            for r in runs:
                if r.get("tab"):
                    r_parts.append("<w:r><w:tab/></w:r>")
                else:
                    t_text = html.escape(str(r.get("text", "")))
                    rpr_xml = r.get("rpr", "")
                    r_parts.append(f'<w:r>{rpr_xml}<w:t xml:space="preserve">{t_text}</w:t></w:r>')
            run_xml = "".join(r_parts)
        else:
            text = p.get("text", "")
            escaped_text = html.escape(str(text))
            run_xml = f'<w:r><w:t xml:space="preserve">{escaped_text}</w:t></w:r>'
        p_xml = f"<w:p>{p_pr_xml}{run_xml}</w:p>"
        p_elements.append(p_xml)

    body_content = "".join(p_elements)
    # A4 页面设置与页边距 (top/bottom: 1200 dxa ~ 2.1cm, left/right: 1440 dxa ~ 2.5cm)
    body_content += '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1200" w:right="1440" w:bottom="1200" w:left="1440"/></w:sectPr>'

    doc_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n'
        f"  <w:body>{body_content}</w:body>\n"
        "</w:document>"
    )

    types_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""

    rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    doc_rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", types_xml.strip().encode("utf-8"))
        z.writestr("_rels/.rels", rels_xml.strip().encode("utf-8"))
        z.writestr("word/_rels/document.xml.rels", doc_rels_xml.strip().encode("utf-8"))
        z.writestr("word/document.xml", doc_xml.encode("utf-8"))
        z.writestr("word/styles.xml", styles_xml.strip().encode("utf-8"))

    return path


def _get_styles_xml(style: TocStyle) -> str:
    """根据 TocStyle 动态生成 styles.xml 样式配置。"""
    if style == TocStyle.MODERN_ACCENT:
        font_ascii = "Calibri"
        font_ea = "微软雅黑"
        h1_color = "1F4E79"
        h1_jc = "left"
        h1_sz = "30"
        item_color = "1F2937"
        item_sz = "20"
        line_rule = 'w:line="280" w:lineRule="auto"'
    elif style == TocStyle.REGULATORY_LISTING:
        font_ascii = "Times New Roman"
        font_ea = "宋体"
        h1_color = "000000"
        h1_jc = "center"
        h1_sz = "24"
        item_color = "000000"
        item_sz = "21"
        line_rule = 'w:line="240" w:lineRule="auto"'
    elif style == TocStyle.MINIMAL_CLEAN:
        font_ascii = "Arial"
        font_ea = "微软雅黑"
        h1_color = "334155"
        h1_jc = "left"
        h1_sz = "26"
        item_color = "334155"
        item_sz = "20"
        line_rule = 'w:line="270" w:lineRule="auto"'
    else:  # CLASSIC_DOT
        font_ascii = "Times New Roman"
        font_ea = "宋体"
        h1_color = "000000"
        h1_jc = "center"
        h1_sz = "28"
        item_color = "111827"
        item_sz = "21"
        line_rule = 'w:line="260" w:lineRule="auto"'

    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal" w:default="1">
    <w:name w:val="Normal"/>
    <w:rPr>
      <w:rFonts w:ascii="{font_ascii}" w:eastAsia="{font_ea}" w:hAnsi="{font_ascii}" w:cs="{font_ascii}"/>
      <w:sz w:val="{item_sz}"/>
      <w:szCs w:val="{item_sz}"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:pPr><w:jc w:val="{h1_jc}"/><w:spacing w:before="180" w:after="220"/></w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="{font_ascii}" w:eastAsia="{font_ea}" w:hAnsi="{font_ascii}" w:cs="{font_ascii}"/>
      <w:b/><w:sz w:val="{h1_sz}"/><w:szCs w:val="{h1_sz}"/><w:color w:val="{h1_color}"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="SubHeader">
    <w:name w:val="SubHeader"/>
    <w:pPr><w:spacing w:before="0" w:after="140"/></w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="{font_ascii}" w:eastAsia="{font_ea}" w:hAnsi="{font_ascii}" w:cs="{font_ascii}"/>
      <w:sz w:val="18"/><w:szCs w:val="18"/><w:color w:val="64748B"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="TocItem">
    <w:name w:val="TocItem"/>
    <w:pPr><w:spacing {line_rule}/></w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="{font_ascii}" w:eastAsia="{font_ea}" w:hAnsi="{font_ascii}" w:cs="{font_ascii}"/>
      <w:sz w:val="{item_sz}"/><w:szCs w:val="{item_sz}"/><w:color w:val="{item_color}"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="FooterNote">
    <w:name w:val="FooterNote"/>
    <w:pPr><w:spacing w:before="180" w:after="40"/></w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="{font_ascii}" w:eastAsia="{font_ea}" w:hAnsi="{font_ascii}" w:cs="{font_ascii}"/>
      <w:sz w:val="17"/><w:szCs w:val="17"/><w:color w:val="94A3B8"/>
    </w:rPr>
  </w:style>
</w:styles>"""


def create_toc_pages_batch(
    toc_specs: Sequence[dict],
    style: TocStyle = TocStyle.REGULATORY_LISTING,
) -> list[bytes]:
    """批量生成所有分卷目录页 PDF 字节流。

    利用真实 Tab Leader 右对齐制表位实现完美点导线排版，绝无字符串手工拼凑。
    """
    if not toc_specs:
        return []

    styles_xml = _get_styles_xml(style)
    leader_type = "none" if style == TocStyle.MINIMAL_CLEAN else "dot"
    tab_pos = 9026  # A4 可用宽度制表位

    results: list[bytes] = [b""] * len(toc_specs)
    if not is_wincom_supported():
        return results

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_base = Path(tmp_dir)
        docx_pdf_pairs = []

        for idx, spec in enumerate(toc_specs):
            paragraphs = []
            # 顶部大标题
            paragraphs.append({"text": "目    录", "pstyle": "Heading1"})

            # 副标题 / 分卷概要
            v_title = spec.get("vol_title", "")
            if v_title:
                paragraphs.append({"text": v_title, "pstyle": "SubHeader"})

            items = spec["items"]
            max_items = 24
            tab_xml = f'<w:tabs><w:tab w:val="right" w:leader="{leader_type}" w:pos="{tab_pos}"/></w:tabs>'

            for title, page_num, lvl in items[:max_items]:
                indent_str = "    " * (max(1, min(3, lvl)) - 1)
                full_title = f"{indent_str}{title}"
                is_bold = (lvl == 1 and style in (TocStyle.MODERN_ACCENT, TocStyle.REGULATORY_LISTING))
                title_rpr = "<w:rPr><w:b/></w:rPr>" if is_bold else ""
                page_color = "7F7F7F" if style == TocStyle.MINIMAL_CLEAN else "111827"
                page_rpr = f'<w:rPr><w:color w:val="{page_color}"/></w:rPr>'

                paragraphs.append({
                    "pstyle": "TocItem",
                    "tabs": tab_xml,
                    "runs": [
                        {"text": full_title, "rpr": title_rpr},
                        {"tab": True},
                        {"text": str(page_num), "rpr": page_rpr},
                    ],
                })

            if len(items) > max_items:
                omitted = len(items) - max_items
                paragraphs.append({
                    "text": f"[+ 尚有 {omitted} 个条目未在概览页展开...]",
                    "pstyle": "FooterNote",
                })

            if spec.get("footer_text"):
                paragraphs.append({"text": spec["footer_text"], "pstyle": "FooterNote"})

            t_docx = tmp_base / f"toc_{idx}.docx"
            t_pdf = tmp_base / f"toc_{idx}.pdf"
            _build_minimal_docx(t_docx, paragraphs=paragraphs, styles_xml=styles_xml)
            docx_pdf_pairs.append((t_docx, t_pdf))

        session = WordSession()
        try:
            session.start()
            word = session.word
            for idx, (doc_p, pdf_p) in enumerate(docx_pdf_pairs):
                try:
                    d = word.Documents.Open(str(doc_p), ReadOnly=True, ConfirmConversions=False)
                    d.ExportAsFixedFormat(str(pdf_p), 17)
                    d.Close(False)
                    if pdf_p.exists() and pdf_p.stat().st_size > 0:
                        with open(pdf_p, "rb") as f:
                            results[idx] = f.read()
                except Exception:
                    pass
        except Exception:
            pass
        finally:
            session.stop()

    return results
