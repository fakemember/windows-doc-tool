"""
用于动态生成独立分卷目录页的 DOCX/PDF 渲染模块。
利用 Word COM / OOXML 引擎生成 100% 具备本地矢量字体子集嵌入、支持中文、带点导线的专业级目录页。
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

        p_pr_xml = f"<w:pPr>{''.join(p_pr_children)}</w:pPr>" if p_pr_children else ""
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


def create_toc_pages_batch(
    toc_specs: Sequence[dict],
) -> list[bytes]:
    """批量生成所有分卷目录页 PDF 字节流。
    
    单次启动 WordSession 实例批量转换，并在临时目录下安全保存，彻底规避反复开关 COM RPC。
    """
    if not toc_specs:
        return []

    # 临床中文字体样式配置（全量使用微软默认且必带的 微软雅黑 + Calibri）
    styles_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal" w:default="1">
    <w:name w:val="Normal"/>
    <w:rPr>
      <w:rFonts w:ascii="Calibri" w:eastAsia="微软雅黑" w:hAnsi="Calibri" w:cs="Calibri"/>
      <w:sz w:val="20"/>
      <w:szCs w:val="20"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:pPr><w:jc w:val="center"/><w:spacing w:before="180" w:after="240"/></w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Calibri" w:eastAsia="微软雅黑" w:hAnsi="Calibri" w:cs="Calibri"/>
      <w:b/><w:sz w:val="32"/><w:szCs w:val="32"/><w:color w:val="1F4E79"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="TocItem">
    <w:name w:val="TocItem"/>
    <w:pPr><w:spacing w:line="280" w:lineRule="auto"/></w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Calibri" w:eastAsia="微软雅黑" w:hAnsi="Calibri" w:cs="Calibri"/>
      <w:sz w:val="19"/><w:szCs w:val="19"/><w:color w:val="262626"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="FooterNote">
    <w:name w:val="FooterNote"/>
    <w:pPr><w:spacing w:before="160" w:after="40"/></w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Calibri" w:eastAsia="微软雅黑" w:hAnsi="Calibri" w:cs="Calibri"/>
      <w:sz w:val="17"/><w:szCs w:val="17"/><w:color w:val="7F7F7F"/>
    </w:rPr>
  </w:style>
</w:styles>"""

    results: list[bytes] = [b""] * len(toc_specs)
    if not is_wincom_supported():
        return results

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_base = Path(tmp_dir)
        docx_pdf_pairs = []

        for idx, spec in enumerate(toc_specs):
            paragraphs = []
            # 顶部仅保留优雅简洁的中文大标题，居中显示
            paragraphs.append({"text": "目    录", "pstyle": "Heading1"})

            items = spec["items"]
            max_items = 24
            for title, page_num, lvl in items[:max_items]:
                indent = "    " * (max(1, min(3, lvl)) - 1)
                paragraphs.append({
                    "text": f"{indent}{title}  ......  {page_num}",
                    "pstyle": "TocItem",
                })

            if len(items) > max_items:
                omitted = len(items) - max_items
                paragraphs.append({
                    "text": f"[+ 尚有 {omitted} 个子表未在概览页展开...]",
                    "pstyle": "TocItem",
                })

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
