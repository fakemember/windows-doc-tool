"""
测试用例 docx 构造工厂。
无需 python-docx，直接用 zipfile 与 OOXML 最小规范构造合法 docx 文档。
"""

from __future__ import annotations

import io
import zipfile
import html
from pathlib import Path
from typing import Any, Mapping, Sequence


DEFAULT_STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal" w:default="1">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:pPr><w:outlineLvl w:val="0"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="DengXian" w:eastAsia="等线"/><w:b/><w:sz w:val="32"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:pPr><w:outlineLvl w:val="1"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="DengXian" w:eastAsia="等线"/><w:b/><w:sz w:val="28"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:pPr><w:outlineLvl w:val="2"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="DengXian" w:eastAsia="等线"/><w:b/><w:sz w:val="24"/></w:rPr>
  </w:style>
</w:styles>
"""

DEFAULT_SETTINGS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
"""

DEFAULT_CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>
"""

DEFAULT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

DEFAULT_DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>
"""

DEFAULT_BODY_SECTPR = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>'


def build_docx(
    path: Path,
    paragraphs: Sequence[Mapping[str, Any]],
    *,
    styles_xml: str | bytes | None = None,
    settings_xml: str | bytes | None = None,
    body_sect_pr: str | None = DEFAULT_BODY_SECTPR,
    extra_files: Mapping[str, bytes] | None = None,
) -> Path:
    """构建最简合法 docx 文件。

    :param path: 输出文件路径
    :param paragraphs: 段落规范字典列表
        支持键值：
        - text: 段落文本
        - outline: 大纲级别 int (0..8)
        - pstyle: 样式 styleId str
        - page_break: True 时在段首插入硬分页符
        - sect_pr: 段落级 sectPr xml 片段
        - bookmark: (id: int, name: str) 段落内书签
        - raw_xml: 完整的 <w:p>...</w:p> 原始 XML，指定时忽略其他字段
    :param styles_xml: 自定义 styles.xml
    :param settings_xml: 自定义 settings.xml
    :param body_sect_pr: body 末尾 sectPr
    :param extra_files: 注入压缩包的其他相对路径文件
    """
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    # 构造 body 段落
    p_elements: list[str] = []
    for p in paragraphs:
        if "raw_xml" in p:
            p_elements.append(p["raw_xml"])
            continue

        p_pr_children: list[str] = []
        if "pstyle" in p:
            p_pr_children.append(f'<w:pStyle w:val="{html.escape(str(p["pstyle"]))}"/>')
        if "outline" in p:
            p_pr_children.append(f'<w:outlineLvl w:val="{int(p["outline"])}"/>')
        if "sect_pr" in p and p["sect_pr"]:
            p_pr_children.append(p["sect_pr"])

        p_pr_xml = f"<w:pPr>{''.join(p_pr_children)}</w:pPr>" if p_pr_children else ""

        runs: list[str] = []
        if p.get("page_break"):
            runs.append('<w:r><w:br w:type="page"/></w:r>')

        bm_start = ""
        bm_end = ""
        if "bookmark" in p and p["bookmark"]:
            bm_id, bm_name = p["bookmark"]
            bm_start = f'<w:bookmarkStart w:id="{bm_id}" w:name="{html.escape(bm_name)}"/>'
            bm_end = f'<w:bookmarkEnd w:id="{bm_id}"/>'

        text = p.get("text", "")
        if text:
            escaped_text = html.escape(str(text))
            runs.append(f'<w:r><w:t xml:space="preserve">{escaped_text}</w:t></w:r>')

        p_xml = f"<w:p>{p_pr_xml}{bm_start}{''.join(runs)}{bm_end}</w:p>"
        p_elements.append(p_xml)

    body_content = "".join(p_elements)
    if body_sect_pr:
        body_content += body_sect_pr

    doc_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n'
        f"  <w:body>{body_content}</w:body>\n"
        "</w:document>"
    )

    styles_data: bytes
    if styles_xml is None:
        styles_data = DEFAULT_STYLES_XML.encode("utf-8")
    elif isinstance(styles_xml, str):
        styles_data = styles_xml.encode("utf-8")
    else:
        styles_data = styles_xml

    settings_data: bytes
    if settings_xml is None:
        settings_data = DEFAULT_SETTINGS_XML.encode("utf-8")
    elif isinstance(settings_xml, str):
        settings_data = settings_xml.encode("utf-8")
    else:
        settings_data = settings_xml

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", DEFAULT_CONTENT_TYPES.strip().encode("utf-8"))
        z.writestr("_rels/.rels", DEFAULT_RELS.strip().encode("utf-8"))
        z.writestr("word/_rels/document.xml.rels", DEFAULT_DOC_RELS.strip().encode("utf-8"))
        z.writestr("word/document.xml", doc_xml.encode("utf-8"))
        z.writestr("word/styles.xml", styles_data)
        z.writestr("word/settings.xml", settings_data)

        if extra_files:
            for k, v in extra_files.items():
                z.writestr(k, v)

    return path
