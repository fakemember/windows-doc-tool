"""
旧目录剥离器单元测试：覆盖 sdt 块、字段段落、前置标题段、孤立换页符吸收及幂等性。
"""

import pytest

from wdt.toc.existing_toc_stripper import strip_existing_toc


def test_strip_sdt_toc_with_heading_and_page_break():
    # 构造含：前置标题段 + SDT旧目录 + 孤立硬分页段 + 正文段
    doc_xml_str = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:body>'
        '<w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr><w:r><w:t>目 录</w:t></w:r></w:p>'
        '<w:sdt>'
        '<w:sdtPr><w:docPartGallery w:val="Table of Contents"/></w:sdtPr>'
        '<w:sdtContent>'
        '<w:p><w:r><w:t>Old TOC Item 1</w:t></w:r></w:p>'
        '</w:sdtContent>'
        '</w:sdt>'
        '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'  # 紧随其后的孤立分页段
        '<w:p><w:r><w:t>Heading 1 Body</w:t></w:r></w:p>'
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>'
        '</w:body></w:document>'
    )
    doc_bytes = doc_xml_str.encode("utf-8")

    styles_xml = (
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:style w:type="paragraph" w:styleId="TOCHeading"><w:name w:val="TOC Heading"/></w:style>'
        '</w:styles>'
    )

    stripped = strip_existing_toc(doc_bytes, styles_xml)

    assert b"<w:sdt>" not in stripped
    assert b"Old TOC Item 1" not in stripped
    assert "目 录".encode("utf-8") not in stripped
    assert b'<w:br w:type="page"/>' not in stripped  # 孤立换页符被吸收
    assert b"Heading 1 Body" in stripped             # 正文完整保留


def test_strip_field_paragraphs_toc():
    # 构造跨段落字段旧目录：TOC \o 段落 -> 条目段落 -> end 闭合段落
    doc_bytes = (
        b'<w:body>'
        b'<w:p><w:r><w:t>Table of Contents</w:t></w:r></w:p>'
        b'<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        b'<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" </w:instrText></w:r>'
        b'<w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>'
        b'<w:p><w:r><w:t>Chapter 1 .... 1</w:t></w:r></w:p>'
        b'<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
        b'<w:p><w:r><w:t>Real Content</w:t></w:r></w:p>'
        b'</w:body>'
    )

    stripped = strip_existing_toc(doc_bytes, "")
    assert b"Table of Contents" not in stripped
    assert b"Chapter 1" not in stripped
    assert b"TOC \\o" not in stripped
    assert b"Real Content" in stripped


def test_preserve_sectpr_paragraph():
    # 删完旧目录后，如果下一个段落包含 sectPr，即使包含换页符也不能删除该段落
    doc_bytes = (
        b'<w:body>'
        b'<w:sdt><w:sdtPr><w:docPartGallery w:val="Table of Contents"/></w:sdtPr>'
        b'<w:sdtContent><w:p><w:r><w:t>Old</w:t></w:r></w:p></w:sdtContent></w:sdt>'
        b'<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr></w:p>'
        b'<w:p><w:r><w:t>After Section</w:t></w:r></w:p>'
        b'</w:body>'
    )

    stripped = strip_existing_toc(doc_bytes, "")
    assert b"<w:sdt>" not in stripped
    assert b"<w:sectPr>" in stripped
    assert b"After Section" in stripped


def test_no_toc_idempotent():
    # 无旧目录原样返回
    doc_bytes = (
        b'<w:body>'
        b'<w:p><w:r><w:t>Hello World</w:t></w:r></w:p>'
        b'</w:body>'
    )
    stripped = strip_existing_toc(doc_bytes, "")
    assert stripped == doc_bytes
