"""
实测页码测量与回写单元测试：覆盖隐藏锚点提取、缺失回退、PAGEREF 精确回写与清理。
"""

import zipfile
from pathlib import Path
from unittest.mock import MagicMock, patch

from wdt.toc.page_measurer import (
    build_heading_anchor_xml,
    build_measurement_copy,
    build_sentinel_xml,
    measure_pages_from_pdf,
    patch_docx_measured_pages,
)


def test_build_measurement_copy(tmp_path: Path):
    """测量副本应剥除页眉/页脚引用、清零其距离并将上边距缩减 226 twips"""
    src = tmp_path / "src.docx"
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:body>'
        '<w:p><w:r><w:t>hi</w:t></w:r></w:p>'
        '<w:sectPr>'
        '<w:footerReference w:type="default" r:id="rId3"/>'
        '<w:headerReference w:type="first" r:id="rId4"/>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="432" w:right="1296" w:bottom="777" w:left="1296" w:header="0" w:footer="720"/>'
        '</w:sectPr>'
        '</w:body></w:document>'
    )
    with zipfile.ZipFile(src, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("word/document.xml", xml)
        z.writestr("word/other.xml", "<x/>")

    dst = tmp_path / "dst.docx"
    build_measurement_copy(src, dst)

    with zipfile.ZipFile(dst) as z:
        out = z.read("word/document.xml").decode("utf-8")
        assert z.read("word/other.xml") == b"<x/>"

    assert "footerReference" not in out
    assert "headerReference" not in out
    assert 'w:footer="0"' in out
    assert 'w:top="206"' in out  # 432 - 226
    assert "<w:t>hi</w:t>" in out  # 正文内容不变


def test_measure_pages_from_pdf_success(tmp_path: Path):
    pdf_file = tmp_path / "test.pdf"
    pdf_file.write_text("dummy", encoding="utf-8")

    page1 = MagicMock()
    page1.extract_text.return_value = "Table of Contents"
    page2 = MagicMock()
    page2.extract_text.return_value = "WDTSENTINEL12345678 Introduction WDTTAG0000"
    page3 = MagicMock()
    page3.extract_text.return_value = "Methods and Results WDTTAG0001"

    with patch("wdt.toc.page_measurer.PdfReader") as mock_reader_cls:
        mock_reader = MagicMock()
        mock_reader.pages = [page1, page2, page3]
        mock_reader_cls.return_value = mock_reader

        pages = measure_pages_from_pdf(
            pdf_path=pdf_file,
            heading_count=2,
            sentinel_id="12345678",
            page_offset=0,
        )
        # 第 2 页是正文首页 => k = 2 - 1 = 1
        # WDTTAG0000 位于物理第 2 页 => 显示页码 = 2 - 1 = 1
        # WDTTAG0001 位于物理第 3 页 => 显示页码 = 3 - 1 = 2
        assert pages == [1, 2]


def test_measure_pages_missing_anchor_returns_none(tmp_path: Path):
    pdf_file = tmp_path / "test_missing.pdf"
    pdf_file.write_text("dummy", encoding="utf-8")

    page1 = MagicMock()
    page1.extract_text.return_value = "Table of Contents"
    page2 = MagicMock()
    page2.extract_text.return_value = "WDTSENTINEL12345678 Introduction WDTTAG0000"
    page3 = MagicMock()
    page3.extract_text.return_value = "No tag here"

    with patch("wdt.toc.page_measurer.PdfReader") as mock_reader_cls:
        mock_reader = MagicMock()
        mock_reader.pages = [page1, page2, page3]
        mock_reader_cls.return_value = mock_reader

        pages = measure_pages_from_pdf(
            pdf_path=pdf_file,
            heading_count=2,
            sentinel_id="12345678",
        )
        assert pages is None


def test_patch_docx_measured_pages(tmp_path: Path):
    docx_file = tmp_path / "test_patch.docx"
    doc_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:body>'
        '<w:p><w:hyperlink w:anchor="_WDToc000001">'
        '<w:r><w:t>Heading 1</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText> PAGEREF _WDToc000001 \\h </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        '<w:r><w:t>0</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        '</w:hyperlink></w:p>'
        # Sentinel 段落
        '<w:p><w:r><w:t>WDTSENTINELabcd1234</w:t></w:r></w:p>'
        # 正文标题段落，含锚点 run
        '<w:p><w:bookmarkStart w:id="1" w:name="_WDToc000001"/>'
        '<w:r><w:rPr><w:sz w:val="2"/><w:color w:val="FFFFFF"/></w:rPr><w:t>WDTTAG0000</w:t></w:r>'
        '<w:r><w:t>Heading 1 Content</w:t></w:r>'
        '<w:bookmarkEnd w:id="1"/></w:p>'
        '</w:body></w:document>'
    )

    with zipfile.ZipFile(docx_file, "w") as z:
        z.writestr("word/document.xml", doc_xml.encode("utf-8"))

    # 回写实测页码 15
    patch_docx_measured_pages(
        docx_path=docx_file,
        measured_pages=[15],
        bookmark_names=["_WDToc000001"],
        sentinel_id="abcd1234",
    )

    with zipfile.ZipFile(docx_file, "r") as z:
        patched_xml = z.read("word/document.xml").decode("utf-8")

    assert "<w:t>15</w:t>" in patched_xml
    assert "<w:t>0</w:t>" not in patched_xml
    assert "WDTTAG0000" not in patched_xml
    assert "WDTSENTINEL" not in patched_xml
    assert "Heading 1 Content" in patched_xml


def test_patch_docx_measured_pages_static_hyperlink(tmp_path: Path):
    docx_file = tmp_path / "test_patch_static.docx"
    doc_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:body>'
        '<w:p><w:hyperlink w:anchor="_TocHeading_001" w:history="1">'
        '<w:r><w:t>Heading 1</w:t></w:r>'
        '<w:r><w:tab/></w:r>'
        '<w:r><w:t>1</w:t></w:r>'
        '</w:hyperlink></w:p>'
        # Sentinel 段落
        '<w:p><w:r><w:t>WDTSENTINELxyz987</w:t></w:r></w:p>'
        # 正文标题段落，含锚点 run
        '<w:p><w:bookmarkStart w:id="1" w:name="_TocHeading_001"/>'
        '<w:r><w:rPr><w:sz w:val="2"/><w:color w:val="FFFFFF"/></w:rPr><w:t>WDTTAG0000</w:t></w:r>'
        '<w:r><w:t>Heading 1 Content</w:t></w:r>'
        '<w:bookmarkEnd w:id="1"/></w:p>'
        '</w:body></w:document>'
    )

    with zipfile.ZipFile(docx_file, "w") as z:
        z.writestr("word/document.xml", doc_xml.encode("utf-8"))

    # 回写实测页码 42
    patch_docx_measured_pages(
        docx_path=docx_file,
        measured_pages=[42],
        bookmark_names=["_TocHeading_001"],
        sentinel_id="xyz987",
    )

    with zipfile.ZipFile(docx_file, "r") as z:
        patched_xml = z.read("word/document.xml").decode("utf-8")

    assert "<w:t>42</w:t>" in patched_xml
    assert "<w:t>1</w:t>" not in patched_xml
    assert "WDTTAG0000" not in patched_xml
    assert "WDTSENTINEL" not in patched_xml
    assert "Heading 1 Content" in patched_xml
