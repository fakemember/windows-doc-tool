"""
tests/toc/test_heading_detector.py
验证标题检测器的 pStyle + outlineLvl 双通道识别、排除目录自身、中文样式支持等。
"""

import zipfile
from pathlib import Path

import pytest

from tests.fixtures_factory import build_docx
from wdt.toc.heading_detector import detect_headings, Heading


def _extract_parts(docx_path: Path) -> tuple[bytes, bytes]:
    with zipfile.ZipFile(docx_path, "r") as z:
        return z.read("word/document.xml"), z.read("word/styles.xml")


def test_detect_pstyle_english(tmp_path: Path):
    """验证英文 heading 1/2/3 样式识别"""
    docx_file = tmp_path / "test_pstyle.docx"
    build_docx(
        docx_file,
        [
            {"text": "Intro", "pstyle": "heading1"},
            {"text": "Body 1", "pstyle": "Heading2"},
            {"text": "Sub body", "pstyle": "HEADING3"},
            {"text": "Normal text"},
        ],
    )
    doc_bytes, styles_bytes = _extract_parts(docx_file)
    headings = detect_headings(doc_bytes, styles_bytes)

    assert len(headings) == 3
    assert headings[0].level == 1
    assert headings[0].text == "Intro"
    assert headings[1].level == 2
    assert headings[1].text == "Body 1"
    assert headings[2].level == 3
    assert headings[2].text == "Sub body"


def test_detect_outline_lvl_direct(tmp_path: Path):
    """验证段落直接携带 outlineLvl（如 SAS Listing 输出无 style）"""
    docx_file = tmp_path / "test_outline.docx"
    build_docx(
        docx_file,
        [
            {"text": "Listing 14.1.1", "outline": 0},
            {"text": "Listing 14.1.2", "outline": 1},
            {"text": "Regular note"},
            {"text": "Listing 14.1.3", "outline": 2},
        ],
    )
    doc_bytes, styles_bytes = _extract_parts(docx_file)
    headings = detect_headings(doc_bytes, styles_bytes)

    assert len(headings) == 3
    assert [h.level for h in headings] == [1, 2, 3]
    assert [h.text for h in headings] == ["Listing 14.1.1", "Listing 14.1.2", "Listing 14.1.3"]


def test_detect_hybrid_channels(tmp_path: Path):
    """双通道混合：部分段落使用 styleId，部分段落直接带 outlineLvl"""
    docx_file = tmp_path / "test_hybrid.docx"
    build_docx(
        docx_file,
        [
            {"text": "Chapter 1", "pstyle": "Heading1"},
            {"text": "Section 1.1", "outline": 1},
            {"text": "Normal paragraph"},
            {"text": "Chapter 2", "pstyle": "Heading1"},
        ],
    )
    doc_bytes, styles_bytes = _extract_parts(docx_file)
    headings = detect_headings(doc_bytes, styles_bytes)

    assert len(headings) == 3
    assert [h.level for h in headings] == [1, 2, 1]
    assert [h.text for h in headings] == ["Chapter 1", "Section 1.1", "Chapter 2"]
    # 严格升序
    assert headings[0].p_start < headings[1].p_start < headings[2].p_start


def test_exclude_toc_styles(tmp_path: Path):
    """排除目录自身样式（如 TOC 1, TOC Heading, 目录等），即使含 outlineLvl 也跳过"""
    custom_styles = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="paragraph" w:styleId="Heading1">
        <w:name w:val="heading 1"/>
        <w:pPr><w:outlineLvl w:val="0"/></w:pPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="TOC">
        <w:name w:val="TOC Heading"/>
        <w:pPr><w:outlineLvl w:val="9"/></w:pPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="TOC1">
        <w:name w:val="toc 1"/>
        <w:pPr><w:outlineLvl w:val="0"/></w:pPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="MyMuLu">
        <w:name w:val="目录标题"/>
        <w:pPr><w:outlineLvl w:val="0"/></w:pPr>
      </w:style>
    </w:styles>
    """
    docx_file = tmp_path / "test_toc_exclude.docx"
    build_docx(
        docx_file,
        [
            {"text": "Table of Contents", "pstyle": "TOC"},
            {"text": "Entry 1", "pstyle": "TOC1", "outline": 0},
            {"text": "中文目录项", "pstyle": "MyMuLu"},
            {"text": "Real Heading 1", "pstyle": "Heading1"},
        ],
        styles_xml=custom_styles,
    )
    doc_bytes, styles_bytes = _extract_parts(docx_file)
    headings = detect_headings(doc_bytes, styles_bytes)

    assert len(headings) == 1
    assert headings[0].text == "Real Heading 1"
    assert headings[0].level == 1


def test_chinese_heading_styles(tmp_path: Path):
    """支持中文标题样式名（如 name='标题 1'，或 styleId 为中文）"""
    chinese_styles = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="paragraph" w:styleId="cn_h1">
        <w:name w:val="标题 1"/>
      </w:style>
      <w:style w:type="paragraph" w:styleId="标题2">
        <w:name w:val="标题 2"/>
      </w:style>
      <w:style w:type="paragraph" w:styleId="custom_style">
        <w:name w:val="我的自定义样式"/>
        <w:pPr><w:outlineLvl w:val="2"/></w:pPr>
      </w:style>
    </w:styles>
    """
    docx_file = tmp_path / "test_chinese.docx"
    build_docx(
        docx_file,
        [
            {"text": "第1章 概述", "pstyle": "cn_h1"},
            {"text": "1.1 背景", "pstyle": "标题2"},
            {"text": "1.1.1 细节", "pstyle": "custom_style"},
        ],
        styles_xml=chinese_styles,
    )
    doc_bytes, styles_bytes = _extract_parts(docx_file)
    headings = detect_headings(doc_bytes, styles_bytes)

    assert len(headings) == 3
    assert [h.level for h in headings] == [1, 2, 3]
    assert [h.text for h in headings] == ["第1章 概述", "1.1 背景", "1.1.1 细节"]


def test_text_unescape(tmp_path: Path):
    """测试多 run 拼接与 XML 转义字符还原"""
    docx_file = tmp_path / "test_escape.docx"
    raw_p = (
        '<w:p><w:pPr><w:outlineLvl w:val="0"/></w:pPr>'
        '<w:r><w:t>A &amp; B</w:t></w:r>'
        '<w:r><w:t> &lt;Special&gt;</w:t></w:r></w:p>'
    )
    build_docx(docx_file, [{"raw_xml": raw_p}])
    doc_bytes, styles_bytes = _extract_parts(docx_file)
    headings = detect_headings(doc_bytes, styles_bytes)

    assert len(headings) == 1
    assert headings[0].text == "A & B <Special>"
