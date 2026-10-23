"""
tests/toc/test_docx_stream_patcher.py
综合验证 TOC 目录注入引擎（docx_stream_patcher.py）：
4种样式注入与 XML well-formed、书签避让、动态/静态模式、分节重启、占位符替换等。
"""

import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import pytest

from tests.fixtures_factory import build_docx
from wdt.toc.docx_stream_patcher import (
    NoHeadingError,
    Position,
    TocConfig,
    TocReport,
    inject_toc,
)
from wdt.toc.page_calculator import PageMode
from wdt.toc.styles_templates import TocStyle

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def _read_docx_xmls(docx_path: Path) -> tuple[str, str, str]:
    with zipfile.ZipFile(docx_path, "r") as z:
        doc = z.read("word/document.xml").decode("utf-8")
        styles = z.read("word/styles.xml").decode("utf-8")
        settings = (
            z.read("word/settings.xml").decode("utf-8")
            if "word/settings.xml" in z.namelist()
            else ""
        )
    return doc, styles, settings


def test_no_heading_error(tmp_path: Path):
    """当文档中无符合深度条件的标题时抛出 NoHeadingError"""
    docx_file = tmp_path / "empty_headings.docx"
    build_docx(docx_file, [{"text": "只有普通段落"}])

    config = TocConfig(depth=1)
    with pytest.raises(NoHeadingError):
        inject_toc(docx_file, config)


@pytest.mark.parametrize("style", list(TocStyle))
def test_all_styles_injection_and_well_formed(tmp_path: Path, style: TocStyle):
    """验证 4 种样式注入后 document.xml 和 styles.xml 均为 well-formed OOXML"""
    docx_file = tmp_path / f"test_style_{style.value}.docx"
    build_docx(
        docx_file,
        [
            {"text": "第一章 引言", "pstyle": "Heading1"},
            {"text": "1.1 背景", "pstyle": "Heading2"},
            {"text": "第二章 方法", "pstyle": "Heading1"},
        ],
    )

    config = TocConfig(
        style=style,
        depth=2,
        title="目 录",
        position=Position.DOCUMENT_START,
        page_mode=PageMode.STATIC_RESTART,
    )
    report = inject_toc(docx_file, config)

    assert report.entries == 3
    assert len(report.pages) == 3

    doc_str, styles_str, _ = _read_docx_xmls(docx_file)

    # 校验 XML 合法性 (well-formed)
    doc_root = ET.fromstring(doc_str)
    styles_root = ET.fromstring(styles_str)

    # styles.xml 中注入了 TOC1, TOC2, TOC, a7
    style_ids = [s.attrib.get(f"{{{W_NS}}}styleId") for s in styles_root.findall("w:style", NS)]
    assert "TOC" in style_ids
    assert "TOC1" in style_ids
    assert "TOC2" in style_ids
    assert "a7" in style_ids

    # document.xml 中包含超链接以及书签
    hyperlinks = doc_root.findall(".//w:hyperlink", NS)
    assert len(hyperlinks) == 3


def test_bookmark_collision_avoidance(tmp_path: Path):
    """验证书签避让：现有高 id 书签（如 id=900）不冲突，现有 _WDToc 前缀自动换名"""
    docx_file = tmp_path / "test_bm_avoid.docx"
    build_docx(
        docx_file,
        [
            {"text": "预设高ID段落", "bookmark": (900, "_WDToc000001")},
            {"text": "新标题 1", "pstyle": "Heading1"},
            {"text": "新标题 2", "pstyle": "Heading1"},
        ],
    )

    config = TocConfig(depth=1, position=Position.DOCUMENT_START)
    report = inject_toc(docx_file, config)
    assert report.entries == 2

    doc_str, _, _ = _read_docx_xmls(docx_file)
    doc_root = ET.fromstring(doc_str)

    bm_starts = doc_root.findall(".//w:bookmarkStart", NS)
    bm_ids = [int(bm.attrib[f"{{{W_NS}}}id"]) for bm in bm_starts]
    bm_names = [bm.attrib[f"{{{W_NS}}}name"] for bm in bm_starts]

    # 原有的 900 与 _WDToc000001 必须保留
    assert 900 in bm_ids
    assert "_WDToc000001" in bm_names

    # 新书签 id 必须严格大于 900
    new_ids = [i for i in bm_ids if i != 900]
    assert all(i > 900 for i in new_ids)

    # 新书签前缀自动升级（不与 _WDToc 重叠，改为 _WDTocX...）
    new_names = [n for n in bm_names if n != "_WDToc000001"]
    assert all(n.startswith("_WDTocX") for n in new_names)


def test_dynamic_vs_static_update_fields(tmp_path: Path):
    """验证动态模式写入 updateFields，静态模式清除 updateFields"""
    docx_file = tmp_path / "test_dynamic.docx"
    build_docx(
        docx_file,
        [{"text": "标题 1", "pstyle": "Heading1"}],
        settings_xml='<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
    )

    # 1. 动态模式
    config_dyn = TocConfig(page_mode=PageMode.DYNAMIC)
    inject_toc(docx_file, config_dyn)
    _, _, settings_dyn = _read_docx_xmls(docx_file)
    assert 'w:updateFields w:val="true"' in settings_dyn

    # 2. 静态模式清洗
    config_stat = TocConfig(page_mode=PageMode.STATIC_RESTART)
    inject_toc(docx_file, config_stat)
    _, _, settings_stat = _read_docx_xmls(docx_file)
    assert "updateFields" not in settings_stat


def test_replace_placeholder_mode(tmp_path: Path):
    """验证 REPLACE_PLACEHOLDER 模式原位替换占位符区域"""
    docx_file = tmp_path / "test_placeholder.docx"
    build_docx(
        docx_file,
        [
            {"text": "封面文档"},
            {"page_break": True},
            {"text": "Table of Contents"},
            {"text": "中心空段落"},
            {"text": "From Microsoft Word"},
            {"page_break": True},
            {"text": "正文标题一", "outline": 0},
            {"page_break": True},
            {"text": "正文标题二", "outline": 0},
        ],
    )

    config = TocConfig(
        position=Position.REPLACE_PLACEHOLDER,
        page_mode=PageMode.STATIC_IN_PLACE,
    )
    report = inject_toc(docx_file, config)
    assert report.entries == 2

    doc_str, _, _ = _read_docx_xmls(docx_file)
    # 占位符 From Microsoft Word 已被替换去除
    assert "From Microsoft Word" not in doc_str
    # 标题正文仍在
    assert "正文标题一" in doc_str
    assert "正文标题二" in doc_str

    # 验证 document.xml 结构合法
    ET.fromstring(doc_str)


def test_restart_section_and_pgnumtype(tmp_path: Path):
    """
    验证 STATIC_RESTART 分节：
    注入后 document.xml 出现两个 sectPr 且正文 sectPr 带 pgNumType w:start="1"
    """
    docx_file = tmp_path / "test_restart_sect.docx"
    build_docx(
        docx_file,
        [
            {"text": "前言"},
            {"text": "章节一", "pstyle": "Heading1"},
            {"text": "章节二", "pstyle": "Heading1"},
        ],
    )

    config = TocConfig(
        position=Position.DOCUMENT_START,
        page_mode=PageMode.STATIC_RESTART,
    )
    inject_toc(docx_file, config)

    doc_str, _, _ = _read_docx_xmls(docx_file)
    doc_root = ET.fromstring(doc_str)

    # 必须出现 2 个 sectPr：目录独立节 + 正文节
    sect_prs = doc_root.findall(".//w:sectPr", NS)
    assert len(sect_prs) == 2

    # 正文 sectPr 必须带有 pgNumType start="1"
    body_sect = sect_prs[-1]
    pg_num = body_sect.find("w:pgNumType", NS)
    assert pg_num is not None
    assert pg_num.attrib.get(f"{{{W_NS}}}start") == "1"


def test_position_before_first_heading_and_bookmark(tmp_path: Path):
    """验证 BEFORE_FIRST_HEADING 与 BOOKMARK 插入位置"""
    docx_file = tmp_path / "test_positions.docx"
    build_docx(
        docx_file,
        [
            {"text": "引言段落", "bookmark": (10, "TargetBM")},
            {"text": "第一章", "pstyle": "Heading1"},
        ],
    )

    # 1. 验证 BOOKMARK 位置插入
    config_bm = TocConfig(
        position=Position.BOOKMARK,
        bookmark_name="TargetBM",
        title="目录书签后",
    )
    inject_toc(docx_file, config_bm)
    doc_str, _, _ = _read_docx_xmls(docx_file)
    ET.fromstring(doc_str)
    assert "目录书签后" in doc_str

    # 2. 验证 BEFORE_FIRST_HEADING 位置插入
    docx_file2 = tmp_path / "test_pos_heading.docx"
    build_docx(
        docx_file2,
        [
            {"text": "封面说明"},
            {"text": "核心标题", "pstyle": "Heading1"},
        ],
    )
    config_heading = TocConfig(
        position=Position.BEFORE_FIRST_HEADING,
        title="标题前目录",
    )
    inject_toc(docx_file2, config_heading)
    doc_str2, _, _ = _read_docx_xmls(docx_file2)
    ET.fromstring(doc_str2)
    assert "标题前目录" in doc_str2


def test_backup_behavior(tmp_path: Path):
    """验证 backup 参数控制 .bak 自动备份行为"""
    docx_file = tmp_path / "test_backup.docx"
    build_docx(docx_file, [{"text": "备份测试标题", "pstyle": "Heading1"}])

    # 1. 默认 backup=True 生成 .bak
    config_backup = TocConfig(backup=True)
    report = inject_toc(docx_file, config_backup)
    bak_path = docx_file.with_name(docx_file.name + ".bak")
    assert bak_path.exists()
    assert report.backup_path == bak_path

    # 2. backup=False 不生成新备份
    docx_file2 = tmp_path / "test_no_backup.docx"
    build_docx(docx_file2, [{"text": "无备份标题", "pstyle": "Heading1"}])
    config_no_bak = TocConfig(backup=False)
    report2 = inject_toc(docx_file2, config_no_bak)
    bak_path2 = docx_file2.with_name(docx_file2.name + ".bak")
    assert not bak_path2.exists()
    assert report2.backup_path is None

