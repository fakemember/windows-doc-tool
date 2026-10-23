"""
tests/toc/test_page_calculator.py
验证各页码模式（STATIC_RESTART, STATIC_OFFSET, STATIC_IN_PLACE, DYNAMIC）的计算精度与公式推导结论。
"""

import zipfile
from pathlib import Path

from tests.fixtures_factory import build_docx
from wdt.toc.page_calculator import PageMode, compute_pages


def _get_doc_bytes(docx_path: Path) -> bytes:
    with zipfile.ZipFile(docx_path, "r") as z:
        return z.read("word/document.xml")


def test_page_calculator_dynamic():
    """DYNAMIC 模式直接返回空列表"""
    pages = compute_pages(b"<dummy>", [10, 20], 5, PageMode.DYNAMIC)
    assert pages == []


def test_page_calculator_restart_and_offset(tmp_path: Path):
    """
    构造场景：
    第 1 页：封面
    硬分页符 -> 进入第 2 页
    目录插在第 2 页位置 (insert_pos)
    正文标题 1（正文首段，位于 insert_pos 后）
    硬分页符
    正文标题 2（位于第 2 页）
    硬分页符
    正文标题 3（位于第 3 页）
    """
    docx_file = tmp_path / "test_pages.docx"
    build_docx(
        docx_file,
        [
            {"text": "封面文档"},
            {"page_break": True},  # 换到第 2 页
            {"text": "目录占位段落"},
            {"text": "正文标题 1", "outline": 0},
            {"page_break": True},  # 换页
            {"text": "正文标题 2", "outline": 0},
            {"page_break": True},  # 换页
            {"text": "正文标题 3", "outline": 0},
        ],
    )
    doc_bytes = _get_doc_bytes(docx_file)

    # 找到"目录占位段落"的位置作为 insert_pos
    m_toc_slot = doc_bytes.find("目录占位段落".encode("utf-8"))
    # 找到 3 个标题的起始位置
    m_h1 = doc_bytes.find("正文标题 1".encode("utf-8"))
    m_h2 = doc_bytes.find("正文标题 2".encode("utf-8"))
    m_h3 = doc_bytes.find("正文标题 3".encode("utf-8"))

    p1 = doc_bytes.rfind(b"<w:p", 0, m_h1)
    p2 = doc_bytes.rfind(b"<w:p", 0, m_h2)
    p3 = doc_bytes.rfind(b"<w:p", 0, m_h3)

    insert_pos = doc_bytes.rfind(b"<w:p", 0, m_toc_slot)

    # 1. 验证 STATIC_RESTART：正文首页标题必须为 1，其后每页 +1
    pages_restart = compute_pages(doc_bytes, [p1, p2, p3], insert_pos, PageMode.STATIC_RESTART)
    assert pages_restart == [1, 2, 3]

    # 2. 验证 STATIC_OFFSET：page_offset=3 时，首页为 1 + 3 = 4
    pages_offset = compute_pages(
        doc_bytes, [p1, p2, p3], insert_pos, PageMode.STATIC_OFFSET, page_offset=3
    )
    assert pages_offset == [4, 5, 6]


def test_page_calculator_static_in_place(tmp_path: Path):
    """验证 STATIC_IN_PLACE 模式复刻基座公式：1 + 2 + br_count + (sect_count - 1)"""
    docx_file = tmp_path / "test_in_place.docx"
    build_docx(
        docx_file,
        [
            {"text": "封面"},
            {"page_break": True},
            {"text": "Table of Contents"},
            {"text": "From Microsoft Word"},
            {"page_break": True},
            {"text": "标题一", "outline": 0},
            {"page_break": True},
            {"text": "标题二", "outline": 0},
        ],
    )
    doc_bytes = _get_doc_bytes(docx_file)

    m_h1 = doc_bytes.find("标题一".encode("utf-8"))
    m_h2 = doc_bytes.find("标题二".encode("utf-8"))
    p1 = doc_bytes.rfind(b"<w:p", 0, m_h1)
    p2 = doc_bytes.rfind(b"<w:p", 0, m_h2)

    # p1 之前有 2 个 br (封面后、占位符后)，0 个 sectPr
    # 根据公式: 1 + 2 + br_count + (sect_count - 1)
    # p1: 1 + 2 + 2 + (0 - 1) = 4
    # p2: 1 + 2 + 3 + (0 - 1) = 5
    pages_in_place = compute_pages(doc_bytes, [p1, p2], 0, PageMode.STATIC_IN_PLACE)
    assert pages_in_place == [4, 5]


def test_page_calculator_with_body_sect_pr():
    """测试正文段落间包含 sectPr 的分节换页场景"""
    fake_doc = (
        b"<w:body>"
        b"<w:p><w:t>TOC</w:t></w:p>"  # insert_pos = 10
        b"<w:p><w:t>Heading 1</w:t></w:p>"  # pos = 30
        b"<w:p><w:pPr><w:sectPr/></w:pPr><w:t>End of Sect 1</w:t></w:p>"
        b"<w:p><w:t>Heading 2</w:t></w:p>"  # pos = 110
        b"<w:sectPr/>"
        b"</w:body>"
    )
    pages = compute_pages(fake_doc, [30, 110], 10, PageMode.STATIC_RESTART)
    # Heading 1 前无 sectPr/br: page = 1
    # Heading 2 前跨越了一个正文内部 sectPr: page = 1 + 1 = 2
    assert pages == [1, 2]
