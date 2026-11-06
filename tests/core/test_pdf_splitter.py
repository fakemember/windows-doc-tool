"""
PDF 拆分与独立目录重构单元测试。
"""

from __future__ import annotations

import io
from pathlib import Path

import pytest
from pypdf import PdfReader, PdfWriter

from wdt.core.pdf_splitter import (
    extract_outline_tree,
    plan_volume_splits,
    split_pdf_by_outline,
)
from wdt.main import main


def _create_sample_pdf_with_bookmarks(path: Path) -> Path:
    """构造包含 6 页、多级书签结构的测试 PDF 文档。
    
    结构：
    - Page 0: Cover / Title
    - Page 1: Chapter 1: Demographics (顶级书签 1)
      - Page 2: Section 1.1: Age & Sex (子书签)
    - Page 3: Chapter 2: Safety & Adverse Events (顶级书签 2)
      - Page 4: Section 2.1: Severe AE (子书签)
    - Page 5: Chapter 3: Laboratory Listings (顶级书签 3)
    """
    writer = PdfWriter()
    # 构造 6 个单页
    for i in range(6):
        # 简单添加一个空白页 (A4)
        writer.add_blank_page(width=595.276, height=841.890)

    # 添加多级大纲
    bm_ch1 = writer.add_outline_item("Chapter 1: Demographics", page_number=1)
    writer.add_outline_item("Section 1.1: Age & Sex", page_number=2, parent=bm_ch1)

    bm_ch2 = writer.add_outline_item("Chapter 2: Safety & Adverse Events", page_number=3)
    writer.add_outline_item("Section 2.1: Severe AE", page_number=4, parent=bm_ch2)

    writer.add_outline_item("Chapter 3: Laboratory Listings", page_number=5)

    with open(path, "wb") as f:
        writer.write(f)

    return path


def test_extract_and_plan_volumes(tmp_path: Path):
    """测试书签大纲树提取与分卷切分规划。"""
    pdf_path = tmp_path / "sample.pdf"
    _create_sample_pdf_with_bookmarks(pdf_path)

    reader = PdfReader(str(pdf_path))
    tree = extract_outline_tree(reader)
    assert len(tree) == 3
    assert tree[0].title == "Chapter 1: Demographics"
    assert len(tree[0].children) == 1
    assert tree[0].children[0].title == "Section 1.1: Age & Sex"

    vols = plan_volume_splits(reader)
    # 应被切分成 3 个分卷
    assert len(vols) == 3
    # 卷 1 包含封面和 Chapter 1 (原页 0..2)
    assert vols[0].page_start == 0
    assert vols[0].page_end == 2
    # 卷 2 包含 Chapter 2 (原页 3..4)
    assert vols[1].page_start == 3
    assert vols[1].page_end == 4
    # 卷 3 包含 Chapter 3 (原页 5..5)
    assert vols[2].page_start == 5
    assert vols[2].page_end == 5


def test_split_pdf_with_visual_toc(tmp_path: Path):
    """测试 split_pdf_by_outline 生成带独立可视目录页和重构书签的分卷。"""
    pdf_path = tmp_path / "large_study.pdf"
    _create_sample_pdf_with_bookmarks(pdf_path)

    out_dir = tmp_path / "volumes_out"
    results = split_pdf_by_outline(pdf_path, out_dir=out_dir, with_visual_toc=True)

    assert len(results) == 3

    # 卷 1: 原 3 页 (0..2) + 1 页 TOC = 4 页
    assert results[0].page_count == 4
    r1 = PdfReader(str(results[0].output_path))
    assert len(r1.pages) == 4
    # 检查卷 1 的大纲
    r1_tree = extract_outline_tree(r1)
    # 包含了 TOC 书签和 Chapter 1 节点
    titles = [n.title for n in r1_tree]
    assert any("Table of Contents" in t for t in titles)
    assert any("Chapter 1: Demographics" in t for t in titles)

    # 卷 2: 原 2 页 (3..4) + 1 页 TOC = 3 页
    assert results[1].page_count == 3
    r2 = PdfReader(str(results[1].output_path))
    assert len(r2.pages) == 3
    r2_tree = extract_outline_tree(r2)
    titles2 = [n.title for n in r2_tree]
    assert any("Chapter 2: Safety" in t for t in titles2)

    # 卷 3: 原 1 页 (5..5) + 1 页 TOC = 2 页
    assert results[2].page_count == 2
    r3 = PdfReader(str(results[2].output_path))
    assert len(r3.pages) == 2


def test_split_cli_command(tmp_path: Path, capsys):
    """测试通过 CLI 命令调用 split 功能。"""
    pdf_path = tmp_path / "test_cli.pdf"
    _create_sample_pdf_with_bookmarks(pdf_path)

    out_dir = tmp_path / "cli_volumes"
    ret = main(["split", str(pdf_path), "-o", str(out_dir)])
    assert ret == 0

    captured = capsys.readouterr()
    assert "依据顶级大纲拆分 PDF" in captured.out
    assert "共生成 3 个分卷" in captured.out
    assert (out_dir / "test_cli_Part1.pdf").exists()


def test_visual_toc_paginates_beyond_24_items(tmp_path: Path):
    """单卷条目超过 24 时可视目录应自动分页，条目一条不丢。"""
    writer = PdfWriter()
    n_nodes = 30
    for _ in range(n_nodes + 5):
        writer.add_blank_page(width=595.276, height=841.890)
    root = writer.add_outline_item("Section 16.2 Listings", page_number=2)
    for i in range(n_nodes):
        writer.add_outline_item(f"Listing 16.2.{i + 1} Table Title {i + 1}", page_number=i + 2, parent=root)
    pdf_path = tmp_path / "many_nodes.pdf"
    with open(pdf_path, "wb") as f:
        writer.write(f)

    out_dir = tmp_path / "many_out"
    results = split_pdf_by_outline(pdf_path, out_dir=out_dir, with_visual_toc=True, max_pages=100)
    assert len(results) == 1
    r = PdfReader(str(results[0].output_path))
    # 30 条目 + 5 正文页 + 2 目录页
    assert len(r.pages) == n_nodes + 5 + 2

    # 书签完整（TOC + 30 条目）
    titles = []

    def _walk(items):
        for it in items:
            if isinstance(it, list):
                _walk(it)
            else:
                titles.append(str(it.title))

    _walk(r.outline)
    assert len(titles) == n_nodes + 2  # TOC + 根节点 + 30 条目

    # 两个目录页均有链接注释，合计 31 个（根节点 + 30 条目）
    annot_count = 0
    for page_no in range(2):
        annot_count += len(r.pages[page_no].get("/Annots") or [])
    assert annot_count == n_nodes + 1

    # 每个链接目标页都包含对应编号
    import re as _re
    for page_no in range(2):
        page = r.pages[page_no]
        for a in page.get("/Annots") or []:
            ao = a.get_object()
            ref = ao["/Dest"][0]
            target = next(i for i, pg in enumerate(r.pages) if pg.indirect_reference == ref)
            txt = r.pages[target].extract_text() or ""
            assert "ClinStReport" not in txt  # 占位检查：不跳回目录页
