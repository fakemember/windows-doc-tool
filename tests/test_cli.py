"""
CLI 端到端集成测试用例。
"""

from __future__ import annotations

import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import pytest
from pypdf import PdfReader

from tests.fixtures_factory import build_docx
from wdt.core.engine_finder import EngineNotFoundError, find_engine
from wdt.main import main

try:
    find_engine()
    _HAS_SOFFICE = True
except Exception:
    _HAS_SOFFICE = False


def test_main_no_args(capsys):
    """测试无参数调用打印帮助信息并退出 0。"""
    ret = main([])
    assert ret == 0
    captured = capsys.readouterr()
    assert "usage: wdt" in captured.out or "Windows Doc Tool" in captured.out


def test_detect_cmd(capsys):
    """测试 detect 子命令环境诊断正常运行。"""
    ret = main(["detect"])
    captured = capsys.readouterr()
    assert "Windows Doc Tool (wdt) - 环境诊断" in captured.out
    if _HAS_SOFFICE:
        assert ret == 0
        assert "已就绪" in captured.out
    else:
        assert ret == 1
        assert "未找到可用的 LibreOffice" in captured.out


def test_detect_no_engine(monkeypatch, capsys):
    """测试 detect 在未找到引擎时的错误提示与退出码 1。"""
    def _mock_find_engine(**kwargs):
        raise EngineNotFoundError("LibreOffice not found")

    monkeypatch.setattr("wdt.main.find_engine", _mock_find_engine)
    ret = main(["detect"])
    assert ret == 1
    captured = capsys.readouterr()
    assert "未找到可用的 LibreOffice" in captured.out
    assert "安装提示" in captured.out


def test_toc_e2e(tmp_path: Path):
    """测试 toc 子命令端到端注入：生成备份、解析 document.xml、校验 _WDToc 书签。"""
    doc_path = tmp_path / "report.docx"
    build_docx(
        doc_path,
        paragraphs=[
            {"text": "Chapter 1 Introduction", "outline": 0, "pstyle": "Heading1"},
            {"text": "Paragraph 1 text content"},
            {"text": "Section 1.1 Background", "outline": 1, "pstyle": "Heading2"},
            {"text": "Paragraph 2 text content"},
            {"text": "Chapter 2 Methods", "outline": 0, "pstyle": "Heading1", "page_break": True},
            {"text": "Paragraph 3 text content"},
            {"text": "Chapter 3 Results", "outline": 0, "pstyle": "Heading1", "page_break": True},
            {"text": "Paragraph 4 text content"},
        ],
    )

    ret = main(["toc", str(doc_path), "--style", "modern", "--depth", "2"])
    assert ret == 0

    # 1. 断言 .bak 备份文件生成
    bak_path = doc_path.with_name(doc_path.name + ".bak")
    assert bak_path.is_file()

    # 2. 断言注入后 document.xml 可正常解析且含 _WDToc 书签
    with zipfile.ZipFile(doc_path, "r") as z:
        xml_bytes = z.read("word/document.xml")

    root = ET.fromstring(xml_bytes)
    assert root is not None
    assert b"_WDToc" in xml_bytes


def test_toc_bookmark_position_and_no_backup(tmp_path: Path):
    """测试 toc bookmark:NAME 插入位置与 --no-backup 选项。"""
    doc_path = tmp_path / "doc_bm.docx"
    build_docx(
        doc_path,
        paragraphs=[
            {"text": "TOC Placeholder", "bookmark": (10, "MY_TOC_POS")},
            {"text": "Heading One", "outline": 0, "pstyle": "Heading1"},
        ],
    )

    ret = main(
        [
            "toc",
            str(doc_path),
            "--position",
            "bookmark:MY_TOC_POS",
            "--no-backup",
        ]
    )
    assert ret == 0

    # 断言未生成备份
    bak_path = doc_path.with_name(doc_path.name + ".bak")
    assert not bak_path.exists()

    with zipfile.ZipFile(doc_path, "r") as z:
        xml_bytes = z.read("word/document.xml")
    assert b"_WDToc" in xml_bytes


@pytest.mark.skipif(not _HAS_SOFFICE, reason="系统未安装 soffice")
def test_run_pipeline(tmp_path: Path):
    """测试 run 一体化流水线：从 docx 注入目录并导出 PDF。"""
    doc_path = tmp_path / "pipeline_doc.docx"
    build_docx(
        doc_path,
        paragraphs=[
            {"text": "Pipeline Heading 1", "outline": 0, "pstyle": "Heading1"},
            {"text": "Body content on first page"},
            {"text": "Pipeline Heading 2", "outline": 0, "pstyle": "Heading1", "page_break": True},
            {"text": "Body content on second page"},
        ],
    )

    out_dir = tmp_path / "out_pipeline"
    ret = main(["run", str(doc_path), "-o", str(out_dir), "--pdf"])
    assert ret == 0

    stem_docx = out_dir / "pipeline_doc.docx"
    stem_pdf = out_dir / "pipeline_doc.pdf"

    assert stem_docx.is_file()
    assert stem_pdf.is_file()
    assert stem_pdf.stat().st_size > 0

    with zipfile.ZipFile(stem_docx, "r") as z:
        xml_bytes = z.read("word/document.xml")
    assert b"_WDToc" in xml_bytes


def test_run_pdf_dynamic_conflict(tmp_path: Path):
    """测试 run 命令下 --pdf 与 --page-mode dynamic 互斥约束返回非零。"""
    doc_path = tmp_path / "conflict.docx"
    build_docx(doc_path, paragraphs=[{"text": "H1", "outline": 0, "pstyle": "Heading1"}])

    ret = main(["run", str(doc_path), "--pdf", "--page-mode", "dynamic"])
    assert ret != 0


@pytest.mark.skipif(not _HAS_SOFFICE, reason="系统未安装 soffice")
def test_convert_pages_slice(tmp_path: Path):
    """测试 convert --pages 1-1 单文件截取输出 _p1.pdf。"""
    doc_path = tmp_path / "report_pages.docx"
    build_docx(
        doc_path,
        paragraphs=[
            {"text": "First Page Content"},
            {"text": "Second Page Content", "page_break": True},
        ],
    )

    out_dir = tmp_path / "out_slice"
    ret = main(["convert", str(doc_path), "-o", str(out_dir), "--pages", "1-1"])
    assert ret == 0

    expected_pdf = out_dir / "report_pages_p1.pdf"
    assert expected_pdf.is_file()
    assert expected_pdf.stat().st_size > 0

    reader = PdfReader(str(expected_pdf))
    assert len(reader.pages) == 1


def test_cli_argument_errors(tmp_path: Path):
    """测试各种参数错误场景返回非零且不抛出调用栈。"""
    # 1. 目标文件不存在
    ret = main(["convert", str(tmp_path / "non_existing_file.docx")])
    assert ret != 0

    # 2. 坏 --pages 格式
    valid_doc = tmp_path / "valid.docx"
    build_docx(valid_doc, paragraphs=[{"text": "Some text"}])

    ret = main(["convert", str(valid_doc), "--pages", "bad-format"])
    assert ret != 0

    ret = main(["convert", str(valid_doc), "--pages", "5-2"])
    assert ret != 0

    ret = main(["convert", str(valid_doc), "--pages", "0-5"])
    assert ret != 0

    # 3. 目录模式下使用 --pages
    ret = main(["convert", str(tmp_path), "--pages", "1-2"])
    assert ret != 0

    # 4. toc 无标题文件报错
    no_h_doc = tmp_path / "no_heading.docx"
    build_docx(no_h_doc, paragraphs=[{"text": "No headings here"}])
    ret = main(["toc", str(no_h_doc)])
    assert ret != 0

    # 5. toc bookmark 位置缺少名称
    ret = main(["toc", str(valid_doc), "--position", "bookmark:"])
    assert ret != 0


@pytest.mark.skipif(not _HAS_SOFFICE, reason="系统未安装 soffice")
def test_convert_batch_rtf_to_docx(tmp_path: Path):
    """测试 convert 批量转换 RTF 到 DOCX 目录。"""
    sub_in = tmp_path / "inputs"
    sub_in.mkdir()
    rtf1 = sub_in / "test1.rtf"
    rtf2 = sub_in / "test2.rtf"
    rtf1.write_bytes(b"{\\rtf1\\ansi Test 1 Content\\par}")
    rtf2.write_bytes(b"{\\rtf1\\ansi Test 2 Content\\par}")

    out_dir = tmp_path / "outputs"
    ret = main(["convert", str(sub_in), "-o", str(out_dir), "--to", "docx"])
    assert ret == 0

    assert (out_dir / "test1.docx").is_file()
    assert (out_dir / "test2.docx").is_file()

