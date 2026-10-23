"""
流水线 (Pipeline) 模块单元与集成测试用例。
"""

from __future__ import annotations

import shutil
import zipfile
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from tests.fixtures_factory import build_docx
from wdt.core.converter_worker import FileResult
from wdt.core.engine_finder import Engine, find_engine
from wdt.pipeline import (
    PipelineEvent,
    PipelineOptions,
    Stage,
    run_pipeline,
)
from wdt.toc.docx_stream_patcher import TocConfig
from wdt.toc.page_calculator import PageMode

_HAS_SOFFICE = shutil.which("soffice") is not None


def test_pipeline_validation_errors(tmp_path: Path):
    """测试 Pipeline 参数配置互斥与缺失校验。"""
    doc_path = tmp_path / "test.docx"
    build_docx(doc_path, paragraphs=[{"text": "Sample", "outline": 0, "pstyle": "Heading1"}])

    # 1. run 模式缺少 toc 配置
    opt_no_toc = PipelineOptions(mode="run", toc=None)
    with pytest.raises(ValueError, match="必须指定 options.toc"):
        run_pipeline([doc_path], opt_no_toc)

    # 2. toc 模式缺少 toc 配置
    opt_toc_no_toc = PipelineOptions(mode="toc", toc=None)
    with pytest.raises(ValueError, match="必须指定 options.toc"):
        run_pipeline([doc_path], opt_toc_no_toc)

    # 3. dynamic + make_pdf 互斥报错
    opt_dynamic_pdf = PipelineOptions(
        mode="run",
        make_pdf=True,
        toc=TocConfig(page_mode=PageMode.DYNAMIC),
    )
    with pytest.raises(ValueError, match="dynamic \\+ make_pdf 组合不支持"):
        run_pipeline([doc_path], opt_dynamic_pdf)

    # 4. convert 模式未选择任何输出格式
    opt_empty_convert = PipelineOptions(
        mode="convert",
        make_docx=False,
        make_pdf=False,
    )
    with pytest.raises(ValueError, match="必须至少启用 make_docx 或 make_pdf"):
        run_pipeline([doc_path], opt_empty_convert)


def test_pipeline_toc_mode_e2e(tmp_path: Path):
    """测试 toc 模式下的注入与事件产生。"""
    doc_path = tmp_path / "toc_test.docx"
    build_docx(
        doc_path,
        paragraphs=[
            {"text": "Heading 1 Intro", "outline": 0, "pstyle": "Heading1"},
            {"text": "Text body 1"},
            {"text": "Heading 2 Details", "outline": 1, "pstyle": "Heading2"},
        ],
    )

    opt = PipelineOptions(
        mode="toc",
        toc=TocConfig(depth=2),
    )
    events: list[PipelineEvent] = []
    results = run_pipeline([doc_path], opt, progress_cb=events.append)

    assert len(results) == 1
    assert results[0].stage == Stage.TOC
    assert results[0].status == "ok"
    assert "条目: 2" in results[0].message
    assert len(events) == 1

    # 验证书签注入
    with zipfile.ZipFile(doc_path, "r") as z:
        xml_bytes = z.read("word/document.xml")
    assert b"_WDToc" in xml_bytes


def test_pipeline_toc_cancellation(tmp_path: Path):
    """测试 toc 模式中途取消。"""
    doc1 = tmp_path / "d1.docx"
    doc2 = tmp_path / "d2.docx"
    build_docx(doc1, paragraphs=[{"text": "H1", "outline": 0, "pstyle": "Heading1"}])
    build_docx(doc2, paragraphs=[{"text": "H1", "outline": 0, "pstyle": "Heading1"}])

    opt = PipelineOptions(mode="toc", toc=TocConfig(depth=1))
    cancel_flag = False

    def on_event(ev: PipelineEvent):
        nonlocal cancel_flag
        if ev.status == "ok":
            cancel_flag = True

    results = run_pipeline([doc1, doc2], opt, progress_cb=on_event, should_cancel=lambda: cancel_flag)

    assert len(results) == 2
    assert results[0].status == "ok"
    assert results[1].status == "cancelled"


@pytest.mark.skipif(not _HAS_SOFFICE, reason="系统未安装 soffice")
def test_pipeline_run_mode_full(tmp_path: Path):
    """测试 run 模式完整三阶段：DOCX 就绪 -> TOC 注入 -> PDF 导出。"""
    doc_path = tmp_path / "run_doc.docx"
    build_docx(
        doc_path,
        paragraphs=[
            {"text": "Section One", "outline": 0, "pstyle": "Heading1"},
            {"text": "Paragraph A"},
            {"text": "Section Two", "outline": 0, "pstyle": "Heading1", "page_break": True},
            {"text": "Paragraph B"},
        ],
    )

    out_dir = tmp_path / "run_out"
    opt = PipelineOptions(
        mode="run",
        out_dir=out_dir,
        make_docx=True,
        make_pdf=True,
        toc=TocConfig(depth=1),
    )

    events: list[PipelineEvent] = []
    results = run_pipeline([doc_path], opt, progress_cb=events.append)

    # 应包含 CONVERT (准备复制), TOC (注入), PDF (导出) 阶段
    stages = [e.stage for e in results]
    assert Stage.CONVERT in stages
    assert Stage.TOC in stages
    assert Stage.PDF in stages
    assert all(e.status == "ok" for e in results)

    dest_docx = out_dir / "run_doc.docx"
    dest_pdf = out_dir / "run_doc.pdf"
    assert dest_docx.is_file()
    assert dest_pdf.is_file()
    assert dest_pdf.stat().st_size > 0


@pytest.mark.skipif(not _HAS_SOFFICE, reason="系统未安装 soffice")
def test_pipeline_force_and_skip(tmp_path: Path):
    """测试 force=False 时跳过已存在文件，force=True 时重新生成。"""
    doc_path = tmp_path / "force_doc.docx"
    build_docx(doc_path, paragraphs=[{"text": "H1", "outline": 0, "pstyle": "Heading1"}])

    out_dir = tmp_path / "out_force"
    opt = PipelineOptions(
        mode="run",
        out_dir=out_dir,
        make_docx=True,
        make_pdf=True,
        toc=TocConfig(depth=1),
        force=False,
    )

    # 首次执行全部生成
    res1 = run_pipeline([doc_path], opt)
    assert all(e.status == "ok" for e in res1)

    # 再次执行，force=False，目标已存在应当产生 skipped 事件
    res2 = run_pipeline([doc_path], opt)
    skipped_stages = [e.stage for e in res2 if e.status == "skipped"]
    assert Stage.PDF in skipped_stages

    # 开启 force=True，重新执行生成
    opt_force = PipelineOptions(
        mode="run",
        out_dir=out_dir,
        make_docx=True,
        make_pdf=True,
        toc=TocConfig(depth=1),
        force=True,
    )
    res3 = run_pipeline([doc_path], opt_force)
    assert not any(e.status == "skipped" for e in res3)
    assert all(e.status == "ok" for e in res3)
