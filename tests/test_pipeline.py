"""
流水线 (Pipeline) 模块单元与集成测试用例。
"""

from __future__ import annotations

import zipfile
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from tests.fixtures_factory import build_docx
from wdt.core.converter_worker import FileResult
from wdt.core.engine_finder import EngineNotFoundError
from wdt.pipeline import (
    PipelineEvent,
    PipelineOptions,
    Stage,
    run_pipeline,
)
from wdt.toc.docx_stream_patcher import TocConfig
from wdt.toc.page_calculator import PageMode


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

    # 3. convert 模式未选择任何输出格式
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

    def cb(ev: PipelineEvent):
        events.append(ev)

    res = run_pipeline([doc_path], opt, progress_cb=cb)
    assert len(res) == 1
    assert res[0].status == "ok"
    assert res[0].stage == Stage.TOC
    assert len(events) == 1


def test_pipeline_toc_cancellation(tmp_path: Path):
    """测试协作式取消机制在 toc 模式下正常工作。"""
    doc_path = tmp_path / "cancel.docx"
    build_docx(doc_path, paragraphs=[{"text": "Title", "outline": 0, "pstyle": "Heading1"}])

    opt = PipelineOptions(mode="toc", toc=TocConfig(depth=1))
    events = run_pipeline([doc_path], opt, should_cancel=lambda: True)

    assert len(events) == 1
    assert events[0].status == "cancelled"


def _make_fake_runner(fail: bool = False) -> MagicMock:
    runner = MagicMock()

    def fake_convert(pairs, target, progress_cb=None, should_cancel=None, **kwargs):
        results = []
        for src, dst in pairs:
            if should_cancel and should_cancel():
                r = FileResult(src=src, dst=None, ok=False, elapsed=0.0, error="cancelled")
            elif fail:
                r = FileResult(src=src, dst=None, ok=False, elapsed=0.1, error="mock error")
            else:
                dst.parent.mkdir(parents=True, exist_ok=True)
                dst.write_bytes(b"converted")
                r = FileResult(src=src, dst=dst, ok=True, elapsed=0.1, error="")
            results.append(r)
            if progress_cb:
                progress_cb(len(results), len(pairs), r)
        return results

    runner.convert.side_effect = fake_convert
    return runner


def test_pipeline_run_mode_full(tmp_path: Path, monkeypatch):
    """测试 run 模式完整流程：RTF -> DOCX -> TOC -> PDF。"""
    src_rtf = tmp_path / "doc.rtf"
    src_rtf.write_text(r"{\rtf1\ansi Hello World}")

    out_dir = tmp_path / "out"

    runner = _make_fake_runner()
    monkeypatch.setattr("wdt.pipeline.WincomRunner", lambda: runner)
    monkeypatch.setattr("wdt.pipeline.is_wincom_supported", lambda: True)

    fake_report = MagicMock()
    fake_report.entries = 1
    fake_report.elapsed = 0.05
    fake_report.pages = [1]
    fake_report.sentinel = ""
    monkeypatch.setattr("wdt.pipeline.inject_toc", lambda *a, **k: fake_report)

    opt = PipelineOptions(
        mode="run",
        out_dir=out_dir,
        make_docx=True,
        make_pdf=True,
        toc=TocConfig(depth=1),
    )

    events: list[PipelineEvent] = []
    run_pipeline([src_rtf], opt, progress_cb=events.append)

    stages = [e.stage for e in events]
    assert Stage.CONVERT in stages
    assert Stage.TOC in stages
    assert Stage.PDF in stages
    terminal = [e for e in events if e.status != "running"]
    assert all(e.status == "ok" for e in terminal)
    # running 中间事件必须先于同阶段 ok 事件出现
    seq = [(e.stage, e.status) for e in events]
    for stage in (Stage.CONVERT, Stage.PDF):
        first_run = next(i for i, s in enumerate(seq) if s == (stage, "running"))
        first_ok = next(i for i, s in enumerate(seq) if s == (stage, "ok"))
        assert first_run < first_ok


def test_pipeline_force_and_skip(tmp_path: Path, monkeypatch):
    """测试已存在目标文件时的 skip 逻辑与 force 覆盖逻辑。"""
    src_rtf = tmp_path / "report.rtf"
    src_rtf.write_text(r"{\rtf1}")
    out_dir = tmp_path / "out"
    out_dir.mkdir()
    existing_docx = out_dir / "report.docx"
    existing_docx.write_bytes(b"existing")

    runner = _make_fake_runner()
    monkeypatch.setattr("wdt.pipeline.WincomRunner", lambda: runner)
    monkeypatch.setattr("wdt.pipeline.is_wincom_supported", lambda: True)

    opt_skip = PipelineOptions(mode="convert", out_dir=out_dir, make_docx=True, force=False)
    events_skip = run_pipeline([src_rtf], opt_skip)
    assert events_skip[0].status == "skipped"

    opt_force = PipelineOptions(mode="convert", out_dir=out_dir, make_docx=True, force=True)
    events_force = run_pipeline([src_rtf], opt_force)
    assert events_force[0].status == "ok"


def test_pipeline_unsupported_environment_raises(tmp_path: Path, monkeypatch):
    """当环境不支持 Word COM 时，流水线抛出异常。"""
    doc = tmp_path / "a.doc"
    doc.write_bytes(b"x")
    monkeypatch.setattr("wdt.pipeline.is_wincom_supported", lambda: False)

    opt = PipelineOptions(mode="convert", make_docx=True)
    with pytest.raises(EngineNotFoundError, match="Word COM"):
        run_pipeline([doc], opt)
