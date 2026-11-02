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


# ---------------------------------------------------------------------------
# 双引擎路由 (auto 仲裁 / 规则 C / 显式引擎 / wincom 旁路 measured 2-Pass)
# ---------------------------------------------------------------------------

_FAKE_LO_ENGINE = Engine(soffice=Path("/usr/bin/soffice"), source="path", version=(26, 0))


def _make_fake_runner() -> MagicMock:
    """构造行为与真实 Runner 一致的假调度器：写占位产物并按序上报进度。"""
    runner = MagicMock()

    def _convert(pairs, target, gray=False, progress_cb=None, should_cancel=None):
        results = []
        for i, (src, dst) in enumerate(pairs, 1):
            dst.parent.mkdir(parents=True, exist_ok=True)
            dst.write_bytes(b"wdt-fake-output")
            res = FileResult(src=src, dst=dst, ok=True, elapsed=0.1, error="")
            results.append(res)
            if progress_cb:
                progress_cb(i, len(pairs), res)
        return results

    runner.convert.side_effect = _convert
    return runner


def _patch_runners(monkeypatch) -> tuple[MagicMock, MagicMock]:
    """替换 pipeline 内的双引擎 Runner 与支持探测，返回 (lo_runner, wc_runner)。"""
    lo_runner = _make_fake_runner()
    wc_runner = _make_fake_runner()
    monkeypatch.setattr(
        "wdt.pipeline.BatchRunner", lambda engine=None, jobs=None: lo_runner
    )
    monkeypatch.setattr("wdt.pipeline.WincomRunner", lambda: wc_runner)
    monkeypatch.setattr("wdt.pipeline.is_wincom_supported", lambda: True)
    return lo_runner, wc_runner


def test_pipeline_auto_rule_c_upgrades_batch(tmp_path: Path, monkeypatch):
    """规则 C：任务含 PDF 导出时，批内任一文件仲裁为 wincom 则整批统一升级。"""
    small = tmp_path / "table.doc"
    small.write_bytes(b"a" * 10)
    big = tmp_path / "listing.rtf"
    big.write_bytes(b"b" * 2000)

    lo_runner, wc_runner = _patch_runners(monkeypatch)
    monkeypatch.setattr(
        "wdt.pipeline.arbitrate_engine",
        lambda **kw: "wincom" if kw["file_bytes"] > 1000 else "libreoffice",
    )

    opt = PipelineOptions(
        mode="convert", make_docx=True, make_pdf=True, engine_mode="auto"
    )
    events = run_pipeline([small, big], opt, engine=_FAKE_LO_ENGINE)

    assert wc_runner.convert.call_count == 2
    for call in wc_runner.convert.call_args_list:
        assert len(call.args[0]) == 2
    assert lo_runner.convert.call_count == 0
    assert any("[Word COM]" in e.message for e in events)
    assert all(e.status == "ok" for e in events)


def test_pipeline_auto_split_groups_and_skip_wincom_normalize(
    tmp_path: Path, monkeypatch
):
    """auto 仅转 DOCX：按文件分组路由，wincom 产物跳过 LibreOffice 专用规范化。"""
    small = tmp_path / "table.doc"
    small.write_bytes(b"a" * 10)
    big = tmp_path / "listing.rtf"
    big.write_bytes(b"b" * 2000)

    lo_runner, wc_runner = _patch_runners(monkeypatch)
    monkeypatch.setattr(
        "wdt.pipeline.arbitrate_engine",
        lambda **kw: "wincom" if kw["file_bytes"] > 1000 else "libreoffice",
    )

    opt = PipelineOptions(
        mode="convert", make_docx=True, make_pdf=False, engine_mode="auto"
    )
    with patch("wdt.pipeline._safe_normalize_converted") as mock_norm:
        events = run_pipeline([small, big], opt, engine=_FAKE_LO_ENGINE)

    lo_pairs = lo_runner.convert.call_args_list[0].args[0]
    wc_pairs = wc_runner.convert.call_args_list[0].args[0]
    assert [p[0] for p in lo_pairs] == [small]
    assert [p[0] for p in wc_pairs] == [big]
    assert mock_norm.call_count == 1
    assert any("[LibreOffice]" in e.message for e in events)
    assert any("[Word COM]" in e.message for e in events)
    assert all(e.status == "ok" for e in events)


def test_pipeline_explicit_wincom_unsupported_raises(tmp_path: Path, monkeypatch):
    """显式指定 wincom 但环境不支持时必须直接报错，不静默回退。"""
    doc = tmp_path / "a.doc"
    doc.write_bytes(b"x")
    monkeypatch.setattr("wdt.pipeline.is_wincom_supported", lambda: False)

    opt = PipelineOptions(mode="convert", make_docx=True, engine_mode="wincom")
    with pytest.raises(ValueError, match="wincom"):
        run_pipeline([doc], opt)


def test_pipeline_wincom_bypasses_2pass_measured(tmp_path: Path, monkeypatch):
    """测试 wincom 引擎自动 bypass measured 临时 PDF 测算，避免超大文件 2-Pass 冗余耗时。"""
    doc = tmp_path / "fast_m.docx"
    build_docx(doc, paragraphs=[{"text": "H1", "outline": 0, "pstyle": "Heading1"}])

    wc_runner = _make_fake_runner()
    monkeypatch.setattr("wdt.pipeline.WincomRunner", lambda: wc_runner)
    monkeypatch.setattr("wdt.pipeline.is_wincom_supported", lambda: True)

    with patch("wdt.pipeline._perform_page_measurement") as mock_measure:
        opt = PipelineOptions(
            mode="toc",
            toc=TocConfig(depth=1, page_mode=PageMode.STATIC_MEASURED),
            engine_mode="wincom",
        )
        events = run_pipeline([doc], opt)

        # 确保没有调用 2-Pass 临时 PDF 测算函数
        mock_measure.assert_not_called()
        assert events[0].stage == Stage.TOC
        assert events[0].status == "ok"
        assert "单遍直出" in events[0].message

        # 验证段落中没有注入 WDTTAG 隐藏锚点（书签保持纯净）
        import zipfile
        with zipfile.ZipFile(doc, "r") as zf:
            xml_content = zf.read("word/document.xml").decode("utf-8")
        assert "WDTTAG" not in xml_content
