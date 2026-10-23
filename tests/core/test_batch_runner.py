import shutil
from pathlib import Path
from unittest.mock import patch

import pytest
from wdt.core.batch_runner import BatchRunner, _default_jobs
from wdt.core.converter_worker import FileResult
from wdt.core.engine_finder import find_engine

_HAS_SOFFICE = shutil.which("soffice") is not None


def test_default_jobs_clamping():
    with patch("os.cpu_count", return_value=16), \
         patch("wdt.core.batch_runner._get_physical_memory_gb", return_value=32.0):
        # min(16-1, 8, 32//1.5=21) -> 8
        assert _default_jobs() == 8

    with patch("os.cpu_count", return_value=4), \
         patch("wdt.core.batch_runner._get_physical_memory_gb", return_value=4.0):
        # min(4-1=3, 8, 4//1.5=2) -> 2
        assert _default_jobs() == 2

    with patch("os.cpu_count", return_value=1), \
         patch("wdt.core.batch_runner._get_physical_memory_gb", return_value=1.0):
        # 1核 1GB -> max(1, ...)
        assert _default_jobs() >= 1


@pytest.mark.skipif(not _HAS_SOFFICE, reason="系统未安装 soffice")
def test_batch_runner_docx_and_pdf_with_failure_isolation(tmp_path: Path):
    engine = find_engine()
    runner = BatchRunner(engine=engine, jobs=2)

    # 构造 4 个文件：
    # 1. 正常 rtf (doc1.rtf)
    # 2. 伪 .doc 后缀的真 rtf (doc2.doc)
    # 3. 正常 rtf (doc3.rtf)
    # 4. 损坏的垃圾文件 (corrupt.doc)，无法被解析
    doc1 = tmp_path / "doc1.rtf"
    doc2 = tmp_path / "doc2.doc"
    doc3 = tmp_path / "doc3.rtf"
    doc_bad = tmp_path / "corrupt.doc"

    doc1.write_bytes(b"{\\rtf1\\ansi\\deff0 {\\fonttbl{\\f0 Arial;}}\\f0\\fs24 Doc 1 Content\\par}")
    doc2.write_bytes(b"{\\rtf1\\ansi\\deff0 {\\fonttbl{\\f0 Arial;}}\\f0\\fs24 Pseudo Doc Content\\par}")
    doc3.write_bytes(b"{\\rtf1\\ansi\\deff0 {\\fonttbl{\\f0 Arial;}}\\f0\\fs24 Doc 3 Content\\par}")
    doc_bad.write_bytes(b"PK\x03\x04corrupted_non_zip_data_that_fails_load")

    out_dir = tmp_path / "output_docx"
    pairs = [
        (doc1, out_dir / "doc1.docx"),
        (doc2, out_dir / "doc2.docx"),
        (doc_bad, out_dir / "corrupt.docx"),
        (doc3, out_dir / "doc3.docx"),
    ]

    cb_calls: list[tuple[int, int, FileResult]] = []

    def progress_callback(done: int, total: int, res: FileResult):
        cb_calls.append((done, total, res))

    # 执行 docx 批量转换
    results = runner.convert(pairs, target="docx", progress_cb=progress_callback)

    # 1. 结果顺序必须与输入 pairs 保持一致
    assert [r.src for r in results] == [doc1, doc2, doc_bad, doc3]

    # 2. 回调次数必须严格等于总文件数 4
    assert len(cb_calls) == 4
    assert [c[0] for c in cb_calls] == [1, 2, 3, 4]
    assert all(c[1] == 4 for c in cb_calls)

    # 3. 坏文件被隔离重试后失败，但不拖垮正常文件
    assert results[0].ok is True
    assert results[0].dst == out_dir / "doc1.docx"
    assert (out_dir / "doc1.docx").is_file()
    assert (out_dir / "doc1.docx").stat().st_size > 0

    assert results[1].ok is True
    assert results[1].dst == out_dir / "doc2.docx"
    assert (out_dir / "doc2.docx").is_file()
    assert (out_dir / "doc2.docx").stat().st_size > 0

    assert results[2].ok is False
    assert results[2].dst is None
    assert results[2].error != ""

    assert results[3].ok is True
    assert results[3].dst == out_dir / "doc3.docx"
    assert (out_dir / "doc3.docx").is_file()
    assert (out_dir / "doc3.docx").stat().st_size > 0


@pytest.mark.skipif(not _HAS_SOFFICE, reason="系统未安装 soffice")
def test_batch_runner_pdf_gray(tmp_path: Path):
    engine = find_engine()
    runner = BatchRunner(engine=engine, jobs=2)

    doc1 = tmp_path / "g1.rtf"
    doc2 = tmp_path / "g2.doc"  # 伪装为 .doc 的真 rtf

    doc1.write_bytes(b"{\\rtf1\\ansi Gray 1}")
    doc2.write_bytes(b"{\\rtf1\\ansi Gray 2}")

    out_dir = tmp_path / "output_pdf"
    pairs = [
        (doc1, out_dir / "g1.pdf"),
        (doc2, out_dir / "g2.pdf"),
    ]

    cb_results = []
    results = runner.convert(
        pairs,
        target="pdf",
        gray=True,
        progress_cb=lambda d, t, r: cb_results.append(r),
    )

    assert len(results) == 2
    assert len(cb_results) == 2
    for r in results:
        assert r.ok, f"Failed: {r.error}"
        assert r.dst is not None and r.dst.is_file()
        assert r.dst.stat().st_size > 0


def test_batch_runner_empty():
    engine = find_engine() if _HAS_SOFFICE else None
    if engine:
        runner = BatchRunner(engine=engine)
        assert runner.convert([], target="docx") == []


def test_batch_runner_cancel_immediate(tmp_path: Path):
    """测试任务开始前直接触发取消。"""
    engine = find_engine() if _HAS_SOFFICE else None
    if not engine:
        from wdt.core.engine_finder import Engine
        engine = Engine(soffice=Path("/dummy/soffice"), source="mock", version=(7, 5))

    runner = BatchRunner(engine=engine, jobs=2)
    p1 = tmp_path / "a.rtf"
    p2 = tmp_path / "b.rtf"
    p1.write_bytes(b"dummy")
    p2.write_bytes(b"dummy")
    pairs = [(p1, tmp_path / "a.docx"), (p2, tmp_path / "b.docx")]

    results = runner.convert(pairs, target="docx", should_cancel=lambda: True)
    assert len(results) == 2
    assert all(not r.ok and r.error == "cancelled" for r in results)


def test_batch_runner_cancel_during_execution(tmp_path: Path):
    """测试在执行过程中完成首个 chunk 后触发取消。"""
    from concurrent.futures import Future
    from unittest.mock import MagicMock
    from wdt.core.engine_finder import Engine

    engine = Engine(soffice=Path("/dummy/soffice"), source="mock", version=(7, 5))
    runner = BatchRunner(engine=engine, jobs=2)

    p1 = tmp_path / "doc1.rtf"
    p2 = tmp_path / "doc2.rtf"
    pairs = [(p1, tmp_path / "doc1.docx"), (p2, tmp_path / "doc2.docx")]

    cancel_flag = False

    def should_cancel_check():
        return cancel_flag

    f1 = Future()
    f1.set_result([FileResult(src=p1, dst=pairs[0][1], ok=True, elapsed=0.01, error="")])
    f2 = Future()

    mock_executor = MagicMock()
    mock_executor.submit.side_effect = [f1, f2]

    def on_f1_done(done, total, res):
        nonlocal cancel_flag
        cancel_flag = True

    with patch("wdt.core.batch_runner.ProcessPoolExecutor", return_value=mock_executor), \
         patch("wdt.core.batch_runner.as_completed", return_value=[f1]):
        results = runner.convert(
            pairs,
            target="docx",
            progress_cb=on_f1_done,
            should_cancel=should_cancel_check,
        )

    assert mock_executor.shutdown.call_count >= 1
    mock_executor.shutdown.assert_any_call(wait=False, cancel_futures=True)
    assert len(results) == 2
    assert results[0].ok is True
    assert results[1].ok is False
    assert results[1].error == "cancelled"


