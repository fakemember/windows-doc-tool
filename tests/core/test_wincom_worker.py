from __future__ import annotations

import sys
from pathlib import Path
from unittest.mock import MagicMock, call, patch

import pytest

import wdt.core.wincom_worker as ww
from wdt.core.converter_worker import FileResult
from wdt.core.wincom_worker import (
    WINCOM_DEFAULT_TIMEOUT_SEC,
    WincomRunner,
    WordSession,
    convert_file_com,
    export_pdf_com,
    is_wincom_supported,
    msoAutomationSecurityForceDisable,
    msoFeatureInstallNone,
    wdAlertsNone,
    wdExportCreateHeadingBookmarks,
    wdExportCreateNoBookmarks,
    wdExportFormatPDF,
    wdExportOptimizeForPrint,
    wdFormatPDF,
    wdFormatXMLDocument,
)


# ---------------------------------------------------------------------------
# 测试 1: 跨平台/无 pywin32 环境下的安全隔离与常量兜底
# ---------------------------------------------------------------------------

def test_constants_fallback():
    """验证模块常量始终具有正确兜底定义。"""
    assert wdFormatXMLDocument == 16
    assert wdFormatPDF == 17
    assert wdAlertsNone == 0
    assert msoAutomationSecurityForceDisable == 3
    assert msoFeatureInstallNone == 0
    assert wdExportFormatPDF == 17
    assert wdExportOptimizeForPrint == 0
    assert wdExportCreateNoBookmarks == 0
    assert wdExportCreateHeadingBookmarks == 1
    assert WINCOM_DEFAULT_TIMEOUT_SEC == 3600


def test_unsupported_platform_or_missing_pywin32():
    """在非 win32 或 pywin32 不存在时，is_wincom_supported 返回 False，调用抛出 NotImplementedError。"""
    with patch.object(ww, "is_wincom_supported", return_value=False):
        assert ww.is_wincom_supported() is False

        with pytest.raises(NotImplementedError, match="Word COM 引擎仅支持 Windows 环境且需安装 pywin32"):
            WordSession()

        with pytest.raises(NotImplementedError, match="Word COM 引擎仅支持 Windows 环境且需安装 pywin32"):
            WincomRunner()

        with pytest.raises(NotImplementedError, match="Word COM 引擎仅支持 Windows 环境且需安装 pywin32"):
            convert_file_com(Path("a.doc"), Path("a.docx"), "docx")

        with pytest.raises(NotImplementedError, match="Word COM 引擎仅支持 Windows 环境且需安装 pywin32"):
            export_pdf_com(Path("a.docx"), Path("a.pdf"))


# ---------------------------------------------------------------------------
# 测试 2: WordSession 装甲参数设置与生命周期
# ---------------------------------------------------------------------------

def test_word_session_armor_settings_and_pid_tracking():
    """验证 WordSession 启动时设置的五大装甲属性以及 PID 的正确捕获。"""
    mock_word = MagicMock()
    mock_pythoncom = MagicMock()
    mock_win32com = MagicMock()
    mock_win32com.client.Dispatch.return_value = mock_word

    with patch.object(ww, "is_wincom_supported", return_value=True), \
         patch.object(ww, "pythoncom", mock_pythoncom), \
         patch.object(ww, "win32com", mock_win32com), \
         patch.object(ww, "_get_winword_pids", side_effect=[{1001}, {1001, 2002}]):

        session = WordSession()
        session.start()

        # 检查 COM 初始化
        mock_pythoncom.CoInitialize.assert_called_once()
        mock_win32com.client.Dispatch.assert_called_once_with("Word.Application")

        # 检查 PID 捕获
        assert session.pid == 2002

        # 检查五大装甲属性
        assert mock_word.Visible is False
        assert mock_word.DisplayAlerts == wdAlertsNone
        assert mock_word.ScreenUpdating is False
        assert mock_word.AutomationSecurity == msoAutomationSecurityForceDisable
        assert mock_word.FeatureInstall == msoFeatureInstallNone

        # 正常关闭
        session.stop()
        mock_word.Quit.assert_called_once_with(SaveChanges=0)
        assert session.word is None
        assert session.pid is None
        mock_pythoncom.CoUninitialize.assert_called_once()


def test_word_session_kill_protection():
    """验证 kill 时调用 taskkill，且严禁触碰 COM 代理，置空并 CoUninitialize。"""
    mock_word = MagicMock()
    mock_pythoncom = MagicMock()

    with patch.object(ww, "is_wincom_supported", return_value=True), \
         patch.object(ww, "pythoncom", mock_pythoncom), \
         patch.object(ww, "_taskkill_pid") as mock_taskkill:

        session = WordSession()
        session.word = mock_word
        session.pid = 9999

        session.kill()

        # 杀进程树
        mock_taskkill.assert_called_once_with(9999)

        # COM 代理置空，未调用任何 mock_word 方法（如 mock_word.Quit）
        assert session.word is None
        assert session.pid is None
        mock_word.Quit.assert_not_called()
        mock_pythoncom.CoUninitialize.assert_called_once()


# ---------------------------------------------------------------------------
# 测试 3: convert_file_com 文件格式映射与正常处理
# ---------------------------------------------------------------------------

def test_convert_file_com_saveas2_format_mapping(tmp_path: Path):
    """测试 convert_file_com 将 docx 映射为 16，将 pdf 映射为 17。"""
    src_doc = tmp_path / "sample.rtf"
    src_doc.write_text("dummy rtf", encoding="ascii")
    dst_docx = tmp_path / "out" / "sample.docx"
    dst_pdf = tmp_path / "out" / "sample.pdf"

    mock_doc = MagicMock()
    mock_word = MagicMock()
    mock_word.Documents.Open.return_value = mock_doc

    # 模拟 SaveAs2 创建临时产物
    def fake_saveas2(FileName, FileFormat):
        Path(FileName).write_bytes(b"converted-content")

    mock_doc.SaveAs2.side_effect = fake_saveas2

    session = MagicMock()
    session.is_alive.return_value = True
    session.word = mock_word

    with patch.object(ww, "is_wincom_supported", return_value=True):
        # 1. 目标 docx -> wdFormatXMLDocument (16)
        res_docx = convert_file_com(src_doc, dst_docx, target="docx", session=session)
        assert res_docx.ok is True
        assert res_docx.dst == dst_docx
        assert dst_docx.is_file()
        assert dst_docx.read_bytes() == b"converted-content"
        # 校验 SaveAs2 传入 FileFormat=16
        save_call_docx = mock_doc.SaveAs2.call_args
        assert save_call_docx.kwargs["FileFormat"] == wdFormatXMLDocument
        mock_doc.Close.assert_called_with(SaveChanges=0)

        # 2. 目标 pdf -> wdFormatPDF (17)
        res_pdf = convert_file_com(src_doc, dst_pdf, target="pdf", session=session)
        assert res_pdf.ok is True
        assert res_pdf.dst == dst_pdf
        assert dst_pdf.is_file()
        save_call_pdf = mock_doc.SaveAs2.call_args
        assert save_call_pdf.kwargs["FileFormat"] == wdFormatPDF


def test_convert_file_com_invalid_target(tmp_path: Path):
    """不支持的目标格式抛出 ValueError。"""
    src = tmp_path / "sample.rtf"
    src.write_text("rtf", encoding="ascii")
    with patch.object(ww, "is_wincom_supported", return_value=True):
        with pytest.raises(ValueError, match="Unsupported target"):
            convert_file_com(src, tmp_path / "sample.txt", target="txt")  # type: ignore


# ---------------------------------------------------------------------------
# 测试 4: export_pdf_com 标题书签与参数校验
# ---------------------------------------------------------------------------

def test_export_pdf_com_parameters(tmp_path: Path):
    """测试 export_pdf_com 使用 ExportAsFixedFormat 导出并传递标题书签参数。"""
    src_docx = tmp_path / "test.docx"
    src_docx.write_bytes(b"docx content")
    dst_pdf = tmp_path / "test.pdf"

    mock_doc = MagicMock()
    mock_word = MagicMock()
    mock_word.Documents.Open.return_value = mock_doc

    def fake_export(OutputFileName, **kwargs):
        Path(OutputFileName).write_bytes(b"%PDF-fake")

    mock_doc.ExportAsFixedFormat.side_effect = fake_export

    session = MagicMock()
    session.is_alive.return_value = True
    session.word = mock_word

    with patch.object(ww, "is_wincom_supported", return_value=True):
        # 默认 create_bookmarks=True -> wdExportCreateHeadingBookmarks (1)
        res = export_pdf_com(src_docx, dst_pdf, create_bookmarks=True, session=session)
        assert res.ok is True
        assert res.dst == dst_pdf
        assert dst_pdf.is_file()

        export_call = mock_doc.ExportAsFixedFormat.call_args
        assert export_call.kwargs["ExportFormat"] == wdExportFormatPDF
        assert export_call.kwargs["CreateBookmarks"] == wdExportCreateHeadingBookmarks
        assert export_call.kwargs["OptimizeFor"] == wdExportOptimizeForPrint

        # create_bookmarks=False -> wdExportCreateNoBookmarks (0)
        res2 = export_pdf_com(src_docx, dst_pdf, create_bookmarks=False, session=session)
        assert res2.ok is True
        export_call2 = mock_doc.ExportAsFixedFormat.call_args
        assert export_call2.kwargs["CreateBookmarks"] == wdExportCreateNoBookmarks


# ---------------------------------------------------------------------------
# 测试 5: 超时看门狗与异常查杀保护
# ---------------------------------------------------------------------------

def test_convert_file_com_timeout_triggers_kill(tmp_path: Path):
    """验证当发生超时时，触发看门狗杀进程，返回错误且不抛出未捕获异常。"""
    src = tmp_path / "hang.rtf"
    src.write_text("hang", encoding="ascii")
    dst = tmp_path / "hang.pdf"

    session = WordSession.__new__(WordSession)
    session.word = MagicMock()
    session.pid = 8888
    session._killed = False

    def hang_open(*args, **kwargs):
        import time
        time.sleep(0.5)
        raise RuntimeError("RPC disconnect simulation")

    session.word.Documents.Open.side_effect = hang_open

    with patch.object(ww, "is_wincom_supported", return_value=True), \
         patch.object(ww, "_taskkill_pid") as mock_kill, \
         patch.object(ww, "_pid_alive", return_value=False), \
         patch.object(ww, "_spawn_watchdog") as mock_wd, \
         patch.object(ww, "pythoncom", MagicMock()):

        mock_wd.return_value.poll.return_value = 0  # 看门狗已触发查杀
        # 设置超时极短 (0.1秒)
        res = convert_file_com(src, dst, target="pdf", timeout_sec=0.1, session=session)

        assert res.ok is False
        assert "超时" in res.error
        assert mock_kill.called
        assert session.word is None


def test_convert_file_com_timeout_unkillable_word_advises_reboot(tmp_path: Path):
    """看门狗触发但 Word 进程无法终止时，错误信息应提示重启系统。"""
    src = tmp_path / "hang.rtf"
    src.write_text("hang", encoding="ascii")
    dst = tmp_path / "hang.pdf"

    session = WordSession.__new__(WordSession)
    session.word = MagicMock()
    session.pid = 8888
    session._killed = False

    def hang_open(*args, **kwargs):
        import time
        time.sleep(0.5)
        raise RuntimeError("RPC disconnect simulation")

    session.word.Documents.Open.side_effect = hang_open

    with patch.object(ww, "is_wincom_supported", return_value=True), \
         patch.object(ww, "_taskkill_pid"), \
         patch.object(ww, "_pid_alive", return_value=True), \
         patch.object(ww, "_spawn_watchdog") as mock_wd, \
         patch("time.sleep"), \
         patch.object(ww, "pythoncom", MagicMock()):

        mock_wd.return_value.poll.return_value = 0  # 看门狗已触发查杀
        res = convert_file_com(src, dst, target="pdf", timeout_sec=0.1, session=session)

        assert res.ok is False
        assert "超时" in res.error
        assert "重启" in res.error


# ---------------------------------------------------------------------------
# 测试 6: WincomRunner 串行调度、单文件失败隔离与实例重拉
# ---------------------------------------------------------------------------

def test_wincom_runner_failure_isolation_and_restart(tmp_path: Path):
    """测试 WincomRunner 单文件崩溃时不中断批处理，自动拉起新 Word 实例并完成全部文件。"""
    f1 = tmp_path / "doc1.rtf"
    f2 = tmp_path / "doc2.rtf"  # 这个文件将触发崩溃/查杀
    f3 = tmp_path / "doc3.rtf"
    for f in (f1, f2, f3):
        f.write_text(f"content of {f.name}", encoding="ascii")

    pairs = [
        (f1, tmp_path / "out" / "doc1.docx"),
        (f2, tmp_path / "out" / "doc2.docx"),
        (f3, tmp_path / "out" / "doc3.docx"),
    ]

    session_start_count = 0

    def fake_convert(src, dst, target, timeout_sec, session):
        nonlocal session_start_count
        if src.name == "doc2.rtf":
            # 模拟崩溃并杀掉会话
            session.kill()
            return FileResult(src=src, dst=None, ok=False, elapsed=0.1, error="RPC crashed")
        else:
            dst.parent.mkdir(parents=True, exist_ok=True)
            dst.write_bytes(b"dummy docx")
            return FileResult(src=src, dst=dst, ok=True, elapsed=0.1, error="")

    with patch.object(ww, "is_wincom_supported", return_value=True), \
         patch.object(ww, "convert_file_com", side_effect=fake_convert):

        runner = WincomRunner()
        assert runner.jobs == 1

        with patch.object(WordSession, "start", autospec=True) as mock_start, \
             patch.object(WordSession, "stop", autospec=True) as mock_stop:

            def set_word(self):
                self.word = MagicMock()
                self._killed = False

            mock_start.side_effect = set_word

            cb_calls = []
            results = runner.convert(
                pairs,
                target="docx",
                progress_cb=lambda done, total, res: cb_calls.append((done, total, res)),
            )

        assert len(results) == 3
        # 结果对齐 pairs 原始顺序
        assert results[0].src == f1 and results[0].ok is True
        assert results[1].src == f2 and results[1].ok is False
        assert results[2].src == f3 and results[2].ok is True

        # 回调进度正确推进
        assert len(cb_calls) == 3
        assert [c[0] for c in cb_calls] == [1, 2, 3]

        # 会话应在处理 f1 (1次) 和 f2崩溃后重拉 f3 (第2次)，共 start 2 次
        assert mock_start.call_count == 2
        mock_stop.assert_called_once()


# ---------------------------------------------------------------------------
# 测试 7: WincomRunner 协作式取消 (should_cancel)
# ---------------------------------------------------------------------------

def test_wincom_runner_cancel_before_start(tmp_path: Path):
    """在任务开始前立即取消。"""
    pairs = [(tmp_path / "a.rtf", tmp_path / "a.docx")]
    with patch.object(ww, "is_wincom_supported", return_value=True):
        runner = WincomRunner()
        results = runner.convert(pairs, target="docx", should_cancel=lambda: True)
        assert len(results) == 1
        assert results[0].ok is False
        assert results[0].error == "cancelled"


def test_wincom_runner_cancel_during_execution(tmp_path: Path):
    """在任务执行中间取消，首个成功，后续标记为 cancelled。"""
    f1 = tmp_path / "doc1.rtf"
    f2 = tmp_path / "doc2.rtf"
    pairs = [(f1, tmp_path / "doc1.docx"), (f2, tmp_path / "doc2.docx")]

    cancel_flag = False

    def fake_convert(src, dst, target, timeout_sec, session):
        nonlocal cancel_flag
        cancel_flag = True
        return FileResult(src=src, dst=dst, ok=True, elapsed=0.1, error="")

    with patch.object(ww, "is_wincom_supported", return_value=True), \
         patch.object(WordSession, "start"), \
         patch.object(WordSession, "stop"), \
         patch.object(ww, "convert_file_com", side_effect=fake_convert):

        runner = WincomRunner()
        results = runner.convert(
            pairs,
            target="docx",
            should_cancel=lambda: cancel_flag,
        )

        assert len(results) == 2
        assert results[0].ok is True
        assert results[1].ok is False
        assert results[1].error == "cancelled"
