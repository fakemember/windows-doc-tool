from __future__ import annotations

import csv
import io
import os
import shutil
import subprocess
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Literal, Sequence

from wdt.core.converter_worker import FileResult

# ---------------------------------------------------------------------------
# pywin32 软隔离与跨平台常量兜底
# ---------------------------------------------------------------------------

try:
    import pythoncom  # type: ignore
    import win32com.client  # type: ignore

    _PYWIN32_AVAILABLE = True
except ImportError:
    pythoncom = None  # type: ignore
    win32com = None  # type: ignore
    _PYWIN32_AVAILABLE = False


def is_wincom_supported() -> bool:
    """判定当前环境是否支持 Word COM 自动化。

    需同时满足：Windows 平台 (win32) 且成功导入 pywin32 模块。
    """
    return sys.platform == "win32" and _PYWIN32_AVAILABLE


# Word / Office 常量兜底（无论是否安装 pywin32 均可用，保障测试与跨平台可导入）
# WdSaveFormat
wdFormatDocumentDefault = 16  # .docx
wdFormatXMLDocument = 16  # .docx
wdFormatPDF = 17  # .pdf

# WdAlertLevel
wdAlertsNone = 0

# MsoAutomationSecurity
msoAutomationSecurityForceDisable = 3

# MsoFeatureInstall
msoFeatureInstallNone = 0

# WdExportFormat
wdExportFormatPDF = 17

# WdExportOptimizeFor
wdExportOptimizeForPrint = 0

# WdExportCreateBookmarks
wdExportCreateNoBookmarks = 0
wdExportCreateHeadingBookmarks = 1

# 默认超时：60 分钟 (3600 秒)。
# 基准 (空闲机器) 导出 3,815 页约 20 分钟；CPU 高负载竞争时 Word 单线程导出
# 会被饿到数倍时长，60 分钟为 contention 场景保留充足余量，看门狗仍兜底真挂死。
WINCOM_DEFAULT_TIMEOUT_SEC = 3600


# ---------------------------------------------------------------------------
# 进程管理辅助函数
# ---------------------------------------------------------------------------

def _get_winword_pids() -> set[int]:
    """获取当前所有 WINWORD.EXE 进程 PID 集合。"""
    if sys.platform != "win32":
        return set()
    try:
        res = subprocess.run(
            ["tasklist.exe", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            check=False,
        )
        pids: set[int] = set()
        for row in csv.reader(io.StringIO(res.stdout.strip())):
            if len(row) >= 2 and "winword" in row[0].lower():
                try:
                    pids.add(int(row[1]))
                except ValueError:
                    pass
        return pids
    except Exception:
        return set()


def _taskkill_pid(pid: int) -> None:
    """强制杀除指定 PID 及其整个进程树。"""
    try:
        subprocess.run(
            ["taskkill.exe", "/PID", str(pid), "/T", "/F"],
            capture_output=True,
            check=False,
        )
    except Exception:
        pass


def _pid_alive(pid: int) -> bool:
    """检测指定 PID 的 Windows 进程是否仍存活。"""
    if sys.platform != "win32":
        return False
    try:
        res = subprocess.run(
            ["tasklist.exe", "/FI", f"PID eq {pid}", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            check=False,
        )
        return str(pid) in res.stdout
    except Exception:
        return False


def _spawn_watchdog(pid: int, timeout_sec: float) -> subprocess.Popen:
    """拉起独立看门狗子进程：超时后对 Word 进程树执行 taskkill 并验证终止。

    使用独立进程而非 threading.Timer：主机高负载或主线程阻塞在 COM 调用中时，
    Python 线程可能被饿死而错过超时窗口；独立进程不受本进程调度影响。
    """
    script = (
        "import subprocess,sys,time\n"
        "pid=int(sys.argv[1]); timeout=float(sys.argv[2])\n"
        "time.sleep(max(timeout, 1))\n"
        "for _ in range(3):\n"
        "    subprocess.run(['taskkill.exe','/PID',str(pid),'/T','/F'],capture_output=True,check=False)\n"
        "    time.sleep(15)\n"
        "    out=subprocess.run(['tasklist.exe','/FI',f'PID eq {pid}','/FO','CSV','/NH'],capture_output=True,text=True,check=False).stdout\n"
        "    if str(pid) not in out:\n"
        "        break\n"
    )
    return subprocess.Popen(
        [sys.executable, "-c", script, str(pid), str(timeout_sec)],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def _wait_output(path: Path, timeout_sec: int = 120) -> bool:
    """轮询等待 Word 通过打印管线异步刷新的输出文件落盘。

    ExportAsFixedFormat/SaveAs2 在打印后台处理程序高负载或刚重启后
    可能先行返回、由 spooler 异步完成写盘；此处给予宽限窗口。
    """
    deadline = time.time() + timeout_sec
    while time.time() < deadline:
        try:
            if path.is_file() and path.stat().st_size > 0:
                return True
        except OSError:
            pass
        time.sleep(2)
    return False


# ---------------------------------------------------------------------------
# Word COM 装甲会话上下文
# ---------------------------------------------------------------------------

class WordSession:
    """单个受控 Microsoft Word COM 会话装甲。

    特性：
    - DisplayAlerts = 0 (wdAlertsNone)：静默一切弹窗
    - ScreenUpdating = False：禁用屏幕重绘，提升性能
    - AutomationSecurity = 3 (msoAutomationSecurityForceDisable)：禁用一切宏
    - FeatureInstall = 0 (msoFeatureInstallNone)：静默缺失功能安装请求
    - 记录 WINWORD.EXE PID；
    - 超时或不可恢复挂起时调用 taskkill /PID <pid> /T /F 杀进程树；
    - 查杀后立即置空 COM 代理并调用 pythoncom.CoUninitialize()；
    - 正常退出时安全 Close/Quit。
    """

    def __init__(self) -> None:
        if not is_wincom_supported():
            raise NotImplementedError("Word COM 引擎仅支持 Windows 环境且需安装 pywin32")
        self.word = None
        self.pid: int | None = None
        self._killed: bool = False
        # taskkill 后仍存活的 PID（Word 卡在打印驱动调用时无法立即终止）；
        # 非 None 时提示用户重启系统，这是唯一可靠的外部处置。
        self.kill_failed: int | None = None

    def start(self) -> None:
        """初始化 COM 并启动受装甲保护的 Word.Application。"""
        if self.word is not None:
            return

        if pythoncom is not None:
            try:
                pythoncom.CoInitialize()
            except Exception:
                pass

        pids_before = _get_winword_pids()
        try:
            self.word = win32com.client.Dispatch("Word.Application")
            pids_after = _get_winword_pids()
            new_pids = pids_after - pids_before
            if new_pids:
                self.pid = next(iter(new_pids))

            # 装甲设置
            try:
                self.word.Visible = False
            except Exception:
                pass
            try:
                self.word.DisplayAlerts = wdAlertsNone
            except Exception:
                pass
            try:
                self.word.ScreenUpdating = False
            except Exception:
                pass
            try:
                self.word.AutomationSecurity = msoAutomationSecurityForceDisable
            except Exception:
                pass
            try:
                self.word.FeatureInstall = msoFeatureInstallNone
            except Exception:
                pass
            self._killed = False
        except Exception:
            self.kill()
            raise

    def kill(self) -> None:
        """强制终止 Word 进程树并重置 COM 代理，防止悬空指针 RPC_E_DISCONNECTED。"""
        self._killed = True
        pid = self.pid
        self.pid = None
        if pid is not None:
            _taskkill_pid(pid)
            # 验证进程确实已终止：Word 卡在打印驱动调用中时 taskkill 也无法立即生效
            for _ in range(3):
                if not _pid_alive(pid):
                    break
                time.sleep(10)
            self.kill_failed = pid if _pid_alive(pid) else None

        # 杀掉后严禁再触碰任何已断开的 COM 代理
        self.word = None
        if pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def stop(self) -> None:
        """正常关闭 Word 实例。"""
        if self._killed or self.word is None:
            self.word = None
            self.pid = None
            return

        try:
            self.word.Quit(SaveChanges=0)
        except Exception:
            # 若正常 Quit 失败，保底 taskkill
            if self.pid is not None:
                _taskkill_pid(self.pid)
        finally:
            self.word = None
            self.pid = None
            if pythoncom is not None:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    def is_alive(self) -> bool:
        return not self._killed and self.word is not None


# ---------------------------------------------------------------------------
# 单文件转换与 PDF 导出函数
# ---------------------------------------------------------------------------

def convert_file_com(
    src: Path,
    dst: Path,
    target: Literal["docx", "pdf"],
    timeout_sec: int = WINCOM_DEFAULT_TIMEOUT_SEC,
    session: WordSession | None = None,
) -> FileResult:
    """使用 Word COM 转换单文件 (.doc/.rtf/.docx -> .docx 或 .pdf)。

    - SaveAs2 FileFormat 按目标扩展名/target 映射：
      - docx: wdFormatXMLDocument (16)
      - pdf: wdFormatPDF (17)
    - 超时守护（默认 WINCOM_DEFAULT_TIMEOUT_SEC）：超时后杀掉进程树，返回失败结果。
    """
    if not is_wincom_supported():
        raise NotImplementedError("Word COM 引擎仅支持 Windows 环境且需安装 pywin32")

    if target not in ("docx", "pdf"):
        raise ValueError(f"Unsupported target: {target}")

    src_p = Path(src).resolve()
    dst_p = Path(dst).resolve()
    dst_p.parent.mkdir(parents=True, exist_ok=True)

    fmt_code = wdFormatXMLDocument if target == "docx" else wdFormatPDF

    own_session = False
    sess = session
    if sess is None or not sess.is_alive():
        sess = WordSession()
        sess.start()
        own_session = True

    t_start = time.time()
    doc = None
    error_msg = ""
    # 独立子进程看门狗：超时后强杀 Word 进程树，不受本进程调度/GIL 影响
    watchdog: subprocess.Popen | None = None
    if isinstance(sess.pid, int) and sess.pid > 0:
        watchdog = _spawn_watchdog(sess.pid, timeout_sec)

    try:
        # 打开源文档：禁用宏警告、静默转换提示、以只读模式无界面打开
        doc = sess.word.Documents.Open(
            FileName=str(src_p),
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
            ConfirmConversions=False,
        )

        # 临时文件策略：SaveAs2 到临时文件，成功后再原子移动到目标路径
        tmp_dst = dst_p.with_name(f"{dst_p.stem}_wincom_{os.getpid()}_{int(time.time()*1000)}.{target}")
        if tmp_dst.exists():
            tmp_dst.unlink(missing_ok=True)

        doc.SaveAs2(FileName=str(tmp_dst), FileFormat=fmt_code)

        try:
            doc.Close(SaveChanges=0)
        except Exception:
            pass
        doc = None

        if not _wait_output(tmp_dst):
            # 输出未落盘：重试一次（打印管线瞬时未就绪场景）
            doc = sess.word.Documents.Open(
                FileName=str(src_p),
                ReadOnly=True,
                AddToRecentFiles=False,
                Visible=False,
                ConfirmConversions=False,
            )
            doc.SaveAs2(FileName=str(tmp_dst), FileFormat=fmt_code)
            try:
                doc.Close(SaveChanges=0)
            except Exception:
                pass
            doc = None
            _wait_output(tmp_dst)

        if tmp_dst.is_file() and tmp_dst.stat().st_size > 0:
            if dst_p.exists():
                if dst_p.is_file():
                    dst_p.unlink()
                elif dst_p.is_dir():
                    shutil.rmtree(dst_p)
            shutil.move(str(tmp_dst), str(dst_p))
        else:
            raise RuntimeError(f"Word COM 未生成有效目标文件: {tmp_dst}")

    except Exception as e:
        # 看门狗已自行退出 = 已触发查杀 = 超时
        timed_out = watchdog is not None and watchdog.poll() is not None
        error_msg = "转换超时" if timed_out else f"Word COM 转换异常: {e}"
        # 超时/严重 COM 异常均重置会话（幂等）
        sess.kill()
        if timed_out and getattr(sess, "kill_failed", None):
            error_msg += "，且 Word 进程无法终止（可能卡在打印驱动调用中，建议重启系统后重试）"
    finally:
        if watchdog is not None and watchdog.poll() is None:
            watchdog.terminate()
        if doc is not None and sess.is_alive():
            try:
                doc.Close(SaveChanges=0)
            except Exception:
                pass
            doc = None

        if own_session:
            sess.stop()

    elapsed = time.time() - t_start
    ok = (error_msg == "") and dst_p.is_file() and (dst_p.stat().st_size > 0)
    return FileResult(
        src=src,
        dst=dst_p if ok else None,
        ok=ok,
        elapsed=elapsed,
        error=error_msg if not ok else "",
    )


def export_pdf_com(
    docx_path: Path,
    pdf_path: Path,
    create_bookmarks: bool = True,
    timeout_sec: int = WINCOM_DEFAULT_TIMEOUT_SEC,
    session: WordSession | None = None,
) -> FileResult:
    """使用 Word COM ExportAsFixedFormat 导出高保真 PDF。

    - create_bookmarks=True 时设置 wdExportCreateHeadingBookmarks (1)，生成标题书签；
    - 支持超时守护。
    """
    if not is_wincom_supported():
        raise NotImplementedError("Word COM 引擎仅支持 Windows 环境且需安装 pywin32")

    src_p = Path(docx_path).resolve()
    dst_p = Path(pdf_path).resolve()
    dst_p.parent.mkdir(parents=True, exist_ok=True)

    bookmark_type = (
        wdExportCreateHeadingBookmarks if create_bookmarks else wdExportCreateNoBookmarks
    )

    own_session = False
    sess = session
    if sess is None or not sess.is_alive():
        sess = WordSession()
        sess.start()
        own_session = True

    t_start = time.time()
    doc = None
    error_msg = ""
    # 独立子进程看门狗：超时后强杀 Word 进程树，不受本进程调度/GIL 影响
    watchdog: subprocess.Popen | None = None
    if isinstance(sess.pid, int) and sess.pid > 0:
        watchdog = _spawn_watchdog(sess.pid, timeout_sec)

    try:
        doc = sess.word.Documents.Open(
            FileName=str(src_p),
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
            ConfirmConversions=False,
        )

        tmp_dst = dst_p.with_name(f"{dst_p.stem}_wincom_{os.getpid()}_{int(time.time()*1000)}.pdf")
        if tmp_dst.exists():
            tmp_dst.unlink(missing_ok=True)

        # 刷新文档域代码（包括 TOC、PAGEREF 动态页码域），确保导出的 PDF 具备真实精确页码
        try:
            doc.Fields.Update()
        except Exception:
            pass

        # 与实测基准 (19.9 min / 3,815 页) 保持一致的精简参数集。
        # 切勿追加 DocStructureTags/BitmapMissingFonts 等参数：
        # BitmapMissingFonts=True 在字体缺失时按位图渲染文本，实测导出吞吐下降数倍导致超时。
        doc.ExportAsFixedFormat(
            OutputFileName=str(tmp_dst),
            ExportFormat=wdExportFormatPDF,
            OpenAfterExport=False,
            OptimizeFor=wdExportOptimizeForPrint,
            CreateBookmarks=bookmark_type,
        )

        try:
            doc.Close(SaveChanges=0)
        except Exception:
            pass
        doc = None

        if not _wait_output(tmp_dst):
            # 输出未落盘：重试一次（打印后台处理程序瞬时未就绪场景）
            doc = sess.word.Documents.Open(
                FileName=str(src_p),
                ReadOnly=True,
                AddToRecentFiles=False,
                Visible=False,
                ConfirmConversions=False,
            )
            try:
                doc.Fields.Update()
            except Exception:
                pass
            doc.ExportAsFixedFormat(
                OutputFileName=str(tmp_dst),
                ExportFormat=wdExportFormatPDF,
                OpenAfterExport=False,
                OptimizeFor=wdExportOptimizeForPrint,
                CreateBookmarks=bookmark_type,
            )
            try:
                doc.Close(SaveChanges=0)
            except Exception:
                pass
            doc = None
            _wait_output(tmp_dst)

        if tmp_dst.is_file() and tmp_dst.stat().st_size > 0:
            if dst_p.exists():
                if dst_p.is_file():
                    dst_p.unlink()
                elif dst_p.is_dir():
                    shutil.rmtree(dst_p)
            shutil.move(str(tmp_dst), str(dst_p))
        else:
            raise RuntimeError(f"Word COM 未生成有效 PDF: {tmp_dst}")

    except Exception as e:
        # 看门狗已自行退出 = 已触发查杀 = 超时
        timed_out = watchdog is not None and watchdog.poll() is not None
        error_msg = "导出PDF超时" if timed_out else f"Word COM 导出PDF异常: {e}"
        sess.kill()
        if timed_out and getattr(sess, "kill_failed", None):
            error_msg += "，且 Word 进程无法终止（可能卡在打印驱动调用中，建议重启系统后重试）"
    finally:
        if watchdog is not None and watchdog.poll() is None:
            watchdog.terminate()
        if doc is not None and sess.is_alive():
            try:
                doc.Close(SaveChanges=0)
            except Exception:
                pass
            doc = None

        if own_session:
            sess.stop()

    elapsed = time.time() - t_start
    ok = (error_msg == "") and dst_p.is_file() and (dst_p.stat().st_size > 0)
    return FileResult(
        src=docx_path,
        dst=dst_p if ok else None,
        ok=ok,
        elapsed=elapsed,
        error=error_msg if not ok else "",
    )


# ---------------------------------------------------------------------------
# WincomRunner: 串行受控调度器 (与 BatchRunner 签名对齐)
# ---------------------------------------------------------------------------

class WincomRunner:
    """Word COM 批量转换调度器。

    特性：
    - 串行调度（jobs 恒为 1），杜绝并发打开多个 Word 实例耗尽物理内存；
    - 对外签名与 BatchRunner.convert 完全一致；
    - 跨文件复用同一 Word 实例（降低启动与销毁开销）；
    - 单文件失败/超时/查杀后自动重置并重新拉起新实例继续处理后续文件；
    - 支持 should_cancel 协作式取消；
    - wincom 导出的 DOCX 属于微软原生合规文档，不做任何 normalize 处理。
    """

    def __init__(self, timeout_sec: int = WINCOM_DEFAULT_TIMEOUT_SEC) -> None:
        if not is_wincom_supported():
            raise NotImplementedError("Word COM 引擎仅支持 Windows 环境且需安装 pywin32")
        self.jobs: int = 1
        self.timeout_sec: int = timeout_sec

    def convert(
        self,
        pairs: Sequence[tuple[Path, Path]],
        target: Literal["docx", "pdf"],
        gray: bool = False,
        progress_cb: Callable[[int, int, FileResult], None] | None = None,
        should_cancel: Callable[[], bool] | None = None,
    ) -> list[FileResult]:
        """按顺序串行转换文档列表。

        参数签名与 BatchRunner.convert 严格一致。
        """
        if not pairs:
            return []

        if not is_wincom_supported():
            raise NotImplementedError("Word COM 引擎仅支持 Windows 环境且需安装 pywin32")

        src_list = [Path(p[0]) for p in pairs]
        total = len(src_list)

        if should_cancel and should_cancel():
            return [
                FileResult(src=src, dst=None, ok=False, elapsed=0.0, error="cancelled")
                for src in src_list
            ]

        results: list[FileResult] = []
        session = WordSession()
        session.start()  # 立即拉起受控 Word 实例，供整批文件复用

        try:
            for idx, (src, dst) in enumerate(pairs):
                # 协作式取消检测
                if should_cancel and should_cancel():
                    # 剩余文件标记为 cancelled
                    for rem_src, _ in pairs[idx:]:
                        results.append(
                            FileResult(src=Path(rem_src), dst=None, ok=False, elapsed=0.0, error="cancelled")
                        )
                    break

                # 确保会话可用；若此前发生查杀则重新启动新 Word 实例
                if not session.is_alive():
                    try:
                        session.start()
                    except Exception as e:
                        res = FileResult(
                            src=Path(src),
                            dst=None,
                            ok=False,
                            elapsed=0.0,
                            error=f"Word 实例启动失败: {e}",
                        )
                        results.append(res)
                        if progress_cb:
                            progress_cb(len(results), total, res)
                        continue

                # 转换处理
                # 如果目标是 PDF 且输入已经是 docx，调用 export_pdf_com 保证带大纲书签；
                # 否则统一使用 convert_file_com (SaveAs2)
                src_path = Path(src)
                dst_path = Path(dst)
                if target == "pdf" and src_path.suffix.lower() == ".docx":
                    res = export_pdf_com(
                        docx_path=src_path,
                        pdf_path=dst_path,
                        create_bookmarks=True,
                        timeout_sec=self.timeout_sec,
                        session=session,
                    )
                else:
                    res = convert_file_com(
                        src=src_path,
                        dst=dst_path,
                        target=target,
                        timeout_sec=self.timeout_sec,
                        session=session,
                    )

                results.append(res)
                if progress_cb:
                    progress_cb(len(results), total, res)

        finally:
            session.stop()

        return results
