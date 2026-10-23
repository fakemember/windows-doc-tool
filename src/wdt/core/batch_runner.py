from __future__ import annotations

import os
import sys
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path
from typing import Callable, Literal, Sequence

from wdt.core.chunk_planner import Chunk, plan_chunks
from wdt.core.converter_worker import FileResult, convert_chunk
from wdt.core.engine_finder import Engine


def _get_physical_memory_gb() -> float:
    """获取系统物理内存大小（GB）。"""
    if hasattr(os, "sysconf"):
        try:
            pages = os.sysconf("SC_PHYS_PAGES")
            page_size = os.sysconf("SC_PAGE_SIZE")
            return (pages * page_size) / (1024 ** 3)
        except (ValueError, OSError):
            pass

    if sys.platform == "win32":
        try:
            import ctypes

            class MEMORYSTATUSEX(ctypes.Structure):
                _fields_ = [
                    ("dwLength", ctypes.c_ulong),
                    ("dwMemoryLoad", ctypes.c_ulong),
                    ("ullTotalPhys", ctypes.c_ulonglong),
                    ("ullAvailPhys", ctypes.c_ulonglong),
                    ("ullTotalPageFile", ctypes.c_ulonglong),
                    ("ullAvailPageFile", ctypes.c_ulonglong),
                    ("ullTotalVirtual", ctypes.c_ulonglong),
                    ("ullAvailVirtual", ctypes.c_ulonglong),
                    ("sullAvailExtendedVirtual", ctypes.c_ulonglong),
                ]

            stat = MEMORYSTATUSEX()
            stat.dwLength = ctypes.sizeof(MEMORYSTATUSEX)
            if ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(stat)):
                return stat.ullTotalPhys / (1024 ** 3)
        except Exception:
            pass

    try:
        import psutil
        return psutil.virtual_memory().total / (1024 ** 3)
    except Exception:
        pass

    return 8.0


def _default_jobs() -> int:
    """计算默认并发进程数。

    min(cpu - 1, 8, max(1, 物理内存GB // 1.5))
    """
    cpu = os.cpu_count() or 4
    mem_gb = _get_physical_memory_gb()
    mem_limit = max(1, int(mem_gb // 1.5))
    cpu_limit = max(1, cpu - 1)
    return max(1, min(cpu_limit, 8, mem_limit))


class BatchRunner:
    """文档批量转换调度器。"""

    def __init__(self, engine: Engine, jobs: int | None = None) -> None:
        self.engine = engine
        self.jobs = max(1, jobs) if jobs is not None else _default_jobs()

    def convert(
        self,
        pairs: Sequence[tuple[Path, Path]],
        target: Literal["docx", "pdf"],
        gray: bool = False,
        progress_cb: Callable[[int, int, FileResult], None] | None = None,
        should_cancel: Callable[[], bool] | None = None,
    ) -> list[FileResult]:
        """批量转换文档。

        - 使用 ProcessPoolExecutor 并发执行分块转换；
        - 某分块失败时，对未成功文件逐个单独重试（独立沙箱降级）；
        - 每完成一个文件在主进程触发 progress_cb(done, total, result)；
        - 按输入 pairs 原始顺序返回 FileResult 列表；
        - 支持 should_cancel() 协作式取消机制。
        """
        if not pairs:
            return []

        src_list = [Path(p[0]) for p in pairs]
        dst_map = {Path(p[0]): Path(p[1]) for p in pairs}
        total = len(src_list)

        if should_cancel and should_cancel():
            return [
                FileResult(src=src, dst=None, ok=False, elapsed=0.0, error="cancelled")
                for src in src_list
            ]

        chunks = plan_chunks(src_list, jobs=self.jobs)
        if not chunks:
            return []

        final_results: dict[Path, FileResult] = {}
        needs_retry: list[Path] = []
        done_count = 0
        cancelled = False

        executor = ProcessPoolExecutor(max_workers=self.jobs)
        try:
            # 1. 提交分块转换
            future_to_chunk = {
                executor.submit(
                    convert_chunk,
                    self.engine,
                    chunk,
                    dst_map,
                    target,
                    gray=gray,
                    worker_id=chunk.index % self.jobs,
                ): chunk
                for chunk in chunks
            }

            for future in as_completed(future_to_chunk):
                chunk = future_to_chunk[future]
                try:
                    chunk_results = future.result()
                except Exception as e:
                    chunk_results = [
                        FileResult(
                            src=f,
                            dst=None,
                            ok=False,
                            elapsed=0.0,
                            error=str(e),
                        )
                        for f in chunk.files
                    ]

                for res in chunk_results:
                    if res.ok:
                        final_results[res.src] = res
                        done_count += 1
                        if progress_cb:
                            progress_cb(done_count, total, res)
                    else:
                        # 单文件分块或版本不支持灰度无需重试，直接记为失败
                        if len(chunk.files) == 1 or "LibreOffice >= 7.4" in res.error:
                            final_results[res.src] = res
                            done_count += 1
                            if progress_cb:
                                progress_cb(done_count, total, res)
                        else:
                            needs_retry.append(res.src)

                if should_cancel and should_cancel():
                    cancelled = True
                    try:
                        executor.shutdown(wait=False, cancel_futures=True)
                    except Exception:
                        pass
                    break

            # 2. 失败降级：逐文件重试（单文件单沙箱）
            if not cancelled and needs_retry:
                retry_chunks = [
                    Chunk(
                        index=i,
                        files=[src],
                        total_bytes=src.stat().st_size if src.is_file() else 0,
                    )
                    for i, src in enumerate(needs_retry)
                ]

                retry_futures = {
                    executor.submit(
                        convert_chunk,
                        self.engine,
                        r_chunk,
                        dst_map,
                        target,
                        gray=gray,
                        worker_id=r_chunk.index % self.jobs,
                    ): r_chunk
                    for r_chunk in retry_chunks
                }

                for future in as_completed(retry_futures):
                    r_chunk = retry_futures[future]
                    src = r_chunk.files[0]
                    try:
                        r_results = future.result()
                        res = (
                            r_results[0]
                            if r_results
                            else FileResult(
                                src=src,
                                dst=None,
                                ok=False,
                                elapsed=0.0,
                                error="重试未返回结果",
                            )
                        )
                    except Exception as e:
                        res = FileResult(
                            src=src,
                            dst=None,
                            ok=False,
                            elapsed=0.0,
                            error=f"重试失败: {e}",
                        )

                    final_results[src] = res
                    done_count += 1
                    if progress_cb:
                        progress_cb(done_count, total, res)

                    if should_cancel and should_cancel():
                        cancelled = True
                        try:
                            executor.shutdown(wait=False, cancel_futures=True)
                        except Exception:
                            pass
                        break
        finally:
            try:
                executor.shutdown(wait=not cancelled, cancel_futures=cancelled)
            except Exception:
                pass

        default_err = "cancelled" if cancelled else "未处理"
        return [
            final_results.get(
                src,
                FileResult(src=src, dst=None, ok=False, elapsed=0.0, error=default_err),
            )
            for src in src_list
        ]
