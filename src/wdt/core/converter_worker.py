from __future__ import annotations

import os
import shutil
import signal
import subprocess
import sys
import tempfile
import time
import uuid
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Literal, Mapping, Sequence

from wdt.core.chunk_planner import Chunk
from wdt.core.engine_finder import Engine
from wdt.core.format_sniffer import sniff


@dataclass
class FileResult:
    src: Path
    dst: Path | None
    ok: bool
    elapsed: float
    error: str


def _kill_process_tree(pid: int) -> None:
    """跨平台终止进程及其子进程树。"""
    if sys.platform == "win32":
        try:
            subprocess.run(
                ["taskkill", "/PID", str(pid), "/T", "/F"],
                capture_output=True,
                check=False,
            )
        except Exception:
            pass
    else:
        try:
            pgid = os.getpgid(pid)
            os.killpg(pgid, signal.SIGKILL)
        except (ProcessLookupError, PermissionError):
            try:
                os.kill(pid, signal.SIGKILL)
            except OSError:
                pass
        except Exception:
            pass


def _partition_by_unique_stem(files: list[Path]) -> list[list[Path]]:
    """将同组文件拆分，保证单次 soffice 调用中文件名 stem 不重复。"""
    batches: list[list[Path]] = []
    for f in files:
        placed = False
        for b in batches:
            if not any(existing.stem == f.stem for existing in b):
                b.append(f)
                placed = True
                break
        if not placed:
            batches.append([f])
    return batches


_MIB = 1024 * 1024


def _workload_mb(path: Path) -> float:
    """估算 LibreOffice 实际排版处理的文本量（MiB）。

    docx 是压缩容器，磁盘体积会严重低估真实处理量——实测高压缩比 Listing
    的 document.xml 解压后 431MB 但 docx 仅 6.7MB，按磁盘大小设超时必被误杀。
    因此对 docx 读取 zip 中央目录中 word/document.xml 的解压后大小（无需解压）；
    其余格式无压缩容器，按磁盘大小估算。
    """
    if sniff(path) == "docx":
        try:
            with zipfile.ZipFile(path) as z:
                return z.getinfo("word/document.xml").file_size / _MIB
        except (zipfile.BadZipFile, KeyError):
            pass
    return path.stat().st_size / _MIB


def default_timeout_sec(files: Sequence[Path]) -> int:
    """按处理量估算转换超时。

    基准实测（2,900 页真实 Listing）：RTF→docx 约 1.27s/MiB，docx→PDF
    约 1.35s/MiB，统一取 2s/MiB 留余量，保底 300s。
    """
    return int(300 + 2 * sum(_workload_mb(f) for f in files))


def convert_chunk(
    engine: Engine,
    chunk: Chunk,
    dst_map: Mapping[Path, Path],
    target: Literal["docx", "pdf"],
    gray: bool = False,
    timeout_sec: int | None = None,
    profile_root: Path | None = None,
    worker_id: int = 0,
) -> list[FileResult]:
    """转换一个 Chunk 内的所有文件。

    - 按 format_sniffer.sniff() 结果对文件分组，每组单次调用 soffice；
    - RTF 组添加 --infilter="Rich Text Format"；
    - 每次调用使用独立的临时沙箱 UserInstallation 和临时输出目录 outdir；
    - 进程树超时守护，异常/超时强杀子进程树；
    - 完成后按 dst_map 将产物移动到目标路径。
    """
    if not chunk.files:
        return []

    if target not in ("docx", "pdf"):
        raise ValueError(f"Unsupported target: {target}")

    # LibreOffice < 7.4 不支持 JSON 灰度过滤器
    if gray and engine.version is not None and engine.version < (7, 4):
        return [
            FileResult(
                src=f,
                dst=None,
                ok=False,
                elapsed=0.0,
                error="LibreOffice >= 7.4 required for gray",
            )
            for f in chunk.files
        ]

    if timeout_sec is None:
        timeout_sec = default_timeout_sec(chunk.files)

    # 确定目标格式参数
    if target == "docx":
        target_arg = "docx:MS Word 2007 XML"
    elif gray:
        target_arg = 'pdf:writer_pdf_Export:{"ConvertToGrayScale":{"type":"boolean","value":"true"}}'
    else:
        target_arg = "pdf"

    # 按嗅探格式分组
    groups: dict[str, list[Path]] = {}
    for f in chunk.files:
        fmt = sniff(f)
        groups.setdefault(fmt, []).append(f)

    results: dict[Path, FileResult] = {}
    start_time = time.time()
    deadline = start_time + max(1, timeout_sec)

    for fmt, group_files in groups.items():
        # 拆分为无 stem 重复的批次
        batches = _partition_by_unique_stem(group_files)

        for batch in batches:
            rem_timeout = max(1, int(deadline - time.time()))
            t_batch_start = time.time()

            # 独占沙箱目录与输出目录
            temp_parent = str(profile_root) if profile_root else tempfile.gettempdir()
            unique_tag = f"lo_w{worker_id}_{uuid.uuid4().hex[:8]}"
            temp_dir = tempfile.mkdtemp(prefix=f"{unique_tag}_", dir=temp_parent)

            try:
                profile_dir = Path(temp_dir) / "profile"
                outdir = Path(temp_dir) / "out"
                profile_dir.mkdir(parents=True, exist_ok=True)
                outdir.mkdir(parents=True, exist_ok=True)

                cmd = [
                    str(engine.soffice),
                    "--headless",
                    "--norestore",
                    "--nolockcheck",
                    "--nodefault",
                    "--nologo",
                    "--nofirststartwizard",
                    f"-env:UserInstallation={profile_dir.as_uri()}",
                ]

                if fmt == "rtf":
                    cmd.append("--infilter=Rich Text Format")

                cmd.extend([
                    "--convert-to", target_arg,
                    "--outdir", str(outdir),
                ])
                cmd.extend(str(f) for f in batch)

                popen_kwargs = {}
                if sys.platform != "win32":
                    popen_kwargs["start_new_session"] = True

                proc = subprocess.Popen(
                    cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    **popen_kwargs,
                )

                try:
                    stdout, stderr = proc.communicate(timeout=rem_timeout)
                    elapsed = time.time() - t_batch_start
                    timed_out = False
                except subprocess.TimeoutExpired:
                    _kill_process_tree(proc.pid)
                    try:
                        proc.communicate(timeout=3)
                    except Exception:
                        pass
                    elapsed = time.time() - t_batch_start
                    stdout, stderr = "", "转换超时"
                    timed_out = True
                except Exception as e:
                    _kill_process_tree(proc.pid)
                    try:
                        proc.communicate(timeout=3)
                    except Exception:
                        pass
                    elapsed = time.time() - t_batch_start
                    stdout, stderr = "", str(e)
                    timed_out = False

                for src in batch:
                    if timed_out:
                        results[src] = FileResult(
                            src=src,
                            dst=None,
                            ok=False,
                            elapsed=elapsed,
                            error="转换超时",
                        )
                        continue

                    gen_file = outdir / f"{src.stem}.{target}"
                    if gen_file.is_file() and gen_file.stat().st_size > 0:
                        dst = dst_map.get(src) or dst_map.get(Path(src))
                        if dst is None:
                            dst = src.with_suffix(f".{target}")
                        dst = Path(dst)
                        dst.parent.mkdir(parents=True, exist_ok=True)
                        if dst.exists():
                            if dst.is_file():
                                dst.unlink()
                            elif dst.is_dir():
                                shutil.rmtree(dst)
                        shutil.move(str(gen_file), str(dst))
                        results[src] = FileResult(
                            src=src,
                            dst=dst,
                            ok=True,
                            elapsed=elapsed,
                            error="",
                        )
                    else:
                        err_msg = stderr.strip() if stderr else (stdout.strip() if stdout else "未生成目标文件")
                        results[src] = FileResult(
                            src=src,
                            dst=None,
                            ok=False,
                            elapsed=elapsed,
                            error=err_msg,
                        )

            finally:
                shutil.rmtree(temp_dir, ignore_errors=True)

    # 按照输入 chunk.files 的顺序返回结果
    return [
        results.get(
            f,
            FileResult(
                src=f,
                dst=None,
                ok=False,
                elapsed=0.0,
                error="未知处理错误",
            ),
        )
        for f in chunk.files
    ]
