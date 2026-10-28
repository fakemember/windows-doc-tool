from __future__ import annotations

import os
import re
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
from wdt.core.engine_finder import Engine, clear_staged_update
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

# 预置给每个独立 UserInstallation 的最小注册表配置：
# 1. 显式禁用 MAR 自动更新器 (Update/Enabled = false)，阻止每次新建沙箱 profile 时发起在线更新检查与下载；
# 2. 禁用首次运行初始化与 OpenCL 编译检测 (FirstRun = false, UseOpenCL = false)，避免无头转换被后台初始化拖慢或中断。
_PROFILE_REGISTRY_XCU = """\
<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Common/Misc"><prop oor:name="FirstRun" oor:op="fuse"><value>false</value></prop></item>
<item oor:path="/org.openoffice.Office.Common/Misc"><prop oor:name="UseOpenCL" oor:op="fuse"><value>false</value></prop></item>
<item oor:path="/org.openoffice.Office.Update/Update"><prop oor:name="Enabled" oor:op="fuse"><value>false</value></prop></item>
<item oor:path="/org.openoffice.Office.Update/Update"><prop oor:name="UpdateRunning" oor:op="fuse"><value>false</value></prop></item>
<item oor:path="/org.openoffice.Setup/Office"><prop oor:name="ooSetupInstCompleted" oor:op="fuse"><value>true</value></prop></item>
</oor:items>
"""


def _prepare_profile_dir(profile_dir: Path) -> None:
    """初始化沙箱 UserInstallation 目录并写入禁用自动更新与首次运行检查的配置。"""
    user_dir = profile_dir / "user"
    user_dir.mkdir(parents=True, exist_ok=True)
    (user_dir / "registrymodifications.xcu").write_text(
        _PROFILE_REGISTRY_XCU, encoding="utf-8"
    )


# ---------------------------------------------------------------------------
# SAS RTF 头部声明修复：部分 SAS 输出声明 {\rtf1\ansi 却缺失 \ansicpg936
# （GBK 代码页），且只使用内联 \s1 等样式标记而没有任何 {\stylesheet} 定义。
# 后果：LibreOffice 按 Windows-1252 解码中文导致乱码；转换后标题段落全部
# 退化为 Normal 样式，OOXML 标题检测器无法识别。此处检测并在头部注入缺失
# 声明后转换为暂存副本：
# - \ansicpg936：还原 GBK 文本；
# - {\stylesheet ... \outlinelevelN ... heading N}：让 \sN 标记映射为真正的
#   Heading N 样式（带大纲级别），转换后现有标题检测器零改动可用。
# ---------------------------------------------------------------------------

_RTF_HEADER_SAMPLE = 8192
_RTF_BODY_SAMPLE = 256 * 1024
_RTF_HEX_ESCAPE_RE = re.compile(rb"\\'([0-9a-fA-F]{2})")
# 段落样式引用 \sN（排除 \cs 字符样式与多位数 \s10+；SAS 仅用 1~9 级）
_RTF_SREF_RE = re.compile(rb"(?<![a-zA-Z])\\s([1-9])(?![0-9])")


def _rtf_dbcs_candidates(path: Path) -> bytes:
    """抽样收集 RTF 正文中的双字节候选字节（\'xx 转义值 + 原始高位字节）。"""
    try:
        with open(path, "rb") as f:
            data = f.read(_RTF_BODY_SAMPLE)
    except OSError:
        return b""
    escaped = bytes(int(m.group(1), 16) for m in _RTF_HEX_ESCAPE_RE.finditer(data))
    stripped = _RTF_HEX_ESCAPE_RE.sub(b"", data)
    raw = bytes(b for b in stripped if b >= 0x80)
    return escaped + raw


def _looks_like_gbk(candidate: bytes) -> bool:
    """候选字节流能否按 GB18030 完整解码且以 CJK 汉字为主。"""
    if len(candidate) < 8:
        return False
    try:
        text = candidate.decode("gb18030")
    except UnicodeDecodeError:
        return False
    meaningful = sum(1 for ch in text if not ch.isspace())
    if meaningful == 0:
        return False
    cjk = sum(1 for ch in text if "一" <= ch <= "鿿")
    return cjk / meaningful >= 0.5


def _stylesheet_fragment(levels: Sequence[int]) -> bytes:
    r"""按引用到的样式级别构建 {\stylesheet} 块，\sN -> heading N（outlinelevel N-1）。"""
    parts = []
    for lvl in sorted(set(levels)):
        font_size = max(16, 24 - (lvl - 1) * 2)
        parts.append(
            b"{\\s%d\\keepn\\sb240\\sa60\\b\\fs%d\\outlinelevel%d heading %d;}"
            % (lvl, font_size, lvl - 1, lvl)
        )
    return b"{\\stylesheet" + b"".join(parts) + b"}"


def _rtf_missing_declarations(path: Path) -> tuple[bool, list[int]]:
    """检测 RTF 缺失的头部声明，返回 (是否缺 \\ansicpg, 被引用但未定义的 \\sN 级别)。"""
    try:
        with open(path, "rb") as f:
            head = f.read(_RTF_HEADER_SAMPLE)
    except OSError:
        return False, []
    if b"{\\rtf1\\ansi" not in head:
        return False, []

    needs_ansicpg = b"\\ansicpg" not in head and _looks_like_gbk(
        _rtf_dbcs_candidates(path)
    )

    levels: list[int] = []
    if b"{\\stylesheet" not in head:
        try:
            with open(path, "rb") as f:
                body = f.read(_RTF_BODY_SAMPLE)
        except OSError:
            body = b""
        levels = sorted({int(m.group(1)) for m in _RTF_SREF_RE.finditer(body)})

    return needs_ansicpg, levels


def _rtf_needs_staging(path: Path) -> bool:
    """判断 RTF 是否缺失需注入的头部声明。"""
    needs_ansicpg, levels = _rtf_missing_declarations(path)
    return needs_ansicpg or bool(levels)


def _stage_rtf_fixed(src: Path, stage_dir: Path) -> Path:
    """复制 RTF 到暂存目录，在 {\rtf1\\ansi 锚点后注入缺失的头部声明。"""
    needs_ansicpg, levels = _rtf_missing_declarations(src)
    stage_dir.mkdir(parents=True, exist_ok=True)
    dst = stage_dir / src.name
    anchor = b"{\\rtf1\\ansi"
    with open(src, "rb") as fin, open(dst, "wb") as fout:
        head = fin.read(_RTF_HEADER_SAMPLE)
        idx = head.find(anchor)
        insert_pos = idx + len(anchor) if idx != -1 else len(head)
        fout.write(head[:insert_pos])
        if needs_ansicpg:
            fout.write(b"\\ansicpg936")
        if levels:
            fout.write(_stylesheet_fragment(levels))
        fout.write(head[insert_pos:])
        shutil.copyfileobj(fin, fout)
    return dst


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
    """按处理量与源格式估算转换超时。

    基准实测：含高密表格的 SAS Listing（3,740 页、100,000 行表格）排版
    开销约 13.5~18 s/MiB（rtf/doc 源按磁盘体积计）；docx 源按解压后
    document.xml 体积计约 1.35 s/MiB。统一取 20 s/MiB（rtf/doc）与
    2 s/MiB（docx）留余量，保底 300s。
    """
    rtf_doc_mb = sum(_workload_mb(f) for f in files if sniff(f) != "docx")
    docx_mb = sum(_workload_mb(f) for f in files if sniff(f) == "docx")
    return int(300 + 20 * rtf_doc_mb + 2 * docx_mb)


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
                _prepare_profile_dir(profile_dir)
                outdir.mkdir(parents=True, exist_ok=True)

                # 清除可能在此期间 stage 的更新暂存区，防止 soffice 走自更新而忽略转换
                clear_staged_update(engine.soffice)

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

                # SAS RTF 头部声明修复：缺失 \ansicpg 或缺失 {\stylesheet} 映射的文件，
                # 转换为注入补齐声明的暂存副本（文件名不变，产物命名不受影响）
                if fmt == "rtf":
                    stage_dir = Path(temp_dir) / "rtf_staged"
                    batch_inputs = [
                        _stage_rtf_fixed(src, stage_dir)
                        if _rtf_needs_staging(src)
                        else src
                        for src in batch
                    ]
                else:
                    batch_inputs = list(batch)

                cmd.extend([
                    "--convert-to", target_arg,
                    "--outdir", str(outdir),
                ])
                cmd.extend(str(f) for f in batch_inputs)

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
