from __future__ import annotations

import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Sequence


@dataclass(frozen=True)
class Engine:
    soffice: Path
    source: str  # "bundled" | "path" | "system"
    version: tuple[int, int] | None  # 解析自 `soffice --version`，如 (26, 8)


class EngineNotFoundError(RuntimeError):
    """未找到可用的 LibreOffice / soffice 转换引擎。"""


def _is_executable(p: Path) -> bool:
    try:
        if not p.is_file():
            return False
        if sys.platform != "win32":
            return os.access(p, os.X_OK)
        return True
    except OSError:
        return False


def clear_staged_update(soffice: Path, max_retries: int = 3) -> bool:
    """清除 LibreOffice 相关的更新暂存区及更新状态。

    LibreOffice 启动时若检测到更新暂存区（updating/ 或 Data/settings/updates/0
    下的 pending-service 补丁），会优先唤起 updater.exe 弹出更新对话框并退出，
    导致 --convert-to 等参数被完全忽略，表现为"每次启动都在更新、不干活"。
    同时尝试移除 program/ 下的 updater.exe 与 update_service.exe，从根源阻断更新器。
    """
    removed = False
    prog_dir = soffice.parent
    for updater_bin in ("updater.exe", "update_service.exe", "updater.ini"):
        target_file = prog_dir / updater_bin
        try:
            if target_file.is_file():
                target_file.unlink()
                removed = True
        except OSError:
            pass

    candidate_dirs: set[Path] = set()
    for root in {
        prog_dir,
        prog_dir.parent,
        prog_dir.parent.parent,
        prog_dir.parent.parent.parent,
    }:
        candidate_dirs.add(root / "updating")
        candidate_dirs.add(root / "Data" / "settings" / "updates")

    for staging in candidate_dirs:
        try:
            if not staging.is_dir():
                continue
        except OSError:
            continue
        for attempt in range(max_retries):
            try:
                shutil.rmtree(staging)
                print(
                    f"[提示] 已清除 LibreOffice 更新暂存区，"
                    f"阻止自更新干扰转换: {staging}",
                    file=sys.stderr,
                )
                removed = True
                break
            except OSError:
                if attempt == max_retries - 1:
                    print(
                        f"[警告] 无法清除 LibreOffice 更新暂存区，"
                        f"转换可能被自更新拦截: {staging}",
                        file=sys.stderr,
                    )
                else:
                    time.sleep(0.2)
    return removed


def _parse_version(soffice_path: Path) -> tuple[int, int] | None:
    # 1. 优先尝试从同级目录下的 version.ini 读取，零子进程开销且绝不触发更新器
    try:
        version_ini = soffice_path.parent / "version.ini"
        if version_ini.is_file():
            content = version_ini.read_text(encoding="utf-8", errors="replace")
            m = re.search(r"^ProductKey\s*=\s*LibreOffice\s*(\d+)\.(\d+)", content, re.MULTILINE | re.IGNORECASE)
            if not m:
                m = re.search(r"UpdateID\s*=\s*LibreOffice_(\d+)", content, re.IGNORECASE)
                if m:
                    return int(m.group(1)), 0
            if m:
                return int(m.group(1)), int(m.group(2))
    except Exception:
        pass

    # 2. 从 bootstrap.ini 备用读取 ProductKey
    try:
        bootstrap_ini = soffice_path.parent / "bootstrap.ini"
        if bootstrap_ini.is_file():
            content = bootstrap_ini.read_text(encoding="utf-8", errors="replace")
            m = re.search(r"^ProductKey\s*=\s*LibreOffice\s*(\d+)\.(\d+)", content, re.MULTILINE | re.IGNORECASE)
            if m:
                return int(m.group(1)), int(m.group(2))
    except Exception:
        pass

    # 3. 回退到子进程命令行调用（传入隔离 profile 避免触发更新与污染全局配置）
    try:
        with tempfile.TemporaryDirectory(prefix="wdt_lo_ver_") as tmp_dir:
            profile_url = Path(tmp_dir).resolve().as_uri()
            res = subprocess.run(
                [
                    str(soffice_path),
                    "--headless",
                    "--norestore",
                    "--nolockcheck",
                    "--nodefault",
                    "--nologo",
                    "--nofirststartwizard",
                    f"-env:UserInstallation={profile_url}",
                    "--version",
                ],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                timeout=5,
            )
            if res.returncode != 0:
                return None
            m = re.search(r"(\d+)\.(\d+)", res.stdout)
            if m:
                return int(m.group(1)), int(m.group(2))
    except Exception:
        return None
    return None


def _check_dir_candidates(base_dir: Path) -> Path | None:
    """在目录下检查 soffice 可执行文件，优先 soffice.com。"""
    sub_names = [
        "program/soffice.com",
        "program/soffice.exe",
        "program/soffice",
        "soffice.com",
        "soffice.exe",
        "soffice",
    ]
    for sub in sub_names:
        cand = base_dir / sub
        if _is_executable(cand):
            return cand.resolve()
    return None


def find_bundled_dir() -> Path | None:
    """自动推导内置的便携 LibreOffice 目录（支持打包产物、内部临时目录与当前目录）。"""
    candidates: list[Path] = []
    # 1. 检查可执行文件同级目录 (one-dir 打包典型路径)
    try:
        exe_dir = Path(sys.executable).resolve().parent
        candidates.append(exe_dir / "runtime" / "libreoffice")
        candidates.append(exe_dir / "_internal" / "runtime" / "libreoffice")
    except Exception:
        pass

    # 2. PyInstaller 运行释放目录 (若存在)
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(Path(meipass) / "runtime" / "libreoffice")

    # 3. 当前工作目录下的 runtime/libreoffice
    try:
        candidates.append(Path.cwd() / "runtime" / "libreoffice")
    except Exception:
        pass

    for cand in candidates:
        if cand.is_dir() and _check_dir_candidates(cand):
            return cand
    return None


def find_engine(
    bundled_dir: Path | None = None,
    extra_paths: Sequence[Path] = (),
) -> Engine:
    """智能探测 LibreOffice 引擎。

    探测顺序：
    1. bundled_dir（便携布局，未指定时自动推导）
    2. extra_paths（额外候选路径）
    3. PATH 环境变量
    4. 系统常见安装路径（Linux / Windows / macOS）
    """
    # 1. 检查 bundled_dir
    if bundled_dir is None:
        bundled_dir = find_bundled_dir()

    if bundled_dir is not None:
        b_path = Path(bundled_dir)
        if _is_executable(b_path):
            return _make_engine(b_path.resolve(), "bundled")
        if b_path.is_dir():
            found = _check_dir_candidates(b_path)
            if found:
                return _make_engine(found, "bundled")

    # 2. 检查 extra_paths
    for ep in extra_paths:
        p = Path(ep)
        if _is_executable(p):
            return _make_engine(p.resolve(), "bundled")
        if p.is_dir():
            found = _check_dir_candidates(p)
            if found:
                return _make_engine(found, "bundled")

    # 3. 检查 PATH 环境变量
    path_names = (
        ("soffice.com", "soffice.exe", "soffice", "libreoffice")
        if sys.platform == "win32"
        else ("soffice", "libreoffice")
    )
    for name in path_names:
        found_str = shutil.which(name)
        if found_str:
            return _make_engine(Path(found_str).resolve(), "path")

    # 4. 检查系统常见安装路径
    sys_candidates: list[Path] = []

    # Linux
    opt_matches = sorted(Path("/opt").glob("libreoffice*/program/soffice"), reverse=True)
    sys_candidates.extend(opt_matches)
    for p_str in (
        "/usr/bin/soffice",
        "/usr/bin/libreoffice",
        "/snap/bin/libreoffice",
        "/usr/local/bin/soffice",
    ):
        sys_candidates.append(Path(p_str))

    # Windows
    win_dirs = [
        Path(r"C:\Program Files\LibreOffice\program"),
        Path(r"C:\Program Files (x86)\LibreOffice\program"),
    ]
    prog_files = os.environ.get("ProgramFiles")
    if prog_files:
        win_dirs.append(Path(prog_files) / "LibreOffice" / "program")
    prog_files_x86 = os.environ.get("ProgramFiles(x86)")
    if prog_files_x86:
        win_dirs.append(Path(prog_files_x86) / "LibreOffice" / "program")

    for wd in win_dirs:
        sys_candidates.append(wd / "soffice.com")
        sys_candidates.append(wd / "soffice.exe")

    # macOS
    sys_candidates.append(Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"))

    for cand in sys_candidates:
        if _is_executable(cand):
            return _make_engine(cand.resolve(), "system")

    raise EngineNotFoundError("未能找到 LibreOffice / soffice 可执行文件。")


def _make_engine(soffice: Path, source: str) -> Engine:
    """构建 Engine 前先清除更新暂存区，避免 soffice 启动时走自更新而不执行转换。"""
    clear_staged_update(soffice)
    return Engine(soffice=soffice, source=source, version=_parse_version(soffice))
