from __future__ import annotations

import os
import re
import shutil
import subprocess
import sys
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


def _parse_version(soffice_path: Path) -> tuple[int, int] | None:
    try:
        res = subprocess.run(
            [str(soffice_path), "--version"],
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


def find_engine(
    bundled_dir: Path | None = None,
    extra_paths: Sequence[Path] = (),
) -> Engine:
    """智能探测 LibreOffice 引擎。

    探测顺序：
    1. bundled_dir（便携布局）
    2. extra_paths（额外候选路径）
    3. PATH 环境变量
    4. 系统常见安装路径（Linux / Windows / macOS）
    """
    # 1. 检查 bundled_dir
    if bundled_dir is not None:
        b_path = Path(bundled_dir)
        if _is_executable(b_path):
            soffice = b_path.resolve()
            return Engine(soffice=soffice, source="bundled", version=_parse_version(soffice))
        if b_path.is_dir():
            found = _check_dir_candidates(b_path)
            if found:
                return Engine(soffice=found, source="bundled", version=_parse_version(found))

    # 2. 检查 extra_paths
    for ep in extra_paths:
        p = Path(ep)
        if _is_executable(p):
            soffice = p.resolve()
            return Engine(soffice=soffice, source="bundled", version=_parse_version(soffice))
        if p.is_dir():
            found = _check_dir_candidates(p)
            if found:
                return Engine(soffice=found, source="bundled", version=_parse_version(found))

    # 3. 检查 PATH 环境变量
    path_names = (
        ("soffice.com", "soffice.exe", "soffice", "libreoffice")
        if sys.platform == "win32"
        else ("soffice", "libreoffice")
    )
    for name in path_names:
        found_str = shutil.which(name)
        if found_str:
            soffice = Path(found_str).resolve()
            return Engine(soffice=soffice, source="path", version=_parse_version(soffice))

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
            soffice = cand.resolve()
            return Engine(soffice=soffice, source="system", version=_parse_version(soffice))

    raise EngineNotFoundError("未能找到 LibreOffice / soffice 可执行文件。")
