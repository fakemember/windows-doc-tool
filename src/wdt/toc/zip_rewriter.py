"""
ZIP 重写器：针对小文件采用 zipfile 原位原子重建；针对超大文件（>100MB）采用 7z 多线程增量更新与磁盘余量预检。
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Mapping


def _find_seven_zip(custom_path: Path | None = None) -> str | None:
    """寻找可用的 7za 或 7z 命令路径（优先检查打包/便携运行时内置组件）"""
    if custom_path is not None:
        if Path(custom_path).is_file():
            return str(custom_path)
        found = shutil.which(str(custom_path))
        if found:
            return found

    # 1. 优先检查打包运行时内置的 7za.exe / 7za
    bundled_candidates: list[Path] = []
    try:
        exe_dir = Path(sys.executable).resolve().parent
        bundled_candidates.append(exe_dir / "runtime" / "7za.exe")
        bundled_candidates.append(exe_dir / "runtime" / "7za")
        bundled_candidates.append(exe_dir / "_internal" / "runtime" / "7za.exe")
    except Exception:
        pass

    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        bundled_candidates.append(Path(meipass) / "runtime" / "7za.exe")

    try:
        bundled_candidates.append(Path.cwd() / "runtime" / "7za.exe")
        bundled_candidates.append(Path.cwd() / "runtime" / "7za")
    except Exception:
        pass

    for cand in bundled_candidates:
        if cand.is_file():
            return str(cand.resolve())

    # 2. 查找 PATH 环境变量中的 7za / 7z
    for cmd in ("7za", "7z"):
        found = shutil.which(cmd)
        if found:
            return found
    return None


def rewrite_zip(
    docx_path: Path,
    replacements: Mapping[str, bytes],
    seven_zip: Path | None = None,
    big_threshold: int = 100 * 1024 * 1024,
    tmp_dir: Path | None = None,
) -> None:
    """以原子且安全的方式更新 docx 压缩包中的指定文件。

    :param docx_path: docx 文件路径
    :param replacements: 待更新或新增的条目映射（如 {"word/document.xml": b"..."}）
    :param seven_zip: 可选 7z 可执行程序路径
    :param big_threshold: 超大文件判定阈值（字节），默认 100MB
    :param tmp_dir: 临时工作目录
    """
    docx_path = Path(docx_path).resolve()
    if not docx_path.is_file():
        raise FileNotFoundError(f"docx 文件不存在: {docx_path}")

    file_size = docx_path.stat().st_size
    seven_zip_cmd = _find_seven_zip(seven_zip)

    use_seven_zip = (file_size > big_threshold) and (seven_zip_cmd is not None)

    if not use_seven_zip:
        _rewrite_via_zipfile(docx_path, replacements)
    else:
        _rewrite_via_seven_zip(
            docx_path, replacements, seven_zip_cmd, file_size, tmp_dir
        )


def _rewrite_via_zipfile(docx_path: Path, replacements: Mapping[str, bytes]) -> None:
    """基于标准库 zipfile 全量原子重建压缩包"""
    staging_file = None
    try:
        with tempfile.NamedTemporaryFile(
            dir=docx_path.parent, prefix="wdt_zip_", suffix=".tmp", delete=False
        ) as tmp:
            staging_file = Path(tmp.name)

        written_names: set[str] = set()

        with zipfile.ZipFile(docx_path, "r") as src_zip:
            with zipfile.ZipFile(staging_file, "w", compression=zipfile.ZIP_DEFLATED) as dst_zip:
                for item in src_zip.infolist():
                    if item.filename in replacements:
                        dst_zip.writestr(item.filename, replacements[item.filename])
                        written_names.add(item.filename)
                    else:
                        dst_zip.writestr(item, src_zip.read(item.filename))

                # 补充原 zip 中不存在的新增条目
                for name, data in replacements.items():
                    if name not in written_names:
                        dst_zip.writestr(name, data)

        os.replace(staging_file, docx_path)
        staging_file = None
    finally:
        if staging_file and staging_file.exists():
            try:
                staging_file.unlink()
            except OSError:
                pass


def _rewrite_via_seven_zip(
    docx_path: Path,
    replacements: Mapping[str, bytes],
    seven_zip_cmd: str,
    file_size: int,
    tmp_dir: Path | None,
) -> None:
    """基于 7z 多线程增量更新超大压缩包，先校验磁盘空间并做 staging 复制"""
    # 磁盘余量预检：需 >= 1.3 * 文件大小
    required_space = int(1.3 * file_size)
    target_dir = docx_path.parent
    usage = shutil.disk_usage(target_dir)
    if usage.free < required_space:
        raise RuntimeError(
            f"磁盘空间不足: 当前可用 {usage.free} 字节，原位更新需要至少 {required_space} 字节"
        )

    staging_file = None
    tmp_work_dir = None

    try:
        # 复制原 docx 到 staging 临时文件
        with tempfile.NamedTemporaryFile(
            dir=target_dir, prefix="wdt_7z_", suffix=".tmp", delete=False
        ) as tmp:
            staging_file = Path(tmp.name)
        shutil.copy2(docx_path, staging_file)

        # 准备待更新文件目录
        tmp_work_dir = tempfile.mkdtemp(dir=tmp_dir, prefix="wdt_7z_work_")
        work_path = Path(tmp_work_dir)

        relative_names: list[str] = []
        for rel_name, data in replacements.items():
            file_dest = work_path / rel_name
            file_dest.parent.mkdir(parents=True, exist_ok=True)
            file_dest.write_bytes(data)
            relative_names.append(rel_name)

        # 执行 7z 更新
        cmd = [
            seven_zip_cmd,
            "u",
            "-tzip",
            "-mmt=on",
            "-mx=6",
            str(staging_file),
            *relative_names,
        ]
        res = subprocess.run(
            cmd,
            cwd=str(work_path),
            capture_output=True,
            text=True,
            check=False,
        )

        if res.returncode != 0:
            raise RuntimeError(f"7z 原位更新失败 (退出码 {res.returncode}): {res.stderr}")

        # 原子覆盖
        os.replace(staging_file, docx_path)
        staging_file = None
    finally:
        if staging_file and staging_file.exists():
            try:
                staging_file.unlink()
            except OSError:
                pass
        if tmp_work_dir and os.path.exists(tmp_work_dir):
            try:
                shutil.rmtree(tmp_work_dir)
            except OSError:
                pass
