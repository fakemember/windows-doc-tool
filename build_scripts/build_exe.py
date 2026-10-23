#!/usr/bin/env python3
"""PyInstaller one-dir 目录模式构建脚本。

注意：
- 严格禁止 one-file 单文件打包模式：内置便携运行时体积约 100MB+，
  单文件模式每次冷启动都会解压到系统 Temp 目录，耗时不可接受。
- 本脚本采用 one-dir 模式，将主程序及可选的 runtime/ 打包为绿色文件夹。
- 支持在 Linux 开发机上通过 --dry-run 预览将要执行的完整构建命令。
"""

from __future__ import annotations

import argparse
import os
from pathlib import Path
import shlex
import subprocess
import sys


def build_pyinstaller_args(
    entry: Path,
    runtime_dir: Path | None = None,
    dist_dir: Path | None = None,
    work_dir: Path | None = None,
    name: str = "wdt",
) -> list[str]:
    """组装 PyInstaller one-dir 打包参数列表。"""
    entry = Path(entry)
    dist_dir = Path(dist_dir) if dist_dir else Path("dist")
    work_dir = Path(work_dir) if work_dir else Path("build")

    args = [
        "--noconfirm",
        "--clean",
        "--windowed",
        "--name",
        name,
        "--onedir",
        "--distpath",
        str(dist_dir),
        "--workpath",
        str(work_dir),
    ]

    # 如果指定了 runtime 且目录存在，则作为数据目录收集到产物根下的 runtime/
    if runtime_dir is not None:
        runtime_path = Path(runtime_dir)
        if runtime_path.exists() and runtime_path.is_dir():
            # Windows 分隔符为 ';'，POSIX 系统为 ':'
            sep = os.pathsep
            args.extend(["--add-data", f"{runtime_path.resolve()}{sep}runtime"])
        else:
            print(
                f"警告: 运行时目录 '{runtime_path}' 不存在，跳过收集 runtime。"
                " (源码阶段开发期允许)",
                file=sys.stderr,
            )

    args.append(str(entry.resolve() if entry.exists() else entry))
    return args


def build_exe(
    entry: Path,
    runtime_dir: Path | None = None,
    dist_dir: Path | None = None,
    work_dir: Path | None = None,
    name: str = "wdt",
    dry_run: bool = False,
) -> int:
    """执行 PyInstaller 打包构建。

    Args:
        entry: 程序入口文件路径 (默认 src/wdt/main.py)
        runtime_dir: 便携运行时目录 (若存在则自动打包入 runtime/)
        dist_dir: 产物输出目录 (默认 dist/)
        work_dir: 编译中间缓存目录 (默认 build/)
        name: 生成的可执行文件名 (默认 wdt)
        dry_run: 若为 True，仅打印构建命令，不执行真实编译

    Returns:
        0 表示成功，非 0 表示构建失败。
    """
    args = build_pyinstaller_args(
        entry=entry,
        runtime_dir=runtime_dir,
        dist_dir=dist_dir,
        work_dir=work_dir,
        name=name,
    )

    if dry_run:
        cmd_str = "pyinstaller " + " ".join(shlex.quote(a) for a in args)
        print("[dry-run] PyInstaller 构建命令预览 (one-dir 模式):")
        print(f"  {cmd_str}")
        return 0

    # 优先尝试导入 PyInstaller API
    try:
        import PyInstaller.__main__  # type: ignore

        print(f"正在调用 PyInstaller API 执行构建: name={name}...")
        PyInstaller.__main__.run(args)
        return 0
    except ImportError:
        # 回退到子进程命令行调用
        print("未检测到已导入的 PyInstaller，尝试通过子进程调用 pyinstaller...")
        cmd = [sys.executable, "-m", "PyInstaller"] + args
        try:
            res = subprocess.run(cmd, check=False)
            return res.returncode
        except FileNotFoundError:
            print(
                "错误: 当前环境中未安装 PyInstaller。"
                "请使用 'pip install pyinstaller' 安装，或使用 '--dry-run' 预览命令。",
                file=sys.stderr,
            )
            return 1


def main(argv: list[str] | None = None) -> int:
    project_root = Path(__file__).resolve().parent.parent

    parser = argparse.ArgumentParser(
        description="Windows Doc Tool (wdt) PyInstaller one-dir 目录打包构建脚本"
    )
    parser.add_argument(
        "--entry",
        type=Path,
        default=project_root / "src" / "wdt" / "main.py",
        help="程序入口源码文件路径 (默认: src/wdt/main.py)",
    )
    parser.add_argument(
        "--runtime-dir",
        type=Path,
        default=project_root / "runtime",
        help="便携运行时组件目录 (默认: runtime/)",
    )
    parser.add_argument(
        "--distpath",
        type=Path,
        default=project_root / "dist",
        help="输出产物目录 (默认: dist/)",
    )
    parser.add_argument(
        "--workpath",
        type=Path,
        default=project_root / "build",
        help="编译中间文件目录 (默认: build/)",
    )
    parser.add_argument(
        "--name",
        type=str,
        default="wdt",
        help="输出应用程序名称 (默认: wdt)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="仅打印将要执行的 PyInstaller 打包命令，不执行真实构建",
    )
    args = parser.parse_args(argv)

    return build_exe(
        entry=args.entry,
        runtime_dir=args.runtime_dir,
        dist_dir=args.distpath,
        work_dir=args.workpath,
        name=args.name,
        dry_run=args.dry_run,
    )


if __name__ == "__main__":
    sys.exit(main())
