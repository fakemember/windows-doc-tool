#!/usr/bin/env python3
"""将官方 LibreOffice 便携版目录裁剪为批处理最小运行时。

裁剪策略：
- 只裁剪 `share/` 层的冗余资源（帮助文档、模板库、扩展组件、非 en*.dat/zh*.dat 词典）。
- 绝对保留 `program/` 目录全部核心组件与 DLL 依赖。
- 绝对保留 `share/program/`（若存在）。
- 裁剪后可选执行无头 RTF->DOCX 冒烟验证，失败则自动回滚并报警。
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
from typing import Callable

# 最简 RTF 测试片段（用于冒烟验证）
MINIMAL_RTF = (
    b"{\\rtf1\\ansi\\deff0"
    b"{\\fonttbl{\\f0\\fnil\\fcharset0 Arial;}}"
    b"\\viewkind4\\uc1\\pard\\lang2052\\fs20 Minimal RTF Smoke Test\\par"
    b"}"
)


def _is_non_enzh_dict_item(rel: Path) -> bool:
    """判定 share/dict 下的项目是否为非 en/zh 词典项。"""
    parts = rel.parts
    if len(parts) <= 2:
        # share/dict 根目录本身不删，保留以容纳 en/zh 词典
        return False
    name = rel.name.lower()
    # 保留 en*.dat, zh*.dat 以及常见 en/zh 拼写词典与子目录
    if name.startswith("en") or name.startswith("zh"):
        return False
    return True


@dataclass(frozen=True)
class PruneRule:
    """裁剪规则定义。"""

    name: str
    description: str
    pattern: str
    matcher: Callable[[Path], bool]


# 显式规则列表：只裁剪 share/ 下安全且占体积的组件
PRUNE_RULES: list[PruneRule] = [
    PruneRule(
        name="help",
        description="帮助文档 (share/help)",
        pattern="share/help/**",
        matcher=lambda rel: len(rel.parts) >= 2 and rel.parts[0] == "share" and rel.parts[1] == "help",
    ),
    PruneRule(
        name="template",
        description="内置模板库 (share/template, share/templates)",
        pattern="share/template{,s}/**",
        matcher=lambda rel: len(rel.parts) >= 2
        and rel.parts[0] == "share"
        and rel.parts[1] in ("template", "templates"),
    ),
    PruneRule(
        name="extensions",
        description="扩展插件 (share/extensions)",
        pattern="share/extensions/**",
        matcher=lambda rel: len(rel.parts) >= 2
        and rel.parts[0] == "share"
        and rel.parts[1] == "extensions",
    ),
    PruneRule(
        name="dict_non_enzh",
        description="非 en*.dat/zh*.dat 词典项 (share/dict)",
        pattern="share/dict/** (!en*.dat, !zh*.dat)",
        matcher=lambda rel: len(rel.parts) >= 2
        and rel.parts[0] == "share"
        and rel.parts[1] in ("dict", "dictionaries", "wordbook")
        and _is_non_enzh_dict_item(rel),
    ),
]


def prune_plan(source: Path) -> list[Path]:
    """计算待删除的相对路径清单（纯函数，无副作用）。

    Args:
        source: LibreOffice 便携版根目录。

    Returns:
        待删除的相对路径列表（按字典序升序）。
    """
    source = Path(source)
    if not source.exists():
        return []

    plan: list[Path] = []
    for item in source.rglob("*"):
        rel = item.relative_to(source)
        parts = rel.parts
        # 严格安全边界：非 share/ 目录绝对不删（例如 program/、App/ 等）
        if not parts or parts[0] != "share":
            continue
        # share/program 绝对不删
        if len(parts) >= 2 and parts[1] == "program":
            continue

        for rule in PRUNE_RULES:
            if rule.matcher(rel):
                plan.append(rel)
                break

    return sorted(plan)


def get_dir_size_bytes(path: Path) -> int:
    """统计目录所有常规文件的总字节数。"""
    if not path.exists():
        return 0
    total = 0
    for p in path.rglob("*"):
        if p.is_file() and not p.is_symlink():
            try:
                total += p.stat().st_size
            except OSError:
                pass
    return total


def find_bundled_soffice(output_dir: Path) -> Path | None:
    """在裁剪输出目录中探测可执行 soffice 文件。"""
    program_dir = output_dir / "program"
    candidates = [
        program_dir / "soffice.com",
        program_dir / "soffice.exe",
        program_dir / "soffice",
        program_dir / "soffice.bin",
    ]
    for c in candidates:
        if c.is_file():
            return c
    return None


def run_smoke_test(soffice_path: Path) -> bool:
    """运行最简 RTF->DOCX 转换冒烟测试。

    Args:
        soffice_path: soffice 可执行文件路径。

    Returns:
        若产物有效则返回 True，否则返回 False。
    """
    if not soffice_path.exists():
        print(f"冒烟测试失败: soffice 可执行文件不存在: {soffice_path}", file=sys.stderr)
        return False

    with tempfile.TemporaryDirectory(prefix="wdt_lo_smoke_") as tmp_str:
        tmp_dir = Path(tmp_str)
        rtf_file = tmp_dir / "smoke.rtf"
        rtf_file.write_bytes(MINIMAL_RTF)
        out_dir = tmp_dir / "out"
        out_dir.mkdir(parents=True, exist_ok=True)
        profile_dir = tmp_dir / "profile"
        profile_dir.mkdir(parents=True, exist_ok=True)

        profile_url = profile_dir.resolve().as_uri()
        cmd = [
            str(soffice_path),
            "--headless",
            "--norestore",
            "--nolockcheck",
            "--nodefault",
            "--nologo",
            "--nofirststartwizard",
            f"-env:UserInstallation={profile_url}",
            "--convert-to",
            "docx",
            "--outdir",
            str(out_dir),
            str(rtf_file),
        ]
        try:
            res = subprocess.run(
                cmd,
                capture_output=True,
                timeout=240,
                text=True,
                check=False,
            )
            out_docx = out_dir / "smoke.docx"
            success = out_docx.is_file() and out_docx.stat().st_size > 0
            if not success:
                print(
                    f"冒烟测试未生成有效 DOCX (退出码: {res.returncode}):\n"
                    f"stdout: {res.stdout}\nstderr: {res.stderr}",
                    file=sys.stderr,
                )
            return success
        except (subprocess.TimeoutExpired, OSError) as exc:
            print(f"冒烟测试进程执行异常: {exc}", file=sys.stderr)
            return False


def rollback_deletions(source: Path, output: Path, plan: list[Path]) -> None:
    """回滚删除操作：从 source 恢复 plan 中的所有文件与目录。"""
    for rel in plan:
        src_item = source / rel
        dst_item = output / rel
        if not src_item.exists() and not src_item.is_symlink():
            continue
        if src_item.is_file() or src_item.is_symlink():
            dst_item.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src_item, dst_item, follow_symlinks=False)
        elif src_item.is_dir():
            dst_item.mkdir(parents=True, exist_ok=True)


def execute(
    source: Path,
    output: Path,
    soffice: Path | None = None,
    smoke: bool = True,
) -> list[Path]:
    """复制 source 到 output，按 prune_plan 执行删除，并进行冒烟验证。

    Args:
        source: 源 LibreOffice 便携目录。
        output: 裁剪输出目标目录。
        soffice: 显式指定的 soffice 路径；若为 None 则自动在 output/program 中查找。
        smoke: 是否执行 RTF->DOCX 冒烟验证。

    Returns:
        实际删除的相对路径清单。

    Raises:
        RuntimeError: 冒烟测试失败且已回滚，或必要文件缺失。
    """
    source = Path(source).resolve()
    output = Path(output).resolve()

    if not source.exists():
        raise FileNotFoundError(f"源目录不存在: {source}")

    plan = prune_plan(source)

    # 1. 复制目录树
    output.mkdir(parents=True, exist_ok=True)
    shutil.copytree(source, output, dirs_exist_ok=True)
    size_before = get_dir_size_bytes(output)

    # 2. 逐项删除（按路径深度从深到浅删除）
    sorted_plan = sorted(plan, key=lambda p: len(p.parts), reverse=True)
    for rel in sorted_plan:
        target = output / rel
        if not target.exists() and not target.is_symlink():
            continue
        if target.is_file() or target.is_symlink():
            target.unlink()
        elif target.is_dir():
            try:
                target.rmdir()
            except OSError:
                shutil.rmtree(target, ignore_errors=True)

    # 3. 冒烟验证
    if smoke:
        target_soffice = soffice if soffice is not None else find_bundled_soffice(output)
        if target_soffice is None or not target_soffice.exists():
            print(
                "警告: 未在输出目录中找到 soffice 可执行文件，且未指定 --soffice；"
                "若在非 Windows/CI 环境下构建，请通过 --no-smoke 明确跳过冒烟验证。",
                file=sys.stderr,
            )
            # 冒烟找不到 soffice 时明确报错并回滚
            rollback_deletions(source, output, plan)
            raise RuntimeError(
                "冒烟验证失败: 找不到可执行的 soffice 文件，已回滚删除操作。"
            )

        print(f"正在执行无头转换冒烟验证 ({target_soffice})...")
        smoke_ok = run_smoke_test(target_soffice)
        if not smoke_ok:
            print("警告: 冒烟验证失败，正在回滚本批删除...", file=sys.stderr)
            rollback_deletions(source, output, plan)
            raise RuntimeError("冒烟验证失败，已回滚删除的文件并终止裁剪。")
        print("冒烟验证通过: RTF->DOCX 转换正常。")

    # 4. 统计并打印体积变化
    size_after = get_dir_size_bytes(output)
    mb_before = size_before / (1024 * 1024)
    mb_after = size_after / (1024 * 1024)
    mb_saved = mb_before - mb_after
    pct_saved = (mb_saved / mb_before * 100) if mb_before > 0 else 0.0

    print(f"裁剪完成: 共移除 {len(plan)} 项。")
    print(
        f"裁剪前体积: {mb_before:.1f} MB -> 裁剪后体积: {mb_after:.1f} MB "
        f"(减少 {mb_saved:.1f} MB, -{pct_saved:.1f}%)"
    )
    return plan


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="将官方 LibreOffice 便携版裁剪为批处理最小运行时"
    )
    parser.add_argument(
        "--source",
        type=Path,
        required=True,
        help="官方 LibreOffice 便携版根目录",
    )
    parser.add_argument(
        "--output",
        type=Path,
        required=True,
        help="裁剪后输出的目标目录 (通常为 runtime/libreoffice)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="仅打印待删除清单，不实际修改或复制文件",
    )
    parser.add_argument(
        "--smoke",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="裁剪后是否执行 RTF->DOCX 冒烟验证 (默认开启，--no-smoke 可跳过)",
    )
    parser.add_argument(
        "--soffice",
        type=Path,
        default=None,
        help="显式指定 soffice 可执行文件路径 (可选)",
    )
    args = parser.parse_args(argv)

    if args.dry_run:
        plan = prune_plan(args.source)
        print(f"[dry-run] 扫描源目录: {args.source}")
        print(f"[dry-run] 共规划删除 {len(plan)} 项:")
        for item in plan:
            print(f"  - {item.as_posix()}")
        return 0

    try:
        execute(
            source=args.source,
            output=args.output,
            soffice=args.soffice,
            smoke=args.smoke,
        )
        return 0
    except Exception as e:
        print(f"执行失败: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
