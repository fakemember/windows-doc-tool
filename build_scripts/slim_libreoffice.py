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
import time
from typing import Callable

# 最简 RTF 测试片段（用于冒烟验证）
MINIMAL_RTF = (
    b"{\\rtf1\\ansi\\deff0"
    b"{\\fonttbl{\\f0\\fnil\\fcharset0 Arial;}}"
    b"\\viewkind4\\uc1\\pard\\lang2052\\fs20 Minimal RTF Smoke Test\\par"
    b"}"
)

SMOKE_REGISTRY_XCU = """\
<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Common/Misc"><prop oor:name="FirstRun" oor:op="fuse"><value>false</value></prop></item>
<item oor:path="/org.openoffice.Office.Common/Misc"><prop oor:name="UseOpenCL" oor:op="fuse"><value>false</value></prop></item>
<item oor:path="/org.openoffice.Office.Update/Update"><prop oor:name="Enabled" oor:op="fuse"><value>false</value></prop></item>
<item oor:path="/org.openoffice.Office.Update/Update"><prop oor:name="UpdateRunning" oor:op="fuse"><value>false</value></prop></item>
<item oor:path="/org.openoffice.Setup/Office"><prop oor:name="ooSetupInstCompleted" oor:op="fuse"><value>true</value></prop></item>
</oor:items>
"""


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


# 显式规则列表：只裁剪 share/ 下安全且占体积的组件，
# 以及自更新器可执行文件、配置与暂存区（保留它们会触发启动自更新检查/补丁应用，
# 导致转换请求被更新器拦截或运行时目录被后台进程锁定，必须彻底裁掉）
UPDATER_PROGRAM_FILES = frozenset(
    {
        "updater.exe",
        "update_service.exe",
        "updater.ini",
    }
)

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
    PruneRule(
        name="updater_staging",
        description="更新器补丁暂存区 (updating/**)",
        pattern="updating/**",
        matcher=lambda rel: len(rel.parts) >= 1 and rel.parts[0] == "updating",
    ),
    PruneRule(
        name="updater_components",
        description="内置在线更新器组件 (update-settings.ini, program/updater*)",
        pattern="update-settings.ini | program/{updater.exe,update_service.exe,updater.ini}",
        matcher=lambda rel: (
            len(rel.parts) == 1 and rel.parts[0].lower() == "update-settings.ini"
        )
        or (
            len(rel.parts) == 2
            and rel.parts[0] == "program"
            and rel.parts[1].lower() in UPDATER_PROGRAM_FILES
        ),
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
        if not parts:
            continue
        # share/program 绝对不删
        if parts[0] == "share" and len(parts) >= 2 and parts[1] == "program":
            continue
        # program/ 下仅允许删除在线更新器组件，核心文件与 DLL 绝对保护
        if parts[0] == "program" and (
            len(parts) != 2 or parts[1].lower() not in UPDATER_PROGRAM_FILES
        ):
            continue
        if parts[0] not in ("share", "updating", "program") and parts[0].lower() != "update-settings.ini":
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


def run_smoke_test(soffice_path: Path, attempts: int = 3) -> bool:
    """运行最简 RTF->DOCX 转换冒烟测试（带重试与共享 profile）。

    Windows 全新环境首次无头启动 LO 时存在初始化竞态：LO 24.8 便携版首次
    启动会执行 profile 更新并自我重启，原 soffice.com 进程退出码 0 但转换
    请求被丢弃（CI 知名问题）。对策是所有尝试共享同一个 profile——第一次
    尝试负责"热身"，后续尝试在已初始化的 profile 上转换即可成功。

    Args:
        soffice_path: soffice 可执行文件路径。
        attempts: 最大尝试次数。

    Returns:
        任一尝试产出有效 DOCX 则返回 True，全部失败返回 False。
    """
    if not soffice_path.exists():
        print(f"冒烟测试失败: soffice 可执行文件不存在: {soffice_path}", file=sys.stderr)
        return False

    # ignore_cleanup_errors: 异步存活的 soffice.bin 可能仍占用 profile 文件，
    # 清理失败不应导致整体报错
    with tempfile.TemporaryDirectory(
        prefix="wdt_lo_smoke_", ignore_cleanup_errors=True
    ) as tmp_str:
        tmp_dir = Path(tmp_str)
        rtf_file = tmp_dir / "smoke.rtf"
        rtf_file.write_bytes(MINIMAL_RTF)
        out_dir = tmp_dir / "out"
        out_dir.mkdir(parents=True, exist_ok=True)
        profile_dir = tmp_dir / "profile"
        user_dir = profile_dir / "user"
        user_dir.mkdir(parents=True, exist_ok=True)
        (user_dir / "registrymodifications.xcu").write_text(
            SMOKE_REGISTRY_XCU, encoding="utf-8"
        )
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

        for attempt in range(1, attempts + 1):
            try:
                res = subprocess.run(
                    cmd,
                    capture_output=True,
                    timeout=240,
                    text=True,
                    check=False,
                )
            except subprocess.TimeoutExpired:
                print(f"冒烟测试第 {attempt}/{attempts} 次超时 (>240s)", file=sys.stderr)
                continue
            except OSError as exc:
                print(f"冒烟测试第 {attempt}/{attempts} 次进程执行异常: {exc}", file=sys.stderr)
                continue

            out_docx = out_dir / "smoke.docx"
            if out_docx.is_file() and out_docx.stat().st_size > 0:
                if attempt > 1:
                    print(f"冒烟测试于第 {attempt}/{attempts} 次尝试通过。")
                return True
            print(
                f"冒烟测试第 {attempt}/{attempts} 次未生成有效 DOCX (退出码: {res.returncode}):\n"
                f"stdout: {res.stdout}\nstderr: {res.stderr}",
                file=sys.stderr,
            )
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


def _kill_soffice_under(base_dir: Path) -> None:
    """终止可执行文件位于 base_dir 之下的 soffice/soffice.bin 进程。

    LO 首次启动的更新器可能在冒烟结束后仍以后台进程存活并占用/锁定目录，
    必须在复制或回滚前清场。只杀 base_dir 之下的实例，不影响用户其他 LO 会话。
    """
    base_str = str(base_dir)
    if sys.platform == "win32":
        ps_script = (
            "Get-CimInstance Win32_Process "
            "-Filter \"Name='soffice.exe' OR Name='soffice.bin'\" | "
            f"Where-Object {{ $_.ExecutablePath -like '{base_str}\\*' }} | "
            "ForEach-Object { Stop-Process -Id $_.ProcessId -Force "
            "-ErrorAction SilentlyContinue }"
        )
        for _ in range(2):  # 更新器可能复活，杀两轮
            subprocess.run(
                ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps_script],
                check=False,
                capture_output=True,
            )
            time.sleep(1)
    else:
        subprocess.run(["pkill", "-f", base_str], check=False, capture_output=True)


def execute(
    source: Path,
    output: Path,
    soffice: Path | None = None,
    smoke: bool = True,
) -> list[Path]:
    """复制 source 到 output，按 prune_plan 执行删除，并对 output 副本冒烟验证。

    注意：绝不对 source 原地冒烟——LO 首次启动的更新器会锁住并修改源目录。
    复制带重试（刚解压的 soffice.exe 可能被杀软扫描占用）；冒烟结束后无论
    成败都会清理 output 之下残留的 soffice 进程，避免后续打包步骤撞文件锁。

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

    # 1. 复制目录树（带重试：刚解压的 soffice.exe 可能被杀软扫描占用）
    output.mkdir(parents=True, exist_ok=True)
    for copy_attempt in range(1, 4):
        try:
            shutil.copytree(source, output, dirs_exist_ok=True)
            break
        except PermissionError as exc:
            if copy_attempt == 3:
                raise
            print(
                f"复制被占用 ({exc})，2 秒后重试 ({copy_attempt}/3)...",
                file=sys.stderr,
            )
            time.sleep(2)
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

    # 3. 冒烟验证（仅针对 output 副本，绝不对 source 原地冒烟）
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
        # 清理冒烟唤起的后台 LO/更新器进程，防止占用文件导致回滚或后续打包失败
        _kill_soffice_under(output)
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
