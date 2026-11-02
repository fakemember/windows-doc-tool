"""
Windows Doc Tool (wdt) CLI 入口。

提供环境诊断 (detect)、批量转换 (convert)、目录注入 (toc)、端到端流水线 (run)
以及图形交互界面 (gui) 命令行能力。
纯 Microsoft Word 原生 COM 引擎驱动。
"""

from __future__ import annotations

import argparse
import glob
import multiprocessing
import os
import re
import sys
import time
from pathlib import Path
from typing import Sequence

from wdt.core.engine_finder import (
    EngineNotFoundError,
    WordEngine,
    find_word_engine,
)
from wdt.pipeline import (
    PipelineEvent,
    PipelineOptions,
    Stage,
    resolve_destination,
    run_pipeline,
)
from wdt.toc.docx_stream_patcher import (
    Position,
    TocConfig,
)
from wdt.toc.page_calculator import PageMode
from wdt.toc.styles_templates import TocStyle
from wdt.toc.zip_rewriter import _find_seven_zip

SUPPORTED_EXTENSIONS = (".doc", ".rtf", ".docx")


class WdtArgumentParser(argparse.ArgumentParser):
    """自定义参数解析器，避免参数错误抛出未捕获异常。"""

    def error(self, message: str) -> None:
        self.print_usage(sys.stderr)
        sys.stderr.write(f"错误: {message}\n")
        sys.exit(2)


def is_temporary_file(p: Path) -> bool:
    """判断是否为 Office 或操作系统临时文件。"""
    return p.name.startswith("~$") or p.name.startswith("._")


def collect_target_files(
    target: Path,
    recursive: bool,
    allowed_exts: tuple[str, ...] = SUPPORTED_EXTENSIONS,
) -> list[Path]:
    """收集待处理文件列表，支持单文件与目录递归。"""
    target = Path(target).resolve()
    if not target.exists():
        raise FileNotFoundError(f"目标路径不存在: {target}")

    if target.is_file():
        if target.suffix.lower() not in allowed_exts:
            raise ValueError(
                f"不支持的文件格式: '{target.suffix}'，支持的扩展名: {', '.join(allowed_exts)}"
            )
        return [target]

    if target.is_dir():
        files: list[Path] = []
        for ext in allowed_exts:
            pattern = f"*{ext}" if not recursive else f"**/*{ext}"
            files.extend(target.rglob(f"*{ext}") if recursive else target.glob(f"*{ext}"))
        return [
            f.resolve()
            for f in sorted(set(files))
            if not is_temporary_file(f)
        ]

    raise FileNotFoundError(f"目标路径不是合法文件或目录: {target}")


def parse_pages_arg(pages_str: str) -> tuple[int, int]:
    """解析 --pages A-B 参数为 1-based 起止页码。"""
    m = re.match(r"^(\d+)-(\d+)$", pages_str.strip())
    if not m:
        raise ValueError(
            f"--pages 格式错误: '{pages_str}'，应形如 '1-5' 或 '1-1'"
        )
    start = int(m.group(1))
    end = int(m.group(2))
    if start <= 0 or end <= 0:
        raise ValueError("--pages 页码必须为正整数 (>=1)")
    if start > end:
        raise ValueError(f"--pages 起始页码 ({start}) 不能大于结束页码 ({end})")
    return start, end


def build_toc_config(
    style_str: str,
    depth: int,
    title: str | None,
    position_str: str,
    page_mode_str: str,
    page_offset: int,
    backup: bool,
) -> TocConfig:
    """解析参数构造 TocConfig 对象。"""
    style_map = {
        "classic": TocStyle.CLASSIC_DOT,
        "modern": TocStyle.MODERN_ACCENT,
        "regulatory": TocStyle.REGULATORY_LISTING,
        "minimal": TocStyle.MINIMAL_CLEAN,
    }
    style = style_map.get(style_str, TocStyle.CLASSIC_DOT)

    pos_clean = position_str.strip().lower()
    bookmark_name: str | None = None
    if pos_clean == "document-start":
        position = Position.DOCUMENT_START
    elif pos_clean == "before-first-heading":
        position = Position.BEFORE_FIRST_HEADING
    elif pos_clean == "replace-placeholder":
        position = Position.REPLACE_PLACEHOLDER
    elif pos_clean.startswith("bookmark:"):
        position = Position.BOOKMARK
        bookmark_name = position_str[len("bookmark:"):].strip()
        if not bookmark_name:
            raise ValueError("--position bookmark:NAME 必须指定非空书签名称")
    else:
        raise ValueError(
            f"未知的 --position: '{position_str}'，支持 document-start, before-first-heading, replace-placeholder, bookmark:NAME"
        )

    mode_map = {
        "static-restart": PageMode.STATIC_RESTART,
        "static-offset": PageMode.STATIC_OFFSET,
        "static-in-place": PageMode.STATIC_IN_PLACE,
        "measured": PageMode.STATIC_MEASURED,
        "dynamic": PageMode.DYNAMIC,
    }
    page_mode = mode_map.get(page_mode_str, PageMode.STATIC_RESTART)

    if title is None:
        title_text = "目  录" if style == TocStyle.CLASSIC_DOT else "Table of Contents"
    else:
        title_text = title

    return TocConfig(
        style=style,
        depth=depth,
        title=title_text,
        position=position,
        bookmark_name=bookmark_name,
        page_mode=page_mode,
        page_offset=page_offset,
        backup=backup,
    )


def _attach_windows_console() -> bool:
    """在 Windows 子系统为 windows 的可执行程序中尝试附着父进程控制台。"""
    if sys.platform != "win32":
        return False
    try:
        import ctypes

        kernel32 = ctypes.windll.kernel32
        attached = kernel32.AttachConsole(0xFFFFFFFF) != 0
        if not attached:
            return False

        kernel32.SetConsoleOutputCP(65001)
        kernel32.SetConsoleCP(65001)

        for fd, fileno, mode, name in (
            (1, 1, "w", "stdout"),
            (2, 2, "w", "stderr"),
        ):
            try:
                handle = kernel32.GetStdHandle(-11 if fd == 1 else -12)
                if handle and handle != -1:
                    os_fd = ctypes.cdll.msvcrt._open_osfhandle(handle, 0)
                    if os_fd >= 0:
                        f = open(os_fd, mode, encoding="utf-8", errors="replace", buffering=1)
                        setattr(sys, name, f)
            except Exception:
                pass
        return True
    except Exception:
        return False


# ---------------------------------------------------------------------------
# 子命令实现
# ---------------------------------------------------------------------------

def cmd_detect(args: argparse.Namespace) -> int:
    """执行 detect 环境诊断命令。"""
    word_engine: WordEngine = find_word_engine()
    sz_cmd = _find_seven_zip()
    cpu_count = os.cpu_count() or 1

    print("==================================================")
    print("Windows Doc Tool (wdt) - 环境诊断")
    print("==================================================")

    print("[转换引擎 - Microsoft Word COM]")
    if word_engine.available:
        print("  - 状态: 已就绪")
        print(f"  - 版本: {word_engine.version or '未知'}")
        print(f"  - 路径: {word_engine.app_path or '未直接解析出路径 (通过 COM 调度)'}")
    else:
        print(f"  - 状态: ❌ 不可用 ({word_engine.error or '未安装或无法通过 COM 自动化'})")
        if word_engine.version:
            print(f"  - 检测到的版本: {word_engine.version}")
        if word_engine.app_path:
            print(f"  - 路径: {word_engine.app_path}")

    print("[归档与辅助工具]")
    if sz_cmd:
        print(f"  - 7-Zip (7z/7za): 可用 ({sz_cmd})")
    else:
        print("  - 7-Zip (7z/7za): 未找到（大文件原位更新将使用标准 zipfile）")

    print("[系统硬件资源]")
    print(f"  - CPU 核心数: {cpu_count}")
    print("==================================================")

    return 0 if word_engine.available else 1


def cmd_convert(args: argparse.Namespace) -> int:
    """执行 convert 批量格式转换命令。"""
    target_path = Path(args.target).resolve()
    if not target_path.exists():
        print(f"错误: 目标路径不存在 - {target_path}", file=sys.stderr)
        return 1

    pages_range: tuple[int, int] | None = None
    if args.pages:
        if target_path.is_dir():
            print("错误: --pages 选项仅支持单文件转换，不支持目录模式", file=sys.stderr)
            return 1
        try:
            pages_range = parse_pages_arg(args.pages)
        except ValueError as e:
            print(f"错误: {e}", file=sys.stderr)
            return 1

    target_fmt = args.to
    if pages_range is not None:
        if args.to == "docx":
            raw_argv = getattr(args, "_raw_argv", [])
            if "--to" in raw_argv:
                print("错误: --pages 选项仅在导出 PDF 时有效，不能与 --to docx 组合使用", file=sys.stderr)
                return 1
            target_fmt = "pdf"
        elif args.to == "both":
            print("错误: --pages 选项仅支持 PDF 导出，不能与 --to both 组合使用", file=sys.stderr)
            return 1

    try:
        files = collect_target_files(target_path, args.recursive, allowed_exts=SUPPORTED_EXTENSIONS)
    except Exception as e:
        print(f"错误: {e}", file=sys.stderr)
        return 1

    if not files:
        print(f"在 {target_path} 下未找到任何支持的文档。")
        return 0

    options = PipelineOptions(
        mode="convert",
        out_dir=Path(args.outdir) if args.outdir else None,
        make_docx=(target_fmt in ("docx", "both")),
        make_pdf=(target_fmt in ("pdf", "both")),
        force=args.force,
        pages=pages_range,
        base_dir=target_path if target_path.is_dir() else None,
    )

    if pages_range is not None:
        src = files[0]
        p_start, p_end = pages_range
        p_suffix = f"_p{p_start}" if p_start == p_end else f"_p{p_start}-{p_end}"
        out_base = Path(args.outdir).resolve() if args.outdir else src.parent
        dst_pdf = out_base / f"{src.stem}{p_suffix}.pdf"
        if dst_pdf.exists() and not args.force:
            print(f"[跳过] 目标文件已存在: {dst_pdf}")
            return 0
        print(f"📄 提取页码转换: {src.name} (页码: {p_start}-{p_end}) → {dst_pdf.name}")
        t0 = time.time()
        try:
            events = run_pipeline(files, options)
        except (EngineNotFoundError, ValueError) as e:
            print(f"错误: {e}", file=sys.stderr)
            return 1
        elapsed = time.time() - t0
        failed = [e for e in events if e.status == "failed"]
        if failed:
            print(f"✗ 失败: {src.name} - {failed[0].message}", file=sys.stderr)
            return 1
        print(f"✓ 转换完成 ({elapsed:.2f}s) → {dst_pdf}")
        return 0

    digits = len(str(len(files)))
    done_counter = 0

    def progress_cb(ev: PipelineEvent) -> None:
        nonlocal done_counter
        done_counter += 1
        prefix = f"[{done_counter:>{digits}}/{len(files)}]"
        if ev.status == "ok":
            print(f"  {prefix} ✓ 成功: {ev.src.name} ({ev.message})")
        elif ev.status == "failed":
            print(f"  {prefix} ✗ 失败: {ev.src.name} - {ev.message}")

    print(f"\n🚀 开始转换至 {target_fmt.upper()} 格式（Word COM 引擎，待转换: {len(files)} 个文档）...")
    try:
        events = run_pipeline(files, options, progress_cb=progress_cb)
    except (EngineNotFoundError, ValueError) as e:
        print(f"错误: {e}", file=sys.stderr)
        return 1

    total_success = sum(1 for e in events if e.status == "ok")
    total_failed = sum(1 for e in events if e.status == "failed")
    total_skipped = sum(1 for e in events if e.status == "skipped")

    print("\n" + "=" * 50)
    print("🎉 转换流程结束！")
    print(f"   ✓ 成功: {total_success} 个")
    print(f"   ✗ 失败: {total_failed} 个")
    if total_skipped:
        print(f"   ⏩ 跳过: {total_skipped} 个")
    print("=" * 50)

    return 0 if total_failed == 0 else 1


def cmd_toc(args: argparse.Namespace) -> int:
    """执行 toc DOCX 目录注入命令。"""
    target_files: list[Path] = []
    for pattern in args.files:
        matches = [Path(p) for p in glob.glob(pattern)]
        if matches:
            target_files.extend(matches)
        else:
            target_files.append(Path(pattern))

    target_files = [f.resolve() for f in target_files if f.suffix.lower() == ".docx"]
    if not target_files:
        print("错误: 未指定任何合法的 docx 文件", file=sys.stderr)
        return 1

    try:
        config = build_toc_config(
            style_str=args.style,
            depth=args.depth,
            title=args.title,
            position_str=args.position,
            page_mode_str=args.page_mode,
            page_offset=args.page_offset,
            backup=not args.no_backup,
        )
    except ValueError as e:
        print(f"错误: {e}", file=sys.stderr)
        return 1

    options = PipelineOptions(
        mode="toc",
        toc=config,
    )

    print(f"📑 开始向 {len(target_files)} 个文档注入 TOC 目录 (样式: {args.style}, 深度: {args.depth})...\n")

    digits = len(str(len(target_files)))
    counter = 0

    def progress_cb(ev: PipelineEvent) -> None:
        nonlocal counter
        counter += 1
        prefix = f"[{counter:>{digits}}/{len(target_files)}]"
        if ev.status == "ok":
            print(f"  {prefix} ✓ 成功: {ev.src.name} ({ev.message})")
        elif ev.status == "failed":
            print(f"  {prefix} ✗ 失败: {ev.src.name} - {ev.message}")

    try:
        events = run_pipeline(target_files, options, progress_cb=progress_cb)
    except (EngineNotFoundError, ValueError) as e:
        print(f"错误: {e}", file=sys.stderr)
        return 1

    success_count = sum(1 for e in events if e.status == "ok")
    fail_count = sum(1 for e in events if e.status == "failed")

    print("\n" + "=" * 50)
    print(f"TOC 注入结束！成功: {success_count} 个, 失败: {fail_count} 个")
    print("=" * 50)

    return 0 if fail_count == 0 else 1


def cmd_run(args: argparse.Namespace) -> int:
    """执行 run 端到端流水线命令。"""
    target_path = Path(args.target).resolve()
    if not target_path.exists():
        print(f"错误: 目标路径不存在 - {target_path}", file=sys.stderr)
        return 1

    try:
        files = collect_target_files(target_path, args.recursive, allowed_exts=SUPPORTED_EXTENSIONS)
    except Exception as e:
        print(f"错误: {e}", file=sys.stderr)
        return 1

    if not files:
        print(f"在 {target_path} 下未找到任何支持的文档。")
        return 0

    try:
        toc_config = build_toc_config(
            style_str=args.style,
            depth=args.depth,
            title=args.title,
            position_str=args.position,
            page_mode_str=args.page_mode,
            page_offset=args.page_offset,
            backup=True,
        )
    except ValueError as e:
        print(f"错误: {e}", file=sys.stderr)
        return 1

    options = PipelineOptions(
        mode="run",
        out_dir=Path(args.outdir) if args.outdir else None,
        make_docx=True,
        make_pdf=args.pdf,
        toc=toc_config,
        force=args.force,
        base_dir=target_path if target_path.is_dir() else None,
    )

    print(f"🔄 启动一体化流水线: 待处理 {len(files)} 个文档 (导出PDF: {args.pdf}, Word COM 引擎)...\n")

    def progress_cb(ev: PipelineEvent) -> None:
        stage_name = {
            Stage.CONVERT: "DOCX准备",
            Stage.TOC: "TOC注入",
            Stage.PDF: "PDF导出",
        }.get(ev.stage, ev.stage.value)

        if ev.status == "ok":
            print(f"  ✓ [{stage_name}] 成功: {ev.src.name} ({ev.message})")
        elif ev.status == "failed":
            print(f"  ✗ [{stage_name}] 失败: {ev.src.name} - {ev.message}")
        elif ev.status == "skipped":
            print(f"  ⏩ [{stage_name}] 跳过: {ev.src.name} - {ev.message}")

    try:
        events = run_pipeline(files, options, progress_cb=progress_cb)
    except (EngineNotFoundError, ValueError) as e:
        print(f"错误: {e}", file=sys.stderr)
        return 1

    unique_failed_srcs = {e.src for e in events if e.status == "failed"}
    success_count = len(files) - len(unique_failed_srcs)

    print("\n" + "=" * 50)
    print("🎉 一体化流水线执行完毕！")
    print(f"   ✓ 成功: {success_count} 个")
    print(f"   ✗ 失败: {len(unique_failed_srcs)} 个")
    print("=" * 50)

    return 0 if len(unique_failed_srcs) == 0 else 1


def cmd_gui(args: argparse.Namespace) -> int:
    """执行 gui 启动桌面图形交互界面命令。"""
    try:
        from wdt.ui.main_window import launch_gui
        return launch_gui()
    except ImportError as e:
        print(f"错误: 未安装图形交互依赖 (PySide6): {e}", file=sys.stderr)
        print("提示: 请使用 'uv pip install .[gui]' 安装 GUI 依赖后再启动。", file=sys.stderr)
        return 1


def build_parser() -> WdtArgumentParser:
    """构建顶层 CLI 参数解析器。"""
    parser = WdtArgumentParser(
        prog="wdt",
        description="Windows Doc Tool (wdt) - Windows 原生 Word COM 文档批处理与动态目录注入工具",
    )
    subparsers = parser.add_subparsers(dest="subcommand", metavar="COMMAND")

    # 1. detect
    sub_detect = subparsers.add_parser("detect", help="检查运行环境与 Word COM 引擎")
    sub_detect.set_defaults(func=cmd_detect)

    # 2. convert
    sub_conv = subparsers.add_parser("convert", help="批量文档格式转换")
    sub_conv.add_argument("target", help="目标文件或包含文档的目录")
    sub_conv.add_argument("-r", "--recursive", action="store_true", help="递归扫描子目录")
    sub_conv.add_argument("-o", "--outdir", default=None, help="指定输出目录")
    sub_conv.add_argument(
        "--to",
        choices=["docx", "pdf", "both"],
        default="docx",
        help="目标输出格式（默认 docx）",
    )
    sub_conv.add_argument("--pages", default=None, help="仅转换指定页码区间 A-B（仅限单文件）")
    sub_conv.add_argument("-f", "--force", action="store_true", help="强制覆盖已存在的目标文件")
    sub_conv.set_defaults(func=cmd_convert)

    # 3. toc
    sub_toc = subparsers.add_parser("toc", help="向 DOCX 文档注入格式化 TOC 目录")
    sub_toc.add_argument("files", nargs="+", help="一个或多个待注入的 docx 文件")
    sub_toc.add_argument(
        "--style",
        choices=["classic", "modern", "regulatory", "minimal"],
        default="classic",
        help="目录视觉样式（默认 classic）",
    )
    sub_toc.add_argument(
        "--depth",
        type=int,
        choices=[1, 2, 3],
        default=2,
        help="提取标题层级深度（默认 2）",
    )
    sub_toc.add_argument("--title", default=None, help="自定义目录标题文字")
    sub_toc.add_argument(
        "--position",
        default="document-start",
        help="目录插入位置: document-start|before-first-heading|replace-placeholder|bookmark:NAME",
    )
    sub_toc.add_argument(
        "--page-mode",
        choices=["static-restart", "static-offset", "static-in-place", "measured", "dynamic"],
        default="dynamic",
        help="页码计算模式（默认 dynamic 原生动态域）",
    )
    sub_toc.add_argument("--page-offset", type=int, default=0, help="页码偏移补偿量")
    sub_toc.add_argument("--no-backup", action="store_true", help="就地修改，不保留 .bak 备份文件")
    sub_toc.set_defaults(func=cmd_toc)

    # 4. run
    sub_run = subparsers.add_parser("run", help="端到端一体化流水线 (转换 -> 注入TOC -> 导出PDF)")
    sub_run.add_argument("target", help="目标文件或包含文档的目录")
    sub_run.add_argument("-r", "--recursive", action="store_true", help="递归扫描子目录")
    sub_run.add_argument("-o", "--outdir", default=None, help="指定输出目录")
    sub_run.add_argument(
        "--style",
        choices=["classic", "modern", "regulatory", "minimal"],
        default="classic",
        help="目录视觉样式（默认 classic）",
    )
    sub_run.add_argument(
        "--depth",
        type=int,
        choices=[1, 2, 3],
        default=2,
        help="提取标题层级深度（默认 2）",
    )
    sub_run.add_argument("--title", default=None, help="自定义目录标题文字")
    sub_run.add_argument(
        "--position",
        default="document-start",
        help="目录插入位置: document-start|before-first-heading|replace-placeholder|bookmark:NAME",
    )
    sub_run.add_argument(
        "--page-mode",
        choices=["dynamic", "measured", "static-restart", "static-offset", "static-in-place"],
        default="dynamic",
        help="页码计算模式（默认 dynamic 原生动态域）",
    )
    sub_run.add_argument("--page-offset", type=int, default=0, help="页码偏移补偿量")
    sub_run.add_argument("--pdf", action="store_true", help="级联生成 PDF")
    sub_run.add_argument("-f", "--force", action="store_true", help="强制覆盖已存在的目标文件")
    sub_run.set_defaults(func=cmd_run)

    # 5. gui
    sub_gui = subparsers.add_parser("gui", help="启动图形交互界面")
    sub_gui.set_defaults(func=cmd_gui)

    return parser


def main(argv: Sequence[str] | None = None) -> int:
    """CLI 主入口函数，返回整型退出码。"""
    multiprocessing.freeze_support()

    if argv is None:
        argv = sys.argv[1:]

    has_console = _attach_windows_console()
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if callable(reconfigure):
            try:
                reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass

    # Windows 环境且无终端控制台（双击 exe 启动）且未传参数时，默认启动 GUI
    if not argv and sys.platform == "win32" and not has_console and not os.environ.get("PYTEST_CURRENT_TEST"):
        return cmd_gui()

    parser = build_parser()
    setattr(parser, "_raw_argv", argv)
    if not argv:
        parser.print_help()
        return 0

    args = parser.parse_args(argv)
    setattr(args, "_raw_argv", argv)

    func = getattr(args, "func", None)
    if func:
        return func(args)

    parser.print_help()
    return 0


if __name__ == "__main__":
    sys.exit(main())
