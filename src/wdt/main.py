"""
Windows Doc Tool (wdt) CLI 入口。

提供环境诊断 (detect)、批量转换 (convert)、目录注入 (toc)、端到端流水线 (run)
以及图形交互界面 (gui) 命令行能力。
"""

from __future__ import annotations

import argparse
import glob
import os
import re
import sys
import time
from pathlib import Path
from typing import Sequence

from wdt.core.batch_runner import _default_jobs, _get_physical_memory_gb
from wdt.core.engine_finder import Engine, EngineNotFoundError, find_engine
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
            f"无效的 --pages 格式: '{pages_str}'，应为 'A-B' 形式的正整数区间（例如 1-50）"
        )
    a = int(m.group(1))
    b = int(m.group(2))
    if a < 1 or b < 1:
        raise ValueError(f"页码必须从 1 开始，不能为 0 或负数: {pages_str}")
    if a > b:
        raise ValueError(f"起始页码不能大于结束页码: {a} > {b}")
    return a, b


def parse_position(pos_str: str) -> tuple[Position, str | None]:
    """解析目录插入位置参数，支持 bookmark:NAME 语法。"""
    pos_str = pos_str.strip()
    if pos_str.startswith("bookmark:"):
        bm_name = pos_str[len("bookmark:"):].strip()
        if not bm_name:
            raise ValueError("位置参数 bookmark:NAME 中的书签名称不能为空")
        return Position.BOOKMARK, bm_name

    mapping = {
        "document-start": Position.DOCUMENT_START,
        "before-first-heading": Position.BEFORE_FIRST_HEADING,
        "replace-placeholder": Position.REPLACE_PLACEHOLDER,
    }
    if pos_str in mapping:
        return mapping[pos_str], None
    raise ValueError(f"不支持的目录插入位置: '{pos_str}'")


def build_toc_config(
    style_str: str,
    depth: int,
    title: str | None,
    position_str: str,
    page_mode_str: str,
    page_offset: int,
    backup: bool,
) -> TocConfig:
    """构建 TOC 注入配置对象。"""
    style_map = {
        "classic": TocStyle.CLASSIC_DOT,
        "modern": TocStyle.MODERN_ACCENT,
        "regulatory": TocStyle.REGULATORY_LISTING,
        "minimal": TocStyle.MINIMAL_CLEAN,
    }
    if style_str not in style_map:
        raise ValueError(f"不支持的 TOC 样式: '{style_str}'")
    style = style_map[style_str]

    page_mode_map = {
        "static-restart": PageMode.STATIC_RESTART,
        "static-offset": PageMode.STATIC_OFFSET,
        "static-in-place": PageMode.STATIC_IN_PLACE,
        "dynamic": PageMode.DYNAMIC,
    }
    if page_mode_str not in page_mode_map:
        raise ValueError(f"不支持的页码模式: '{page_mode_str}'")
    page_mode = page_mode_map[page_mode_str]

    position, bm_name = parse_position(position_str)
    doc_title = title if title is not None else "Table of Contents"

    return TocConfig(
        style=style,
        depth=depth,
        title=doc_title,
        position=position,
        bookmark_name=bm_name,
        page_mode=page_mode,
        page_offset=page_offset,
        backup=backup,
    )


# ---------------------------------------------------------------------------
# 子命令实现
# ---------------------------------------------------------------------------


def cmd_detect(args: argparse.Namespace) -> int:
    """执行 detect 环境诊断命令。"""
    bundled_dir = Path("./runtime/libreoffice")
    engine: Engine | None = None
    try:
        engine = find_engine(bundled_dir=bundled_dir if bundled_dir.is_dir() else None)
    except EngineNotFoundError:
        pass

    sz_cmd = _find_seven_zip()
    cpu_count = os.cpu_count() or 1
    mem_gb = _get_physical_memory_gb()
    default_jobs = _default_jobs()

    print("==================================================")
    print("Windows Doc Tool (wdt) - 环境诊断")
    print("==================================================")

    if engine:
        ver_str = f"{engine.version[0]}.{engine.version[1]}" if engine.version else "未知"
        gray_support = "支持" if (engine.version and engine.version >= (7, 4)) else "不支持 (<7.4)"
        print("[转换引擎]")
        print("  - 状态: 已就绪")
        print(f"  - 路径: {engine.soffice}")
        print(f"  - 来源: {engine.source}")
        print(f"  - 版本: {ver_str}")
        print(f"  - 原生灰度导出 (>=7.4): {gray_support}")
    else:
        print("[转换引擎]")
        print("  - 状态: ❌ 未找到可用的 LibreOffice / soffice 转换引擎！")
        print("[安装提示]")
        print("  - Linux:   sudo apt install libreoffice-writer-nogui")
        print("  - macOS:   brew install --cask libreoffice")
        print("  - Windows: 请安装 LibreOffice 或将 portable 运行时放置于 ./runtime/libreoffice")

    print("[归档与辅助工具]")
    if sz_cmd:
        print(f"  - 7-Zip (7z/7za): 可用 ({sz_cmd})")
    else:
        print("  - 7-Zip (7z/7za): 未找到（大文件原位更新将降级为标准 zipfile）")

    print("[系统硬件资源]")
    print(f"  - CPU 核心数: {cpu_count}")
    print(f"  - 可用物理内存: {mem_gb:.1f} GB")
    print(f"  - 推荐默认并发数: {default_jobs}")
    print("==================================================")

    return 0 if engine is not None else 1


def cmd_convert(args: argparse.Namespace) -> int:
    """执行 convert 批量格式转换命令。"""
    target_path = Path(args.target).resolve()
    if not target_path.exists():
        print(f"错误: 目标路径不存在 - {target_path}", file=sys.stderr)
        return 1

    # 参数语义校验
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

    if target_fmt == "docx" and args.gray:
        print("警告: --gray 仅在导出 PDF 时有效，当前目标格式为 docx，已忽略此选项。", file=sys.stderr)

    # 探测 LibreOffice 引擎
    bundled_dir = Path("./runtime/libreoffice")
    try:
        engine = find_engine(bundled_dir=bundled_dir if bundled_dir.is_dir() else None)
    except EngineNotFoundError:
        print("错误: 未找到可用的 LibreOffice / soffice 转换引擎！", file=sys.stderr)
        print("安装提示：Linux 执行 'sudo apt install libreoffice-writer-nogui'", file=sys.stderr)
        return 1

    # 收集文件
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
        gray=args.gray,
        jobs=args.jobs,
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
        events = run_pipeline(files, options, engine=engine)
        elapsed = time.time() - t0
        failed = [e for e in events if e.status == "failed"]
        if failed:
            print(f"✗ 失败: {src.name} - {failed[0].message}", file=sys.stderr)
            return 1
        print(f"✓ 转换完成 ({elapsed:.2f}s) → {dst_pdf}")
        return 0

    # 批量转换
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

    print(f"\n🚀 开始转换至 {target_fmt.upper()} 格式（待转换: {len(files)} 个文档）...")
    events = run_pipeline(files, options, progress_cb=progress_cb, engine=engine)

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

    target_files = [f.resolve() for f in sorted(set(target_files))]
    if not target_files:
        print("错误: 未指定任何待注入目录的 docx 文件", file=sys.stderr)
        return 1

    if args.page_mode == "static-offset" and args.page_offset == 0:
        print("警告: --page-mode static-offset 建议指定 --page-offset N (当前偏移量为 0)", file=sys.stderr)

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
        jobs=1,
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

    events = run_pipeline(target_files, options, progress_cb=progress_cb)

    success_count = sum(1 for e in events if e.status == "ok")
    fail_count = sum(1 for e in events if e.status == "failed")

    print("\n" + "=" * 50)
    print(f"TOC 注入结束！成功: {success_count} 个, 失败: {fail_count} 个")
    print("=" * 50)

    return 0 if fail_count == 0 else 1


def cmd_run(args: argparse.Namespace) -> int:
    """执行 run 端到端流水线命令。"""
    if args.pdf and args.page_mode == "dynamic":
        print(
            "错误: --pdf 不能与 --page-mode dynamic 同时使用（无头 LibreOffice 不会自动刷新动态目录域）",
            file=sys.stderr,
        )
        return 1

    if args.page_mode == "static-offset" and args.page_offset == 0:
        print("警告: --page-mode static-offset 建议指定 --page-offset N (当前偏移量为 0)", file=sys.stderr)

    target_path = Path(args.target).resolve()
    if not target_path.exists():
        print(f"错误: 目标路径不存在 - {target_path}", file=sys.stderr)
        return 1

    bundled_dir = Path("./runtime/libreoffice")
    try:
        engine = find_engine(bundled_dir=bundled_dir if bundled_dir.is_dir() else None)
    except EngineNotFoundError:
        print("错误: 未找到可用的 LibreOffice / soffice 转换引擎！", file=sys.stderr)
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
        gray=args.gray,
        toc=toc_config,
        jobs=args.jobs,
        force=args.force,
        base_dir=target_path if target_path.is_dir() else None,
    )

    print(f"🔄 启动一体化流水线: 待处理 {len(files)} 个文档 (导出PDF: {args.pdf})...\n")

    failed_items: list[str] = []

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
            failed_items.append(f"{ev.src.name} ({stage_name}失败: {ev.message})")
        elif ev.status == "skipped":
            print(f"  ⏩ [{stage_name}] 跳过: {ev.src.name} - {ev.message}")

    events = run_pipeline(files, options, progress_cb=progress_cb, engine=engine)

    # 统计独立成功的文件
    unique_failed_srcs = {e.src for e in events if e.status == "failed"}
    success_count = len(files) - len(unique_failed_srcs)

    print("\n" + "=" * 50)
    print("🎉 一体化流水线执行完毕！")
    print(f"   ✓ 成功完成: {success_count} 个")
    print(f"   ✗ 处理失败: {len(unique_failed_srcs)} 个")
    if failed_items:
        for fi in failed_items:
            print(f"     - {fi}")
    print("=" * 50)

    return 0 if len(unique_failed_srcs) == 0 else 1


def cmd_gui(args: argparse.Namespace) -> int:
    """启动图形界面 (PySide6)。"""
    try:
        from wdt.ui.main_window import launch_gui
    except ImportError as e:
        print(f"错误: 无法启动 GUI 界面，缺少依赖: {e}", file=sys.stderr)
        print('请执行: pip install ".[gui]"', file=sys.stderr)
        return 1
    return launch_gui()


# ---------------------------------------------------------------------------
# CLI 总调度
# ---------------------------------------------------------------------------


def build_parser() -> WdtArgumentParser:
    """构建顶层 CLI 参数解析器。"""
    parser = WdtArgumentParser(
        prog="wdt",
        description="Windows Doc Tool (wdt) - 高性能办公与临床大文件批处理与目录注入工具",
    )
    subparsers = parser.add_subparsers(dest="subcommand", metavar="COMMAND")

    # 1. detect
    sub_detect = subparsers.add_parser("detect", help="检查运行环境与依赖")
    sub_detect.set_defaults(func=cmd_detect)

    # 2. convert
    sub_conv = subparsers.add_parser("convert", help="批量文档格式转换")
    sub_conv.add_argument("target", help="目标文件或包含文档的目录")
    sub_conv.add_argument("-r", "--recursive", action="store_true", help="递归扫描子目录")
    sub_conv.add_argument("-o", "--outdir", default=None, help="指定输出目录")
    sub_conv.add_argument("-j", "--jobs", type=int, default=None, help="并发工作者数")
    sub_conv.add_argument(
        "--to",
        choices=["docx", "pdf", "both"],
        default="docx",
        help="目标输出格式（默认 docx）",
    )
    sub_conv.add_argument("--gray", action="store_true", help="导出灰度 PDF（需要 LO >= 7.4）")
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
        choices=["static-restart", "static-offset", "static-in-place", "dynamic"],
        default="static-restart",
        help="页码计算模式（默认 static-restart）",
    )
    sub_toc.add_argument("--page-offset", type=int, default=0, help="页码偏移补偿量")
    sub_toc.add_argument("--no-backup", action="store_true", help="就地修改，不保留 .bak 备份文件")
    sub_toc.set_defaults(func=cmd_toc)

    # 4. run
    sub_run = subparsers.add_parser("run", help="端到端一体化流水线 (转换 -> 注入TOC -> 导出PDF)")
    sub_run.add_argument("target", help="目标文件或包含文档的目录")
    sub_run.add_argument("-r", "--recursive", action="store_true", help="递归扫描子目录")
    sub_run.add_argument("-o", "--outdir", default=None, help="指定输出目录")
    sub_run.add_argument("-j", "--jobs", type=int, default=None, help="并发工作者数")
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
        choices=["static-restart", "static-offset", "static-in-place", "dynamic"],
        default="static-restart",
        help="页码计算模式（默认 static-restart）",
    )
    sub_run.add_argument("--page-offset", type=int, default=0, help="页码偏移补偿量")
    sub_run.add_argument("--pdf", action="store_true", help="级联生成 PDF")
    sub_run.add_argument("--gray", action="store_true", help="导出的 PDF 为灰度版本")
    sub_run.add_argument("-f", "--force", action="store_true", help="强制覆盖已存在的目标文件")
    sub_run.set_defaults(func=cmd_run)

    # 5. gui
    sub_gui = subparsers.add_parser("gui", help="启动图形交互界面")
    sub_gui.set_defaults(func=cmd_gui)

    return parser


def main(argv: Sequence[str] | None = None) -> int:
    """CLI 主入口函数，返回整型退出码。"""
    if argv is None:
        argv = sys.argv[1:]

    parser = build_parser()
    if not argv:
        parser.print_help()
        return 0

    try:
        args = parser.parse_args(argv)
        setattr(args, "_raw_argv", list(argv))
    except SystemExit as e:
        code = e.code
        if code is None:
            return 0
        if isinstance(code, int):
            return code
        return 1
    except Exception as e:
        sys.stderr.write(f"参数解析错误: {e}\n")
        return 2

    if not hasattr(args, "func"):
        parser.print_help()
        return 0

    try:
        return args.func(args)
    except SystemExit as e:
        code = e.code
        if code is None:
            return 0
        if isinstance(code, int):
            return code
        return 1
    except Exception as e:
        sys.stderr.write(f"执行异常: {e}\n")
        return 1


if __name__ == "__main__":
    sys.exit(main())
