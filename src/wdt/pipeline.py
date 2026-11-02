"""
Windows Doc Tool (wdt) 业务流水线模块。

提供脱离 CLI 与 GUI 交互的统一核心编排 API：
- 批量转换 (convert)
- 目录注入 (toc)
- 级联流水线 (run: 转换 -> TOC注入 -> 导出PDF)
"""

from __future__ import annotations

import logging
import shutil
import tempfile
import time
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Callable, Literal, Protocol, Sequence

from pypdf import PdfReader, PdfWriter

from wdt.core.batch_runner import BatchRunner
from wdt.core.converter_worker import FileResult
from wdt.core.docx_normalizer import normalize_docx
from wdt.core.engine_finder import (
    Engine,
    EngineNotFoundError,
    arbitrate_engine,
    find_engine,
)
from wdt.core.format_sniffer import sniff
from wdt.core.wincom_worker import WincomRunner, is_wincom_supported
from wdt.toc.docx_stream_patcher import (
    NoHeadingError,
    TocConfig,
    TocReport,
    inject_toc,
)
from wdt.toc.page_calculator import PageMode
from wdt.toc.page_measurer import (
    build_measurement_copy,
    measure_pages_from_pdf,
    patch_docx_measured_pages,
)
from wdt.toc.zip_rewriter import _find_seven_zip

logger = logging.getLogger(__name__)


EngineKind = Literal["libreoffice", "wincom"]

_ENGINE_LABELS: dict[EngineKind, str] = {
    "libreoffice": "LibreOffice",
    "wincom": "Word COM",
}


class ConvertRunner(Protocol):
    """BatchRunner / WincomRunner 统一调度契约。"""

    def convert(
        self,
        pairs: Sequence[tuple[Path, Path]],
        target: Literal["docx", "pdf"],
        gray: bool = False,
        progress_cb: Callable[[int, int, FileResult], None] | None = None,
        should_cancel: Callable[[], bool] | None = None,
    ) -> list[FileResult]: ...


def _arbitrate_make_pdf(options: "PipelineOptions") -> bool:
    """仲裁视角下"任务是否含 PDF 导出"：toc 模式的 measured 实测即一次 PDF 导出。"""
    if options.mode in ("convert", "run"):
        return options.make_pdf
    return bool(options.toc and options.toc.page_mode == PageMode.STATIC_MEASURED)


def _resolve_engine_kinds(
    files: Sequence[Path],
    options: "PipelineOptions",
    word_available: bool,
) -> dict[Path, EngineKind]:
    """按 engine_mode 为每个源文件决议转换引擎（auto 走量化仲裁 + 规则 C 批次一致性）。"""
    if options.engine_mode == "libreoffice":
        return {src: "libreoffice" for src in files}
    if options.engine_mode == "wincom":
        if not is_wincom_supported():
            raise ValueError(
                "已指定 wincom 引擎，但当前环境不支持 Word COM"
                "（需 Windows 且安装 Microsoft Word 与 pywin32）。"
            )
        return {src: "wincom" for src in files}

    make_pdf = _arbitrate_make_pdf(options)
    kinds: dict[Path, EngineKind] = {}
    for src in files:
        try:
            size = src.stat().st_size
        except OSError:
            size = 0
        kinds[src] = arbitrate_engine(
            file_bytes=size,
            sniffed=sniff(src),
            make_pdf=make_pdf,
            make_docx=options.make_docx,
            word_available=word_available,
        )
    # 规则 C：任务含 PDF 导出时，批内任一文件仲裁为 wincom 则整批统一升级，
    # 保证同批次交付物（如 1 Table + 1 Listing）的排版与渲染引擎一致。
    if make_pdf and any(k == "wincom" for k in kinds.values()):
        kinds = {src: "wincom" for src in kinds}
    return kinds


class Stage(Enum):
    CONVERT = "convert"
    TOC = "toc"
    PDF = "pdf"


@dataclass
class PipelineEvent:
    stage: Stage
    src: Path
    status: str  # "ok" | "failed" | "skipped" | "cancelled"
    message: str = ""


@dataclass
class PipelineOptions:
    mode: Literal["convert", "toc", "run"]
    out_dir: Path | None = None  # None = 源文件同目录
    make_docx: bool = True  # convert/run
    make_pdf: bool = False
    gray: bool = False
    toc: TocConfig | None = None  # toc/run 模式必填
    jobs: int | None = None
    force: bool = False
    pages: tuple[int, int] | None = None  # 仅单文件 PDF 截取
    base_dir: Path | None = None  # 递归源基准目录，用于镜像相对路径
    engine_mode: Literal["auto", "libreoffice", "wincom"] = "auto"  # 双引擎仲裁模式


def slice_pdf(src_pdf: Path, dst_pdf: Path, page_start: int, page_end: int) -> None:
    """使用 pypdf 提取指定 1-based 页码区间的页面并写入目标 PDF。"""
    reader = PdfReader(str(src_pdf))
    total_pages = len(reader.pages)
    if page_start > total_pages:
        raise ValueError(f"起始页码 {page_start} 超出文档总页数 {total_pages}")

    writer = PdfWriter()
    actual_end = min(page_end, total_pages)
    for idx in range(page_start - 1, actual_end):
        writer.add_page(reader.pages[idx])

    dst_pdf.parent.mkdir(parents=True, exist_ok=True)
    with open(dst_pdf, "wb") as f:
        writer.write(f)


def resolve_destination(
    src: Path,
    target_ext: str,
    out_dir: Path | None,
    base_dir: Path | None = None,
) -> Path:
    """计算目标输出路径，保持相对目录镜像结构。"""
    if out_dir is not None:
        out_base = Path(out_dir).resolve()
        if base_dir is not None and base_dir.is_dir():
            try:
                rel = src.resolve().relative_to(base_dir.resolve())
                return out_base / rel.parent / f"{src.stem}{target_ext}"
            except ValueError:
                pass
        return out_base / f"{src.stem}{target_ext}"
    return src.parent / f"{src.stem}{target_ext}"


def _safe_normalize_converted(docx_file: Path, seven_zip_path: Path | None) -> None:
    """规范化 LibreOffice 转换产物，失败不阻断流水线。"""
    try:
        normalize_docx(docx_file, seven_zip=seven_zip_path)
    except Exception as e:
        logger.warning("DOCX 规范化失败 (%s): %s", docx_file.name, e)


def _perform_page_measurement(
    docx_file: Path,
    report: TocReport,
    config: TocConfig,
    runner: ConvertRunner,
    seven_zip_path: Path | None = None,
) -> tuple[bool, list[int], str]:
    """执行 measured 模式的两遍测量与回写。
    
    返回 (is_measured_success, final_pages, message)
    """
    with tempfile.TemporaryDirectory(prefix="wdt_measure_") as tmp_dir:
        tmp_docx = Path(tmp_dir) / f"{docx_file.stem}_measure.docx"
        tmp_pdf = Path(tmp_dir) / f"{docx_file.stem}_measure.pdf"
        try:
            build_measurement_copy(docx_file, tmp_docx)
        except Exception as e:
            logger.warning("测量副本生成失败: %s，回退估算页码", e)
            patch_docx_measured_pages(
                docx_file,
                measured_pages=report.pages,
                bookmark_names=report.bookmark_names,
                sentinel_id=report.sentinel,
                seven_zip=seven_zip_path,
            )
            return False, report.pages, f"测量副本生成异常({e})，已回退估算页码"

        results = runner.convert([(tmp_docx, tmp_pdf)], target="pdf")
        if not results or not results[0].ok or not tmp_pdf.is_file():
            err = results[0].error if results else "临时PDF转换失败"
            logger.warning("实测临时 PDF 导出失败: %s，回退估算页码", err)
            patch_docx_measured_pages(
                docx_file,
                measured_pages=report.pages,
                bookmark_names=report.bookmark_names,
                sentinel_id=report.sentinel,
                seven_zip=seven_zip_path,
            )
            return False, report.pages, f"临时PDF转换异常({err})，已回退估算页码"

        measured = measure_pages_from_pdf(
            tmp_pdf,
            heading_count=report.entries,
            sentinel_id=report.sentinel,
            page_offset=config.page_offset,
        )
        if measured is not None:
            patch_docx_measured_pages(
                docx_file,
                measured_pages=measured,
                bookmark_names=report.bookmark_names,
                sentinel_id=report.sentinel,
                seven_zip=seven_zip_path,
            )
            return True, measured, "实测页码成功"
        else:
            logger.warning("实测未找齐标题锚点，回退估算页码")
            patch_docx_measured_pages(
                docx_file,
                measured_pages=report.pages,
                bookmark_names=report.bookmark_names,
                sentinel_id=report.sentinel,
                seven_zip=seven_zip_path,
            )
            return False, report.pages, "未找齐标题锚点，已回退估算页码"


def run_pipeline(
    files: Sequence[Path],
    options: PipelineOptions,
    progress_cb: Callable[[PipelineEvent], None] | None = None,
    should_cancel: Callable[[], bool] | None = None,
    engine: Engine | None = None,
) -> list[PipelineEvent]:
    """执行文档处理流水线。

    根据 options.mode 执行相应任务，支持进度上报与协同式取消机制。
    """
    # 1. 约束校验
    if options.mode in ("toc", "run") and options.toc is None:
        raise ValueError(f"{options.mode} 模式必须指定 options.toc 配置")

    if (
        options.mode == "run"
        and options.make_pdf
        and options.toc is not None
        and options.toc.page_mode == PageMode.DYNAMIC
    ):
        raise ValueError(
            "dynamic + make_pdf 组合不支持：无头 LibreOffice 不会自动刷新动态目录域"
        )

    if options.mode == "convert" and not options.make_docx and not options.make_pdf:
        raise ValueError("convert 模式必须至少启用 make_docx 或 make_pdf 之一")

    # 2. 双引擎决议与准备
    needs_engine = options.mode in ("convert", "run") or (
        options.mode == "toc"
        and options.toc is not None
        and options.toc.page_mode == PageMode.STATIC_MEASURED
    )
    engine_kinds: dict[Path, EngineKind] = {}
    if needs_engine:
        # 入口一次性按文件决议引擎，convert / 测量 / pdf 各阶段共用（全生命周期绑定）
        engine_kinds = _resolve_engine_kinds(files, options, is_wincom_supported())
        if any(k == "libreoffice" for k in engine_kinds.values()) and engine is None:
            engine = find_engine()

    events: list[PipelineEvent] = []

    def emit(stage: Stage, src: Path, status: str, message: str = "") -> PipelineEvent:
        ev = PipelineEvent(stage=stage, src=src, status=status, message=message)
        events.append(ev)
        if progress_cb:
            try:
                progress_cb(ev)
            except Exception:
                pass
        return ev

    def is_cancelled() -> bool:
        return bool(should_cancel and should_cancel())

    # 3. 辅助组件
    sz_cmd = _find_seven_zip()
    seven_zip_path = Path(sz_cmd) if sz_cmd else None
    lo_runner: BatchRunner | None = (
        BatchRunner(engine=engine, jobs=options.jobs) if engine else None
    )
    com_runner: WincomRunner | None = None

    def get_runner(kind: EngineKind) -> ConvertRunner:
        nonlocal com_runner
        if kind == "libreoffice":
            if lo_runner is None:
                raise EngineNotFoundError(
                    "任务包含路由至 LibreOffice 的文件，但未找到可用的 soffice 转换引擎。"
                )
            return lo_runner
        if com_runner is None:
            com_runner = WincomRunner()
        return com_runner

    def kind_of(src: Path) -> EngineKind:
        # 未参与决议的文件（如 toc 非 measured 模式）默认 libreoffice
        return engine_kinds.get(src, "libreoffice")

    # -------------------------------------------------------------------------
    # 模式 A: 仅批量转换 (convert)
    # -------------------------------------------------------------------------
    if options.mode == "convert":
        # 单文件页码截取特殊分支
        if options.pages is not None:
            if len(files) != 1:
                raise ValueError("--pages 选项仅支持单文件转换")
            src = files[0]
            p_start, p_end = options.pages
            p_suffix = f"_p{p_start}" if p_start == p_end else f"_p{p_start}-{p_end}"
            out_base = (
                Path(options.out_dir).resolve() if options.out_dir else src.parent
            )
            dst_pdf = out_base / f"{src.stem}{p_suffix}.pdf"

            if dst_pdf.exists() and not options.force:
                emit(Stage.PDF, src, "skipped", f"目标文件已存在: {dst_pdf.name}")
                return events

            if is_cancelled():
                emit(Stage.PDF, src, "cancelled", "用户中止")
                return events

            t0 = time.time()
            kind = kind_of(src)
            with tempfile.TemporaryDirectory(prefix="wdt_pages_") as tmp_dir:
                temp_pdf = Path(tmp_dir) / f"{src.stem}.pdf"
                results = get_runner(kind).convert(
                    [(src, temp_pdf)],
                    target="pdf",
                    gray=options.gray if kind == "libreoffice" else False,
                    should_cancel=should_cancel,
                )
                if is_cancelled() or (results and results[0].error == "cancelled"):
                    emit(Stage.PDF, src, "cancelled", "用户中止")
                    return events
                if not results or not results[0].ok:
                    err = results[0].error if results else "转换 PDF 失败"
                    emit(Stage.PDF, src, "failed", err)
                    return events
                try:
                    slice_pdf(temp_pdf, dst_pdf, p_start, p_end)
                    elapsed = time.time() - t0
                    emit(Stage.PDF, src, "ok", f"{elapsed:.2f}s")
                except Exception as e:
                    emit(Stage.PDF, src, "failed", f"截取页码失败: {e}")
            return events

        # 常规批量转换
        targets_to_run: list[str] = []
        if options.make_docx:
            targets_to_run.append("docx")
        if options.make_pdf:
            targets_to_run.append("pdf")

        for cur_target in targets_to_run:
            stage = Stage.CONVERT if cur_target == "docx" else Stage.PDF
            if is_cancelled():
                for src in files:
                    emit(stage, src, "cancelled", "用户中止")
                continue

            target_ext = f".{cur_target}"
            work_pairs: list[tuple[Path, Path]] = []
            for src in files:
                if src.suffix.lower() == target_ext and options.out_dir is None:
                    continue

                dst = resolve_destination(
                    src, target_ext, options.out_dir, options.base_dir
                )
                if dst.exists() and not options.force:
                    emit(stage, src, "skipped", f"目标已存在: {dst.name}")
                else:
                    work_pairs.append((src, dst))

            if not work_pairs:
                continue

            use_gray = options.gray if cur_target == "pdf" else False

            # 按决议引擎分组调度（LibreOffice 走多进程并发，Word COM 走串行实例队列）
            groups: dict[EngineKind, list[tuple[Path, Path]]] = {
                "libreoffice": [],
                "wincom": [],
            }
            for pair in work_pairs:
                groups[kind_of(pair[0])].append(pair)

            for kind, group_pairs in groups.items():
                if not group_pairs:
                    continue
                label = _ENGINE_LABELS[kind]

                def _on_progress(
                    done: int, total: int, res: FileResult, _label: str = label
                ) -> None:
                    st = (
                        "ok"
                        if res.ok
                        else ("cancelled" if res.error == "cancelled" else "failed")
                    )
                    msg = (
                        f"[{_label}] {res.elapsed:.2f}s"
                        if res.ok
                        else f"[{_label}] {res.error}"
                    )
                    emit(stage, res.src, st, msg)

                convert_results = get_runner(kind).convert(
                    group_pairs,
                    target=cur_target,  # type: ignore[arg-type]
                    gray=use_gray if kind == "libreoffice" else False,
                    progress_cb=_on_progress,
                    should_cancel=should_cancel,
                )
                if cur_target == "docx" and kind == "libreoffice":
                    for r in convert_results or []:
                        if r.ok and r.dst is not None:
                            _safe_normalize_converted(r.dst, seven_zip_path)

        return events

    # -------------------------------------------------------------------------
    # 模式 B: 仅目录注入 (toc)
    # -------------------------------------------------------------------------
    if options.mode == "toc":
        assert options.toc is not None
        for src in files:
            if is_cancelled():
                emit(Stage.TOC, src, "cancelled", "用户中止")
                continue

            if not src.is_file():
                emit(Stage.TOC, src, "failed", f"文件不存在: {src.name}")
                continue

            try:
                toc_cfg = options.toc
                src_engine_kind = kind_of(src)
                if (
                    options.toc.page_mode == PageMode.STATIC_MEASURED
                    and src_engine_kind == "wincom"
                ):
                    # Word COM 原生单遍直出：无需植入 WDTTAG 隐藏测量锚点，保持书签纯净
                    from dataclasses import replace
                    toc_cfg = replace(options.toc, page_mode=PageMode.STATIC_RESTART)

                report: TocReport = inject_toc(
                    src, toc_cfg, seven_zip=seven_zip_path
                )

                measure_msg = ""
                if options.toc.page_mode == PageMode.STATIC_MEASURED:
                    if src_engine_kind == "wincom":
                        # Word COM 原生支持导出时流式刷新域（doc.Fields.Update），
                        # 避免超大文档进行 2-Pass 冗余临时 PDF 测算，实现单遍高效直出。
                        measure_msg = "Word COM 原生单遍直出"
                    else:
                        measure_runner = get_runner(src_engine_kind)
                        succ, final_pages, measure_msg = _perform_page_measurement(
                            src, report, options.toc, measure_runner, seven_zip_path=seven_zip_path
                        )
                        report.pages = final_pages

                pages_summary = (
                    f"{report.pages}"
                    if len(report.pages) <= 5
                    else f"[{report.pages[0]}..{report.pages[-1]}]"
                )
                detail = f"条目: {report.entries}, 页码: {pages_summary}, 耗时: {report.elapsed:.2f}s"
                if measure_msg:
                    detail += f" ({measure_msg})"
                emit(Stage.TOC, src, "ok", detail)
            except NoHeadingError as e:
                emit(Stage.TOC, src, "failed", str(e))
            except Exception as e:
                emit(Stage.TOC, src, "failed", f"未知异常: {e}")

        return events

    # -------------------------------------------------------------------------
    # 模式 C: 端到端一体化流水线 (run)
    # -------------------------------------------------------------------------
    if options.mode == "run":
        assert options.toc is not None

        to_convert: list[tuple[Path, Path]] = []
        ready_docxs: list[tuple[Path, Path]] = []  # (原始源路径, 对应的 DOCX 路径)

        # 阶段 1: 准备 DOCX
        for src in files:
            if is_cancelled():
                emit(Stage.CONVERT, src, "cancelled", "用户中止")
                continue

            dst_docx = resolve_destination(
                src, ".docx", options.out_dir, options.base_dir
            )
            if src.suffix.lower() in (".doc", ".rtf"):
                if dst_docx.exists() and not options.force:
                    emit(Stage.CONVERT, src, "skipped", "目标 DOCX 已存在")
                    ready_docxs.append((src, dst_docx))
                else:
                    to_convert.append((src, dst_docx))
            else:
                if (
                    options.out_dir is not None
                    and src.resolve() != dst_docx.resolve()
                ):
                    dst_docx.parent.mkdir(parents=True, exist_ok=True)
                    if not dst_docx.exists() or options.force:
                        shutil.copy2(src, dst_docx)
                        emit(Stage.CONVERT, src, "ok", "复制到输出目录")
                    else:
                        emit(Stage.CONVERT, src, "skipped", "目标 DOCX 已存在")
                    ready_docxs.append((src, dst_docx))
                else:
                    ready_docxs.append((src, src))

        if to_convert and not is_cancelled():
            conv_groups: dict[EngineKind, list[tuple[Path, Path]]] = {
                "libreoffice": [],
                "wincom": [],
            }
            for pair in to_convert:
                conv_groups[kind_of(pair[0])].append(pair)

            for kind, group_pairs in conv_groups.items():
                if not group_pairs:
                    continue
                label = _ENGINE_LABELS[kind]

                def _on_docx_progress(
                    done: int, total: int, res: FileResult, _label: str = label
                ) -> None:
                    st = (
                        "ok"
                        if res.ok
                        else ("cancelled" if res.error == "cancelled" else "failed")
                    )
                    msg = (
                        f"[{_label}] {res.elapsed:.2f}s"
                        if res.ok
                        else f"[{_label}] {res.error}"
                    )
                    emit(Stage.CONVERT, res.src, st, msg)

                c_results = get_runner(kind).convert(
                    group_pairs,
                    target="docx",
                    progress_cb=_on_docx_progress,
                    should_cancel=should_cancel,
                )
                for r in c_results:
                    if r.ok and r.dst is not None:
                        # wincom 产物为 Word 原生合规 DOCX，跳过 LibreOffice 专用规范化
                        if kind == "libreoffice":
                            _safe_normalize_converted(r.dst, seven_zip_path)
                        ready_docxs.append((r.src, r.dst))

        # 阶段 2: 注入 TOC 目录
        injected_docxs: list[tuple[Path, Path]] = []
        for orig_src, d in ready_docxs:
            if is_cancelled():
                emit(Stage.TOC, orig_src, "cancelled", "用户中止")
                continue

            try:
                toc_cfg = options.toc
                orig_engine_kind = kind_of(orig_src)
                if (
                    options.toc.page_mode == PageMode.STATIC_MEASURED
                    and orig_engine_kind == "wincom"
                ):
                    # Word COM 原生单遍直出：无需植入 WDTTAG 隐藏测量锚点，保持书签纯净
                    from dataclasses import replace
                    toc_cfg = replace(options.toc, page_mode=PageMode.STATIC_RESTART)

                report = inject_toc(d, toc_cfg, seven_zip=seven_zip_path)

                measure_msg = ""
                if options.toc.page_mode == PageMode.STATIC_MEASURED:
                    if orig_engine_kind == "wincom":
                        # Word COM 原生支持导出时流式刷新域（doc.Fields.Update），
                        # 避免超大文档进行 2-Pass 冗余临时 PDF 测算，实现单遍高效直出。
                        measure_msg = "Word COM 原生单遍直出"
                    else:
                        measure_runner = get_runner(orig_engine_kind)
                        succ, final_pages, measure_msg = _perform_page_measurement(
                            d, report, options.toc, measure_runner, seven_zip_path=seven_zip_path
                        )
                        report.pages = final_pages

                injected_docxs.append((orig_src, d))
                msg = f"条目: {report.entries}, 耗时: {report.elapsed:.2f}s"
                if measure_msg:
                    msg += f" ({measure_msg})"
                emit(Stage.TOC, orig_src, "ok", msg)
            except Exception as e:
                emit(Stage.TOC, orig_src, "failed", str(e))

        # 阶段 3: 导出 PDF
        if options.make_pdf:
            pdf_pairs: list[tuple[Path, Path]] = []
            pair_to_orig: dict[Path, Path] = {}

            for orig_src, d in injected_docxs:
                if is_cancelled():
                    emit(Stage.PDF, orig_src, "cancelled", "用户中止")
                    continue

                dst_pdf = d.with_suffix(".pdf")
                if dst_pdf.exists() and not options.force:
                    emit(Stage.PDF, orig_src, "skipped", "目标 PDF 已存在")
                else:
                    pdf_pairs.append((d, dst_pdf))
                    pair_to_orig[d] = orig_src

            if pdf_pairs and not is_cancelled():
                # 按源文件决议的引擎分组导出，与 convert / 测量阶段共用同一引擎
                pdf_groups: dict[EngineKind, list[tuple[Path, Path]]] = {
                    "libreoffice": [],
                    "wincom": [],
                }
                for d, dst_pdf in pdf_pairs:
                    pdf_groups[kind_of(pair_to_orig[d])].append((d, dst_pdf))

                for kind, group_pairs in pdf_groups.items():
                    if not group_pairs:
                        continue
                    label = _ENGINE_LABELS[kind]

                    def _on_pdf_progress(
                        done: int, total: int, res: FileResult, _label: str = label
                    ) -> None:
                        orig = pair_to_orig.get(res.src, res.src)
                        st = (
                            "ok"
                            if res.ok
                            else ("cancelled" if res.error == "cancelled" else "failed")
                        )
                        msg = (
                            f"[{_label}] {res.elapsed:.2f}s"
                            if res.ok
                            else f"[{_label}] {res.error}"
                        )
                        emit(Stage.PDF, orig, st, msg)

                    get_runner(kind).convert(
                        group_pairs,
                        target="pdf",
                        gray=options.gray if kind == "libreoffice" else False,
                        progress_cb=_on_pdf_progress,
                        should_cancel=should_cancel,
                    )

        return events

    raise ValueError(f"不支持的 pipeline mode: {options.mode}")
