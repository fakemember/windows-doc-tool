"""
Windows Doc Tool (wdt) 业务流水线模块。

提供脱离 CLI 与 GUI 交互的统一核心编排 API：
- 批量转换 (convert)
- 目录注入 (toc)
- 级联流水线 (run: 转换 -> TOC注入 -> 导出PDF)
"""

from __future__ import annotations

import shutil
import tempfile
import time
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Callable, Literal, Sequence

from pypdf import PdfReader, PdfWriter

from wdt.core.batch_runner import BatchRunner
from wdt.core.converter_worker import FileResult
from wdt.core.engine_finder import Engine, EngineNotFoundError, find_engine
from wdt.toc.docx_stream_patcher import (
    NoHeadingError,
    TocConfig,
    TocReport,
    inject_toc,
)
from wdt.toc.page_calculator import PageMode
from wdt.toc.zip_rewriter import _find_seven_zip


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

    # 2. 转换引擎准备
    if options.mode in ("convert", "run"):
        if engine is None:
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
    runner = BatchRunner(engine=engine, jobs=options.jobs) if engine else None

    # -------------------------------------------------------------------------
    # 模式 A: 仅批量转换 (convert)
    # -------------------------------------------------------------------------
    if options.mode == "convert":
        assert runner is not None

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
            with tempfile.TemporaryDirectory(prefix="wdt_pages_") as tmp_dir:
                temp_pdf = Path(tmp_dir) / f"{src.stem}.pdf"
                results = runner.convert(
                    [(src, temp_pdf)],
                    target="pdf",
                    gray=options.gray,
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

            def _on_progress(done: int, total: int, res: FileResult) -> None:
                st = (
                    "ok"
                    if res.ok
                    else ("cancelled" if res.error == "cancelled" else "failed")
                )
                emit(
                    stage,
                    res.src,
                    st,
                    f"{res.elapsed:.2f}s" if res.ok else res.error,
                )

            runner.convert(
                work_pairs,
                target=cur_target,  # type: ignore[arg-type]
                gray=use_gray,
                progress_cb=_on_progress,
                should_cancel=should_cancel,
            )

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
                report: TocReport = inject_toc(
                    src, options.toc, seven_zip=seven_zip_path
                )
                pages_summary = (
                    f"{report.pages}"
                    if len(report.pages) <= 5
                    else f"[{report.pages[0]}..{report.pages[-1]}]"
                )
                emit(
                    Stage.TOC,
                    src,
                    "ok",
                    f"条目: {report.entries}, 页码: {pages_summary}, 耗时: {report.elapsed:.2f}s",
                )
            except NoHeadingError as e:
                emit(Stage.TOC, src, "failed", str(e))
            except Exception as e:
                emit(Stage.TOC, src, "failed", f"未知异常: {e}")

        return events

    # -------------------------------------------------------------------------
    # 模式 C: 端到端一体化流水线 (run)
    # -------------------------------------------------------------------------
    if options.mode == "run":
        assert runner is not None
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
            def _on_docx_progress(done: int, total: int, res: FileResult) -> None:
                st = (
                    "ok"
                    if res.ok
                    else ("cancelled" if res.error == "cancelled" else "failed")
                )
                emit(
                    Stage.CONVERT,
                    res.src,
                    st,
                    f"{res.elapsed:.2f}s" if res.ok else res.error,
                )

            c_results = runner.convert(
                to_convert,
                target="docx",
                progress_cb=_on_docx_progress,
                should_cancel=should_cancel,
            )
            for r in c_results:
                if r.ok and r.dst is not None:
                    ready_docxs.append((r.src, r.dst))

        # 阶段 2: 注入 TOC 目录
        injected_docxs: list[tuple[Path, Path]] = []
        for orig_src, d in ready_docxs:
            if is_cancelled():
                emit(Stage.TOC, orig_src, "cancelled", "用户中止")
                continue

            try:
                report = inject_toc(d, options.toc, seven_zip=seven_zip_path)
                injected_docxs.append((orig_src, d))
                emit(
                    Stage.TOC,
                    orig_src,
                    "ok",
                    f"条目: {report.entries}, 耗时: {report.elapsed:.2f}s",
                )
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
                def _on_pdf_progress(done: int, total: int, res: FileResult) -> None:
                    orig = pair_to_orig.get(res.src, res.src)
                    st = (
                        "ok"
                        if res.ok
                        else ("cancelled" if res.error == "cancelled" else "failed")
                    )
                    emit(
                        Stage.PDF,
                        orig,
                        st,
                        f"{res.elapsed:.2f}s" if res.ok else res.error,
                    )

                runner.convert(
                    pdf_pairs,
                    target="pdf",
                    gray=options.gray,
                    progress_cb=_on_pdf_progress,
                    should_cancel=should_cancel,
                )

        return events

    raise ValueError(f"不支持的 pipeline mode: {options.mode}")
