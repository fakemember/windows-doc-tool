"""
GUI 异步流水线工作线程模块。

利用 QThread 承载 run_pipeline 调度，主进程与 UI 保持解耦，实时通知 PipelineEvent。
"""

from __future__ import annotations

from pathlib import Path
from typing import Sequence

from PySide6.QtCore import QThread, Signal

from wdt.pipeline import PipelineEvent, PipelineOptions, run_pipeline


class PipelineWorker(QThread):
    """后台流水线工作线程。"""

    event_signal = Signal(object)      # 发射 PipelineEvent
    finished_signal = Signal(list)    # 发射 list[PipelineEvent]
    error_signal = Signal(str)        # 发射严重未捕获异常文本

    def __init__(
        self,
        files: Sequence[Path],
        options: PipelineOptions,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.files = list(files)
        self.options = options
        self._cancelled = False

    def cancel(self) -> None:
        """协作式请求中止当前流水线。"""
        self._cancelled = True

    def _should_cancel(self) -> bool:
        return self._cancelled

    def _on_event(self, ev: PipelineEvent) -> None:
        self.event_signal.emit(ev)

    def run(self) -> None:
        """执行后台流水线处理。"""
        try:
            results = run_pipeline(
                files=self.files,
                options=self.options,
                progress_cb=self._on_event,
                should_cancel=self._should_cancel,
            )
            self.finished_signal.emit(results)
        except Exception as e:
            self.error_signal.emit(str(e))


class PdfSplitWorker(QThread):
    """后台独立 PDF 拆分工作线程。"""

    progress_signal = Signal(int, int, str)  # current, total, message
    finished_signal = Signal(list)           # list[SplitResult]
    error_signal = Signal(str)

    def __init__(
        self,
        pdf_path: Path,
        out_dir: Path | None,
        max_pages: int = 500,
        with_visual_toc: bool = True,
        style=None,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.pdf_path = Path(pdf_path)
        self.out_dir = Path(out_dir) if out_dir else None
        self.max_pages = max_pages
        self.with_visual_toc = with_visual_toc
        self.style = style

    def run(self) -> None:
        try:
            from wdt.core.pdf_splitter import split_pdf_by_outline
            from wdt.toc.styles_templates import TocStyle

            use_style = self.style or TocStyle.REGULATORY_LISTING
            self.progress_signal.emit(0, 100, f"正在解析大纲并切分: {self.pdf_path.name}")
            results = split_pdf_by_outline(
                src_pdf=self.pdf_path,
                out_dir=self.out_dir,
                max_pages=self.max_pages,
                with_visual_toc=self.with_visual_toc,
                style=use_style,
            )
            self.progress_signal.emit(100, 100, f"拆分完成，共生成 {len(results)} 个分卷")
            self.finished_signal.emit(results)
        except Exception as e:
            self.error_signal.emit(str(e))
