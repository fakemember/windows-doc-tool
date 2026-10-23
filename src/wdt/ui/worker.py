"""
GUI 异步流水线工作线程模块。

利用 QThread 承载 run_pipeline 调度，主进程与 UI 保持解耦，实时通知 PipelineEvent。
"""

from __future__ import annotations

from pathlib import Path
from typing import Sequence

from PySide6.QtCore import QThread, Signal

from wdt.core.engine_finder import Engine
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
        engine: Engine | None = None,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.files = list(files)
        self.options = options
        self.engine = engine
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
                engine=self.engine,
            )
            self.finished_signal.emit(results)
        except Exception as e:
            self.error_signal.emit(str(e))
