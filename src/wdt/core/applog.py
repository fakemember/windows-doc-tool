"""wdt 全局日志配置。

GUI 双击运行的 exe 无控制台，所有诊断输出写入
``%LOCALAPPDATA%\\wdt\\wdt.log``（滚动保留 3 份，单份 5MB）。
CLI 场景额外镜像到 stderr。
"""

from __future__ import annotations

import logging
import os
import sys
from logging.handlers import RotatingFileHandler
from pathlib import Path

_CONFIGURED_ATTR = "_wdt_logging_configured"


def log_path() -> Path:
    base = os.environ.get("LOCALAPPDATA") or str(Path.home())
    d = Path(base) / "wdt"
    d.mkdir(parents=True, exist_ok=True)
    return d / "wdt.log"


def setup_logging() -> Path:
    """配置根 logger：文件滚动 + 有控制台时镜像 stderr。返回日志文件路径。"""
    path = log_path()
    root = logging.getLogger()
    if getattr(root, _CONFIGURED_ATTR, False):
        return path

    root.setLevel(logging.INFO)
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")

    fh = RotatingFileHandler(path, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8")
    fh.setFormatter(fmt)
    root.addHandler(fh)

    if sys.stderr is not None:
        try:
            ch = logging.StreamHandler(sys.stderr)
            ch.setFormatter(fmt)
            root.addHandler(ch)
        except Exception:
            pass

    setattr(root, _CONFIGURED_ATTR, True)
    logging.getLogger(__name__).info("wdt logging initialized: %s", path)
    return path
