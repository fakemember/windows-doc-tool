from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass
class FileResult:
    src: Path
    dst: Path | None
    ok: bool
    elapsed: float
    error: str
