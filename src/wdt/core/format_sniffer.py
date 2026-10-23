from __future__ import annotations

import zipfile
from pathlib import Path

_OLE2_MAGIC = bytes.fromhex("D0CF11E0A1B11AE1")


def sniff(path: Path) -> str:
    """根据二进制特征头嗅探文档真实格式。

    返回 "rtf" | "doc" | "docx" | "unknown"。
    """
    p = Path(path)
    try:
        if not p.is_file():
            return "unknown"
        with open(p, "rb") as f:
            header = f.read(512)
        if b"{\\rtf" in header:
            return "rtf"
        if header.startswith(_OLE2_MAGIC):
            return "doc"
        if zipfile.is_zipfile(p):
            try:
                with zipfile.ZipFile(p, "r") as zf:
                    if "word/document.xml" in zf.namelist():
                        return "docx"
            except Exception:
                return "unknown"
    except Exception:
        return "unknown"
    return "unknown"
