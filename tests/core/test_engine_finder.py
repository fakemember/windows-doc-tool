from __future__ import annotations

import sys
from pathlib import Path
from unittest.mock import patch

import pytest
from wdt.core.engine_finder import WordEngine, find_word_engine


def test_find_word_engine_non_windows(monkeypatch):
    monkeypatch.setattr(sys, "platform", "linux")
    we = find_word_engine()
    assert we.available is False
    assert "仅支持 Windows" in we.error


def test_find_word_engine_windows():
    we = find_word_engine()
    assert isinstance(we, WordEngine)
    if sys.platform == "win32":
        # 如果当前测试机已安装 Word 且有 pywin32
        if we.available:
            assert we.version != ""
            assert we.error == ""
