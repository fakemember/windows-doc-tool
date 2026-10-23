import os
import shutil
import stat
from pathlib import Path
from unittest.mock import patch

import pytest
from wdt.core.engine_finder import (
    Engine,
    EngineNotFoundError,
    _parse_version,
    find_engine,
)


@pytest.mark.skipif(shutil.which("soffice") is None, reason="系统未安装 soffice")
def test_find_real_engine():
    engine = find_engine()
    assert isinstance(engine, Engine)
    assert engine.soffice.exists()
    assert engine.source in ("path", "system")
    assert engine.version is not None
    assert engine.version[0] >= 7


def test_find_engine_bundled(tmp_path: Path):
    bundled = tmp_path / "runtime" / "libreoffice"
    prog = bundled / "program"
    prog.mkdir(parents=True)

    fake_com = prog / "soffice.com"
    fake_exe = prog / "soffice.exe"

    fake_com.write_text("#!/bin/sh\necho 'LibreOffice 25.2.0.1'\n")
    fake_exe.write_text("#!/bin/sh\necho 'LibreOffice 24.8.0.1'\n")

    fake_com.chmod(fake_com.stat().st_mode | stat.S_IEXEC)
    fake_exe.chmod(fake_exe.stat().st_mode | stat.S_IEXEC)

    # 优先使用 soffice.com
    engine = find_engine(bundled_dir=bundled)
    assert engine.source == "bundled"
    assert engine.soffice == fake_com.resolve()
    assert engine.version == (25, 2)


def test_find_engine_extra_paths(tmp_path: Path):
    extra = tmp_path / "custom_lo"
    extra.mkdir()
    fake_soffice = extra / "soffice"
    fake_soffice.write_text("#!/bin/sh\necho 'LibreOffice 7.6.0.0'\n")
    fake_soffice.chmod(fake_soffice.stat().st_mode | stat.S_IEXEC)

    engine = find_engine(extra_paths=[fake_soffice])
    assert engine.source == "bundled"
    assert engine.soffice == fake_soffice.resolve()
    assert engine.version == (7, 6)


def test_find_engine_not_found(tmp_path: Path):
    with patch("shutil.which", return_value=None), \
         patch("pathlib.Path.glob", return_value=[]), \
         patch("wdt.core.engine_finder._is_executable", return_value=False):
        with pytest.raises(EngineNotFoundError):
            find_engine()


def test_parse_version_failure(tmp_path: Path):
    dummy = tmp_path / "dummy"
    dummy.write_text("not an executable")
    assert _parse_version(dummy) is None
