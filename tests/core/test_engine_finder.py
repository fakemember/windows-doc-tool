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
    clear_staged_update,
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


def test_find_engine_bundled(tmp_path: Path, monkeypatch):
    bundled = tmp_path / "runtime" / "libreoffice"
    prog = bundled / "program"
    prog.mkdir(parents=True)

    fake_com = prog / "soffice.com"
    fake_exe = prog / "soffice.exe"

    fake_com.write_text("dummy")
    fake_exe.write_text("dummy")

    fake_com.chmod(fake_com.stat().st_mode | stat.S_IEXEC)
    fake_exe.chmod(fake_exe.stat().st_mode | stat.S_IEXEC)

    monkeypatch.setattr(
        "wdt.core.engine_finder._parse_version",
        lambda p: (25, 2) if p.name == "soffice.com" else (24, 8),
    )

    # 优先使用 soffice.com
    engine = find_engine(bundled_dir=bundled)
    assert engine.source == "bundled"
    assert engine.soffice == fake_com.resolve()
    assert engine.version == (25, 2)


def test_find_engine_extra_paths(tmp_path: Path, monkeypatch):
    extra = tmp_path / "custom_lo"
    extra.mkdir()
    fake_soffice = extra / "soffice"
    fake_soffice.write_text("dummy")
    fake_soffice.chmod(fake_soffice.stat().st_mode | stat.S_IEXEC)

    monkeypatch.setattr("wdt.core.engine_finder._parse_version", lambda p: (7, 6))

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


def test_clear_staged_update_removes_updating_dir(tmp_path: Path):
    """soffice 位于 program/ 下时，应清除安装根目录的 updating/ 暂存区。"""
    prog = tmp_path / "LibreOffice" / "program"
    prog.mkdir(parents=True)
    staging = tmp_path / "LibreOffice" / "updating"
    staging.mkdir()
    (staging / "update.dat").write_text("staged")
    soffice = prog / "soffice.com"
    soffice.write_text("dummy")

    assert clear_staged_update(soffice) is True
    assert not staging.exists()


def test_clear_staged_update_flat_layout(tmp_path: Path):
    """soffice 直接位于安装根目录时，应清除同级 updating/ 暂存区。"""
    staging = tmp_path / "updating"
    staging.mkdir()
    soffice = tmp_path / "soffice"
    soffice.write_text("dummy")

    assert clear_staged_update(soffice) is True
    assert not staging.exists()


def test_clear_staged_update_noop(tmp_path: Path):
    soffice = tmp_path / "program" / "soffice"
    soffice.parent.mkdir(parents=True)
    soffice.write_text("dummy")

    assert clear_staged_update(soffice) is False


def test_find_engine_clears_staged_update(tmp_path: Path, monkeypatch):
    bundled = tmp_path / "runtime" / "libreoffice"
    prog = bundled / "program"
    prog.mkdir(parents=True)
    soffice = prog / "soffice.com"
    soffice.write_text("dummy")
    soffice.chmod(soffice.stat().st_mode | stat.S_IEXEC)

    staging = bundled / "updating"
    staging.mkdir()
    (staging / "update.dat").write_text("staged")

    monkeypatch.setattr("wdt.core.engine_finder._parse_version", lambda p: (25, 2))

    engine = find_engine(bundled_dir=bundled)
    assert engine.source == "bundled"
    assert not staging.exists()
