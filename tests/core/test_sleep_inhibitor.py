from __future__ import annotations

import sys
from unittest.mock import MagicMock, patch

import pytest

from wdt.core.sleep_inhibitor import (
    ES_CONTINUOUS,
    ES_SYSTEM_REQUIRED,
    SleepInhibitor,
)


def test_sleep_inhibitor_non_win32():
    """非 Windows 环境静默降级，不抛出异常。"""
    with patch("sys.platform", "linux"):
        inhibitor = SleepInhibitor("test linux")
        with inhibitor:
            assert inhibitor._active is False
        assert inhibitor._active is False


@pytest.mark.skipif(sys.platform != "win32", reason="仅在 Windows 平台测试真实 Win32 API 交互")
def test_sleep_inhibitor_real_windows():
    """Windows 真实环境下 SetThreadExecutionState 调用测试。"""
    inhibitor = SleepInhibitor("test real windows")
    with inhibitor:
        assert inhibitor._active is True
    assert inhibitor._active is False


def test_sleep_inhibitor_mock_success():
    """模拟 Windows 下 API 调用成功与释放过程。"""
    mock_kernel32 = MagicMock()
    mock_kernel32.SetThreadExecutionState.return_value = 1  # 成功返回非 0

    with (
        patch("sys.platform", "win32"),
        patch("ctypes.WinDLL", return_value=mock_kernel32),
    ):
        inhibitor = SleepInhibitor("test batch")
        with inhibitor:
            assert inhibitor._active is True
            mock_kernel32.SetThreadExecutionState.assert_called_once_with(
                ES_CONTINUOUS | ES_SYSTEM_REQUIRED
            )

        assert inhibitor._active is False
        assert mock_kernel32.SetThreadExecutionState.call_count == 2
        mock_kernel32.SetThreadExecutionState.assert_called_with(ES_CONTINUOUS)


def test_sleep_inhibitor_mock_failure():
    """模拟 Windows 下 API 调用返回 0 (失败) 时不激活也不崩溃。"""
    mock_kernel32 = MagicMock()
    mock_kernel32.SetThreadExecutionState.return_value = 0

    with (
        patch("sys.platform", "win32"),
        patch("ctypes.WinDLL", return_value=mock_kernel32),
        patch("ctypes.get_last_error", return_value=5),
    ):
        inhibitor = SleepInhibitor("test failure")
        with inhibitor:
            assert inhibitor._active is False

        # 因为未激活，__exit__ 不会再次调用恢复状态
        mock_kernel32.SetThreadExecutionState.assert_called_once()
