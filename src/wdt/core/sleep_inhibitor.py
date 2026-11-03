"""Windows 电源防睡眠保障模块。

在长耗时任务（批量转换/PDF导出）期间阻止 Windows 因空闲超时自动休眠。
特性：
- 允许用户正常 Win + L 锁屏与显示器空闲关闭，不阻止用户手动点击睡眠。
- 仅在 Windows 平台生效；非 Windows 环境静默降级（便于跨平台测试）。
- 异常或退出时通过 finally 保障释放恢复。
"""

from __future__ import annotations

import logging
import sys

logger = logging.getLogger(__name__)

ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001


class SleepInhibitor:
    """转换期间阻止系统因空闲超时而自动睡眠。"""

    def __init__(self, reason: str = "wdt batch document processing"):
        self.reason = reason
        self._active = False

    def __enter__(self) -> SleepInhibitor:
        if sys.platform != "win32":
            return self

        try:
            import ctypes
            from ctypes import wintypes

            kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
            set_execution_state = kernel32.SetThreadExecutionState
            set_execution_state.argtypes = [wintypes.DWORD]
            set_execution_state.restype = wintypes.DWORD

            prev_state = set_execution_state(ES_CONTINUOUS | ES_SYSTEM_REQUIRED)
            if prev_state == 0:
                err = ctypes.get_last_error()
                logger.warning(
                    "启用防睡眠失败 (SetThreadExecutionState 返回 0, WinError=%d)", err
                )
            else:
                self._active = True
                logger.info("已启用防系统空闲休眠保障 (reason: %s)", self.reason)
        except Exception as e:
            logger.warning("初始化 SleepInhibitor 异常: %s", e)

        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> bool:
        if sys.platform != "win32" or not self._active:
            return False

        try:
            import ctypes
            from ctypes import wintypes

            kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
            set_execution_state = kernel32.SetThreadExecutionState
            set_execution_state.argtypes = [wintypes.DWORD]
            set_execution_state.restype = wintypes.DWORD

            set_execution_state(ES_CONTINUOUS)
            logger.info("已恢复默认系统电源管理状态")
        except Exception as e:
            logger.warning("恢复电源管理状态异常: %s", e)
        finally:
            self._active = False

        return False
