from __future__ import annotations

import re
import sys
from dataclasses import dataclass
from pathlib import Path


class EngineNotFoundError(RuntimeError):
    """未找到可用的 Word COM 转换引擎。"""


@dataclass(frozen=True)
class WordEngine:
    available: bool
    version: str  # 如 "16.0"，不可用为 ""
    app_path: Path | None
    error: str = ""


def find_word_engine() -> WordEngine:
    """轻量探测 Microsoft Word 引擎安装情况。

    仅在 Windows 平台通过注册表或轻量 COM 查询，绝不调用 CreateObject 启动
    完整 Word.Application 进程，毫秒级返回。
    """
    if sys.platform != "win32":
        return WordEngine(
            available=False,
            version="",
            app_path=None,
            error="Word COM 引擎仅支持 Windows 环境",
        )

    try:
        import winreg
    except ImportError:
        return WordEngine(
            available=False,
            version="",
            app_path=None,
            error="系统缺少 winreg 注册表支持模块",
        )

    # 1. 查询 HKEY_CLASSES_ROOT\Word.Application\CurVer 获取主版本号
    cur_ver: str = ""
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CurVer") as key:
            val, _ = winreg.QueryValueEx(key, None)
            if isinstance(val, str) and val.strip():
                cur_ver = val.strip()
    except OSError:
        pass
    except Exception as e:
        return WordEngine(
            available=False,
            version="",
            app_path=None,
            error=f"读取 Word 注册表失败: {e}",
        )

    if not cur_ver:
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Word.Application") as key:
                cur_ver = "Word.Application"
        except OSError:
            return WordEngine(
                available=False,
                version="",
                app_path=None,
                error="未检测到 Microsoft Word 注册信息",
            )
        except Exception as e:
            return WordEngine(
                available=False,
                version="",
                app_path=None,
                error=f"读取 Word 注册表失败: {e}",
            )

    # 从 ProgID (如 Word.Application.16) 中提取版本字符串 (如 "16.0")
    version_str = ""
    m = re.search(r"Word\.Application\.(\d+)", cur_ver, re.IGNORECASE)
    if m:
        version_str = f"{m.group(1)}.0"

    # 2. 查询 CLSID 及其 LocalServer32 获取可执行文件路径
    clsid: str = ""
    for subkey in (rf"{cur_ver}\CLSID", r"Word.Application\CLSID"):
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, subkey) as key:
                val, _ = winreg.QueryValueEx(key, None)
                if isinstance(val, str) and val.strip():
                    clsid = val.strip()
                    break
        except OSError:
            continue
        except Exception:
            continue

    app_path: Path | None = None
    if clsid:
        for root_hkey in (winreg.HKEY_CLASSES_ROOT, winreg.HKEY_LOCAL_MACHINE):
            sub = (
                rf"CLSID\{clsid}\LocalServer32"
                if root_hkey == winreg.HKEY_CLASSES_ROOT
                else rf"Software\Classes\CLSID\{clsid}\LocalServer32"
            )
            try:
                with winreg.OpenKey(root_hkey, sub) as key:
                    val, _ = winreg.QueryValueEx(key, None)
                    if isinstance(val, str) and val.strip():
                        raw_str = val.strip()
                        raw_clean = raw_str.strip()
                        if raw_clean.startswith('"'):
                            end_quote = raw_clean.find('"', 1)
                            exe_candidate = (
                                raw_clean[1:end_quote]
                                if end_quote != -1
                                else raw_clean[1:]
                            )
                        else:
                            if ".exe" in raw_clean.lower():
                                idx = raw_clean.lower().find(".exe") + 4
                                exe_candidate = raw_clean[:idx].strip()
                            else:
                                exe_candidate = raw_clean.split()[0]
                        if exe_candidate:
                            p = Path(exe_candidate)
                            if p.is_file():
                                app_path = p.resolve()
                                break
            except OSError:
                continue
            except Exception:
                continue

    # 3. 检查 pywin32 是否安装
    try:
        import win32com  # noqa: F401
    except ImportError:
        return WordEngine(
            available=False,
            version=version_str,
            app_path=app_path,
            error="已检测到 Word 安装，但当前 Python 环境未安装 pywin32 依赖",
        )
    except Exception as e:
        return WordEngine(
            available=False,
            version=version_str,
            app_path=app_path,
            error=f"pywin32 环境异常: {e}",
        )

    return WordEngine(
        available=True,
        version=version_str or "已安装",
        app_path=app_path,
        error="",
    )
