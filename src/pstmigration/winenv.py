"""PCの環境（Outlook のビット数など）と、ファイルを掴んでいるプロセスの調査（設計書 §18.5）。

診断用で、失敗しても例外を出さない。
"""

from __future__ import annotations

import contextlib
import ctypes
import os
import platform
import shutil
import struct
import winreg
from ctypes import wintypes
from pathlib import Path

TOOL_BITS = struct.calcsize("P") * 8


def _reg(path: str, name: str, view: int) -> str | None:
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path, 0, winreg.KEY_READ | view) as k:
            return str(winreg.QueryValueEx(k, name)[0])
    except OSError:
        return None


def outlook() -> dict:
    """{"bitness": "x86"|"x64"|None, "version": str|None}。クイック実行版と MSI 版（2010〜）を見る。"""
    v64, v32 = winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY
    c2r = "SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration"
    platform_ = _reg(c2r, "Platform", v64)
    if platform_:
        return {"bitness": platform_.lower(), "version": _reg(c2r, "VersionToReport", v64)}
    for ver in ("16.0", "15.0", "14.0"):
        key = f"SOFTWARE\\Microsoft\\Office\\{ver}\\Outlook"
        b = _reg(key, "Bitness", v64)
        if b is None and _reg(key, "Bitness", v32) is not None:
            b = "x86"  # 32ビット版の MSI は WOW6432Node 側にだけある
        if b:
            return {"bitness": b.lower(), "version": ver}
    return {"bitness": None, "version": None}


def _dir_bytes(path: Path) -> int:
    total = 0
    for root, _, files in os.walk(path):
        for f in files:
            with contextlib.suppress(OSError):
                total += os.path.getsize(os.path.join(root, f))
    return total


def snapshot(work: Path) -> dict:
    """status.json に載せる環境情報。ビット数の不一致・容量不足を台帳から見分けるため。"""
    ol = outlook()
    out = {
        "tool_bits": TOOL_BITS,
        "outlook_bitness": ol["bitness"],
        "office_version": ol["version"],
        "windows": platform.version(),
    }
    try:
        anchor = work if work.exists() else Path(work.anchor)
        out["free_gb"] = round(shutil.disk_usage(anchor).free / 1024**3, 1)
        out["work_gb"] = round(_dir_bytes(work) / 1024**3, 2) if work.exists() else 0
    except OSError:
        pass
    return out


# ---- Restart Manager：ファイルを使用中のプロセス -------------------------------------------------


class _UniqueProcess(ctypes.Structure):
    _fields_ = [("dwProcessId", wintypes.DWORD), ("ProcessStartTime", wintypes.FILETIME)]


class _ProcessInfo(ctypes.Structure):
    _fields_ = [
        ("Process", _UniqueProcess),
        ("strAppName", wintypes.WCHAR * 256),
        ("strServiceShortName", wintypes.WCHAR * 64),
        ("ApplicationType", ctypes.c_int),
        ("AppStatus", wintypes.ULONG),
        ("TSSessionId", wintypes.DWORD),
        ("bRestartable", wintypes.BOOL),
    ]


def _exe_name(pid: int) -> str:
    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    h = k32.OpenProcess(0x1000, False, pid)  # PROCESS_QUERY_LIMITED_INFORMATION
    if not h:
        return ""
    try:
        buf, n = ctypes.create_unicode_buffer(1024), wintypes.DWORD(1024)
        return Path(buf.value).name if k32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(n)) else ""
    finally:
        k32.CloseHandle(h)


def holders(path: Path) -> list[str]:
    """path を開いているプロセス（例 "SearchProtocolHost.exe pid=123"）。調べられなければ空。"""
    try:
        rm = ctypes.WinDLL("rstrtmgr")
        session, key = wintypes.DWORD(), ctypes.create_unicode_buffer(64)
        if rm.RmStartSession(ctypes.byref(session), 0, key):
            return []
        try:
            files = (ctypes.c_wchar_p * 1)(str(path))
            if rm.RmRegisterResources(session, 1, files, 0, None, 0, None):
                return []
            needed, count, reasons = wintypes.UINT(), wintypes.UINT(0), wintypes.DWORD()
            rc = rm.RmGetList(session, ctypes.byref(needed), ctypes.byref(count), None, ctypes.byref(reasons))
            if rc not in (0, 234) or not needed.value:  # 234 = ERROR_MORE_DATA
                return []
            info, count = (_ProcessInfo * needed.value)(), wintypes.UINT(needed.value)
            if rm.RmGetList(session, ctypes.byref(needed), ctypes.byref(count), info, ctypes.byref(reasons)):
                return []
            out = []
            for p in info[: count.value]:
                pid = p.Process.dwProcessId
                name = _exe_name(pid) or p.strAppName
                svc = f"[{p.strServiceShortName}]" if p.strServiceShortName else ""
                out.append(f"{name}{svc} pid={pid}")
            return out
        finally:
            rm.RmEndSession(session)
    except Exception:  # 診断用。調べられなくても本処理は続ける
        return []
