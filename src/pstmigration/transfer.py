"""ファイルサーバ転送（設計書 §8.3、§8.4）。"""

from __future__ import annotations

import hashlib
import logging
import os
import time
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path

import pywintypes
import win32con
import win32file
import win32netcon
import win32wnet

from pstmigration.state import StepError

log = logging.getLogger(__name__)

CHUNK = 4 * 1024 * 1024
RESUME_BACKOFF = 64 * 1024 * 1024  # 再開時に捨てる .partial の末尾
ERROR_SESSION_CREDENTIAL_CONFLICT = 1219


def connect(server_root: str, user: str, password: str) -> bool:
    """NAS のローカルアカウントで共有に接続する（ドライブ文字なし・非永続）。

    失敗してもここでは止めない（ローカルの分割・圧縮は進められるため）。転送時に NETWORK エラーになる。
    """
    if not user:
        return True
    try:
        win32wnet.WNetAddConnection2(win32netcon.RESOURCETYPE_DISK, None, server_root, None, user, password, 0)
        return True
    except pywintypes.error as e:
        if e.winerror == ERROR_SESSION_CREDENTIAL_CONFLICT:
            # 同じNASに別の資格情報で接続済み（ユーザーが手でつないでいる等）。その接続で試す
            log.warning("NAS には別の資格情報で接続済みのため、そのまま使います: %s", server_root)
            return True
        log.warning("NAS に接続できません: %s (%s)", server_root, e.strerror)
        return False


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while chunk := f.read(CHUNK):
            h.update(chunk)
    return h.hexdigest()


def open_exclusive(path: Path):
    """共有なしで開く。他プロセスが開いていれば pywintypes.error。"""
    return win32file.CreateFile(str(path), win32con.GENERIC_READ, 0, None, win32con.OPEN_EXISTING, 0, None)


@contextmanager
def slot(server_root: Path, wait_min: float) -> Iterator[None]:
    """_slots\\slot*.lock のどれかを排他オープンして保持する。_slots が無ければ制限なし。"""
    try:
        slots = sorted((server_root / "_slots").glob("slot*.lock"))
    except OSError as ex:
        raise StepError("NETWORK", str(ex), transient=True) from ex
    if not slots:
        yield
        return
    deadline = time.monotonic() + wait_min * 60
    while True:
        for s in slots:
            try:
                handle = open_exclusive(s)
            except pywintypes.error as ex:
                if ex.winerror in (32, 33):  # 共有違反・ロック違反だけを「枠が使用中」と扱う
                    continue
                code = "ACCESS_DENIED" if ex.winerror == 5 else "NETWORK"
                raise StepError(code, str(ex), transient=True) from ex
            log.info("転送枠を取得: %s", s.name)
            try:
                yield
            finally:
                handle.Close()
            return
        if time.monotonic() > deadline:
            raise StepError("SLOT_BUSY", "転送枠が空きません", transient=True)
        time.sleep(30)


def upload(src: Path, dst: Path, sha256: str, max_mbps: float, *, overwrite: bool = False) -> None:
    """一時名で転送→読み直してSHA-256照合→正式名へ。

    既存の正式名ファイルは上書きしない。overwrite=True（毎回作り直すバックアップ用）なら置き換える。
    """
    # 失敗時の診断用（設計書 §18.5）：どの段階で、何バイト目で、何秒後に
    diag = {"stage": "prepare", "destination": str(dst), "offset": 0, "sent": 0, "size": None}
    t0 = time.monotonic()
    try:
        dst.parent.mkdir(parents=True, exist_ok=True)
        part = dst.with_name(dst.name + ".partial")
        if overwrite:
            part.unlink(missing_ok=True)  # 前回の残りは中身が違うので続きからにしない
        elif dst.exists():
            if sha256_file(dst) == sha256:
                log.info("転送済み: %s", dst)
                return
            raise StepError("REMOTE_CONFLICT", f"同名で内容の異なるファイルがあります: {dst}", final=True)

        size = diag["size"] = src.stat().st_size
        offset = part.stat().st_size if part.exists() else 0
        if offset > size:
            part.unlink()
            offset = 0
        if offset:
            # 前回が通信エラーで途切れていると、末尾は NAS に書けたと報告されただけで中身が欠けていることがある。
            # 少し戻して書き直す。ponytail: 戻し幅は固定。途中まで壊れていれば照合で全体をやり直す
            offset = max(0, offset - RESUME_BACKOFF)
            os.truncate(part, offset)
            log.info("転送再開: %s (%d / %d bytes)", dst.name, offset, size)
        diag.update(stage="write", offset=offset)

        rate = max_mbps * 125_000  # Mbps → bytes/sec
        start, sent = time.monotonic(), 0
        with open(src, "rb") as fi, open(part, "ab") as fo:
            fi.seek(offset)
            while chunk := fi.read(CHUNK):
                fo.write(chunk)
                sent += len(chunk)
                diag["sent"] = sent
                if rate > 0:
                    ahead = sent / rate - (time.monotonic() - start)
                    if ahead > 0:
                        time.sleep(ahead)

        diag["stage"] = "verify"
        remote_sha = sha256_file(part)
        if remote_sha != sha256:
            remote_size = part.stat().st_size
            try:
                local_sha = sha256_file(src)
            except OSError:  # Windows Search 等が掴んでいる。ローカルの変化は断定せず再試行に回す
                local_sha = None
            part.unlink()
            raise StepError(
                "HASH_MISMATCH",
                f"転送後のSHA-256が一致しません: {dst}",
                final=local_sha is not None and local_sha != sha256,
                transient=local_sha is None,
                details={
                    "stage": "transfer_verify",
                    "source": str(src),
                    "destination": str(dst),
                    "expected_sha256": sha256,
                    "remote_sha256": remote_sha,
                    "local_sha256": local_sha,
                    "local_matches_expected": local_sha == sha256,
                    "remote_size": remote_size,
                    "size": size,
                    "resumed_from": offset,
                    "elapsed_s": round(time.monotonic() - t0),
                },
            )
        diag["stage"] = "rename"
        if overwrite:
            os.replace(part, dst)
        else:
            os.rename(part, dst)  # Windows の rename は既存ファイルを上書きしない
        secs = time.monotonic() - t0
        log.info(
            "転送完了: %s (%d bytes, 再開位置 %d, %.0f 秒, %.1f Mbps)",
            dst,
            size,
            offset,
            secs,
            sent * 8 / 1e6 / max(secs, 1e-3),
        )
    except StepError:
        raise
    except OSError as e:
        diag.update(elapsed_s=round(time.monotonic() - t0), winerror=getattr(e, "winerror", None), file=e.filename)
        code = "ACCESS_DENIED" if isinstance(e, PermissionError) else "NETWORK"
        raise StepError(code, str(e), transient=True, details=diag) from e
