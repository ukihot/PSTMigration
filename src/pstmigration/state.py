"""state.json と状態遷移（設計書 §9）。"""

from __future__ import annotations

import json
import logging
import os
import traceback
from datetime import datetime
from pathlib import Path
from typing import Any

from pstmigration import __version__

log = logging.getLogger(__name__)

TERMINAL = {"COMPLETED", "FAILED", "SKIPPED"}


class StepError(Exception):
    """処理段階の失敗。

    transient: 一時的（ネットワーク・容量・ロック等）。試行回数を数えず次回再試行。
    final:     再試行しても無駄。即 FAILED。
    どちらでもない: 試行回数を数え、max_attempts で FAILED。
    """

    def __init__(
        self, code: str, message: str = "", *, transient: bool = False, final: bool = False, details: dict | None = None
    ):
        super().__init__(f"{code}: {message}")
        self.code = code
        self.message = message
        self.transient = transient
        self.final = final
        self.details = details


def now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def read_json(path: Path, default: Any) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        return default


def write_json(path: Path, data: Any) -> None:
    """一時ファイル→置換で原子的に書く。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + ".tmp")
    # MAPI が孤立サロゲートを返しても文字を捨てず、JSON の \uXXXX として保存する。
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8", errors="backslashreplace")
    os.replace(tmp, path)


def new_entry(key: str, path: str, is_default: bool) -> dict:
    return {
        "key": key,
        "source_path": path,
        "is_default_store": is_default,
        "status": "PENDING",
        "checkpoint": None,
        "snapshot_at": None,
        "source_size": None,
        "remote_dir": None,
        "generation": 1,
        "counts": {},
        "folder_counts": {},
        "outputs": {},
        "durations": {},
        "warnings": [],
        "problems": [],
        "attempts": 0,
        "last_error": None,
        "detected_at": now_iso(),
    }


def apply_error(entry: dict, err: StepError, max_attempts: int, *, cause: Exception | None = None) -> None:
    failure = {
        "code": err.code,
        "message": err.message[:500],
        "at": now_iso(),
        "phase": entry["status"],
        "checkpoint": entry.get("checkpoint"),
        "pst_key": entry["key"],
        "tool_version": __version__,
        "generation": entry.get("generation", 1),
    }
    if err.details:
        failure["details"] = err.details
    if cause is not None:
        failure.update(
            exception_type=f"{type(cause).__module__}.{type(cause).__qualname__}",
            exception_message=str(cause)[:2000],
            # 内部例外も含む。メール本文・添付やフレームのローカル変数は収集しない。
            traceback="".join(traceback.format_exception(cause))[-16000:],
        )
    entry["last_error"] = failure
    if err.final:
        entry["status"] = "FAILED"
    elif err.transient:
        entry["status"] = "RETRY_PENDING"
    else:
        entry["attempts"] = entry.get("attempts", 0) + 1
        entry["status"] = "FAILED" if entry["attempts"] >= max_attempts else "RETRY_PENDING"
    failure["retryable"] = entry["status"] == "RETRY_PENDING"
    log.warning("PST %s: %s -> %s", entry["key"], err, entry["status"])


def warn(entry: dict, code: str) -> None:
    if code not in entry["warnings"]:
        entry["warnings"].append(code)


class State:
    def __init__(self, path: Path):
        self.path = path
        self.data = read_json(path, {"psts": {}})

    @property
    def psts(self) -> dict[str, dict]:
        return self.data["psts"]

    def save(self) -> None:
        self.data["tool_version"] = __version__
        self.data["updated_at"] = now_iso()
        write_json(self.path, self.data)
