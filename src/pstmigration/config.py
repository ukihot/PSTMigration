"""config.json の読込と検証（設計書 §12）。"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass
from datetime import UTC, datetime
from pathlib import Path


@dataclass(frozen=True)
class Config:
    cutoff: str = "2021-01-01T00:00:00+09:00"
    server_root: str = r"\\192.168.8.120\PSTBackup"
    # NAS は AD 非参加のため、NAS のローカルアカウントで接続する。空なら現在の資格情報のまま
    server_user: str = "tokunaga"
    server_password: str = "***REMOVED***"
    work_dir: str = r"%LOCALAPPDATA%\PSTMigration\work"
    disk_factor: float = 2.5
    disk_reserve_gb: float = 5
    logoff_snapshot_budget_gb: float = 10
    split_pause_ms: int = 200
    release_wait_sec: int = 900
    sevenzip_level: int = 5
    sevenzip_threads: int = 2
    transfer_max_mbps: float = 50
    slot_wait_min: int = 30
    new_pst_warn_gb: float = 20
    max_attempts: int = 3
    recollect_on_change: bool = False
    log_subject: bool = False

    def __post_init__(self) -> None:
        if self.cutoff_local.tzinfo is None:
            raise ValueError("cutoff にはタイムゾーン（例 +09:00）が必要です")
        if not 0 <= self.sevenzip_level <= 9:
            raise ValueError("sevenzip_level は 0〜9")
        if self.disk_factor < 1 or self.max_attempts < 1 or self.sevenzip_threads < 1:
            raise ValueError("disk_factor / max_attempts / sevenzip_threads は 1 以上")

    @property
    def cutoff_local(self) -> datetime:
        return datetime.fromisoformat(self.cutoff)

    @property
    def cutoff_utc(self) -> datetime:
        return self.cutoff_local.astimezone(UTC)

    @property
    def work_path(self) -> Path:
        return Path(os.path.expandvars(self.work_dir))

    @property
    def server_path(self) -> Path:
        return Path(self.server_root)


def load(path: Path) -> Config:
    """不正な設定は例外（未知キーは TypeError、値不正は ValueError）。"""
    return Config(**json.loads(path.read_text(encoding="utf-8")))
