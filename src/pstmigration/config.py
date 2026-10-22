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
    server_user: str = ""
    server_password: str = ""
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
    # ライブ分割方式（設計書 §16）
    live_split: str = "auto"  # auto: 現行方式で処理できないPSTだけ / always / off
    chunk_max_gb: float = 5
    chunk_min_gb: float = 1
    chunk_disk_factor: float = 2.5  # 束の目標サイズに対して必要な空きの比率（束PST＋7z＋余裕）
    batch_target_mb: float = 100  # 1回の CopyMessages の目標（PR_MESSAGE_SIZE の合計）
    batch_items: int = 500
    plan_size_factor: float = 1.15  # 計画・見積もり用。容量の保証には使わない

    def __post_init__(self) -> None:
        if self.cutoff_local.tzinfo is None:
            raise ValueError("cutoff にはタイムゾーン（例 +09:00）が必要です")
        if not 0 <= self.sevenzip_level <= 9:
            raise ValueError("sevenzip_level は 0〜9")
        if self.disk_factor < 1 or self.max_attempts < 1 or self.sevenzip_threads < 1:
            raise ValueError("disk_factor / max_attempts / sevenzip_threads は 1 以上")
        if self.live_split not in ("auto", "always", "off"):
            raise ValueError("live_split は auto / always / off")
        # 0 以下だと束のコピーが1件も進まず、無限ループになる
        if self.batch_target_mb <= 0 or self.batch_items < 1:
            raise ValueError("batch_target_mb は 0 より大きく、batch_items は 1 以上")
        if not 0 < self.chunk_min_gb <= self.chunk_max_gb:
            raise ValueError("0 < chunk_min_gb <= chunk_max_gb")
        if self.chunk_disk_factor < 1 or self.plan_size_factor < 1:
            raise ValueError("chunk_disk_factor / plan_size_factor は 1 以上")

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
