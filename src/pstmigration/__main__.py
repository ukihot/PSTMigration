"""PSTMigration エントリポイント（設計書 §2.1）。

GPO から無人実行されるため画面は一切出さない。結果はすべて JSON ログ・state・サーバへ。

  PSTMigration.exe                 ログオン時：検出・分割・圧縮・転送
  PSTMigration.exe --snapshot      ログオフ時：元PSTのスナップショット取得
  PSTMigration.exe --scan-only     調査のみ（パイロット用）
  PSTMigration.exe --aggregate DIR 管理者用：DIR\\_admin に台帳CSV等を出力
  PSTMigration.exe --reset KEY     管理者用：FAILED のPSTを再試行対象に戻す
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import sys
import time
from collections.abc import Iterator
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

import pywintypes
import win32api
import win32event
import win32process

from pstmigration import __version__, config, pipeline, report

log = logging.getLogger("pstmigration")

MUTEX_NAME = r"Global\PSTMigration"
PROCESS_MODE_BACKGROUND_BEGIN = 0x00100000
LOG_KEEP_DAYS = 30


class JsonFormatter(logging.Formatter):
    def format(self, record: logging.LogRecord) -> str:
        d = {
            "ts": datetime.now().astimezone().isoformat(timespec="seconds"),
            "level": record.levelname,
            "logger": record.name,
            "msg": record.getMessage(),
        }
        if record.exc_info:
            d["exc"] = self.formatException(record.exc_info)
        return json.dumps(d, ensure_ascii=False)


def setup_logging(log_dir: Path) -> None:
    log_dir.mkdir(parents=True, exist_ok=True)
    cutoff = time.time() - LOG_KEEP_DAYS * 86400
    for old in log_dir.glob("run-*.jsonl"):
        if old.stat().st_mtime < cutoff:
            old.unlink(missing_ok=True)
    fh = logging.FileHandler(log_dir / f"run-{datetime.now():%Y%m%d-%H%M%S}-{os.getpid()}.jsonl", encoding="utf-8")
    fh.setFormatter(JsonFormatter())
    root = logging.getLogger()
    root.setLevel(logging.INFO)
    root.addHandler(fh)
    if sys.stderr is not None:  # 開発時のみ（--noconsole の EXE では None）
        root.addHandler(logging.StreamHandler())


def lower_priority() -> None:
    """CPU・I/O とも低優先度に（業務を妨げない）。"""
    h = win32api.GetCurrentProcess()
    try:
        win32process.SetPriorityClass(h, PROCESS_MODE_BACKGROUND_BEGIN)
    except pywintypes.error:
        win32process.SetPriorityClass(h, win32process.BELOW_NORMAL_PRIORITY_CLASS)


@contextmanager
def single_instance(wait_ms: int) -> Iterator[bool]:
    """同一PCで1インスタンス（他ユーザーのセッション含む）。"""
    try:
        handle = win32event.CreateMutex(None, False, MUTEX_NAME)
    except pywintypes.error:  # 他ユーザーが作成済み（アクセス拒否）
        yield False
        return
    try:
        r = win32event.WaitForSingleObject(handle, wait_ms)
        ok = r in (win32event.WAIT_OBJECT_0, win32event.WAIT_ABANDONED)
        try:
            yield ok
        finally:
            if ok:
                win32event.ReleaseMutex(handle)
    finally:
        handle.Close()


def default_config_path() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).with_name("config.json")
    return Path(__file__).resolve().parents[2] / "config.json"


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(prog="PSTMigration")
    mode = p.add_mutually_exclusive_group()
    mode.add_argument("--snapshot", action="store_true")
    mode.add_argument("--scan-only", action="store_true")
    mode.add_argument("--aggregate", metavar="ROOT")
    mode.add_argument("--reset", metavar="PST_KEY")
    mode.add_argument("--version", action="store_true")
    p.add_argument("--config", type=Path)
    args = p.parse_args(argv)

    if args.version:
        print(__version__)
        return 0

    home = Path(os.environ.get("LOCALAPPDATA", ".")) / "PSTMigration"
    setup_logging(home / "logs")
    log.info("PSTMigration %s 開始 args=%s", __version__, argv if argv is not None else sys.argv[1:])

    try:
        if args.aggregate:
            if not Path(args.aggregate).is_dir():
                log.error("集計対象フォルダーがありません: %s", args.aggregate)
                return 2
            out = report.aggregate(Path(args.aggregate))
            log.info("集計を出力しました: %s", out)
            return 0

        cfg_path = args.config or default_config_path()
        try:
            cfg = config.load(cfg_path)
        except (OSError, ValueError, TypeError) as e:
            log.error("設定ファイルが不正です（%s）: %s", cfg_path, e)
            return 2

        if not args.snapshot:
            lower_priority()  # ログオフ時はログオフ待ちを短くするため通常優先度

        # ログオフ時はログオン側の終了を少し待つ
        with single_instance(60_000 if args.snapshot else 0) as ok:
            if not ok:
                log.info("別のインスタンスが実行中のため終了します")
                return 0
            if args.reset:
                found = pipeline.reset(home, args.reset)
                log.info("reset %s: %s", args.reset, "OK" if found else "該当なし")
                return 0 if found else 1
            if args.snapshot:
                pipeline.run_snapshot(cfg, home)
            else:
                pipeline.Runner(cfg, home, scan_only=args.scan_only).run()
        return 0
    except Exception:
        log.exception("異常終了")
        return 1
    finally:
        log.info("終了")


if __name__ == "__main__":
    sys.exit(main())
