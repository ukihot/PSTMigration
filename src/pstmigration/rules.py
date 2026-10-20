"""Outlook に依存しない判定ロジック（単体テスト対象）。"""

from __future__ import annotations

import hashlib
import os
import re
from collections.abc import Callable
from datetime import UTC, datetime, timedelta
from pathlib import Path
from typing import NamedTuple

# メールフォルダー内でもメールとして扱わないクラス（設計書 §5.1）
NON_MAIL_PREFIXES = (
    "IPM.APPOINTMENT",
    "IPM.CONTACT",
    "IPM.DISTLIST",
    "IPM.TASK",
    "IPM.STICKYNOTE",
    "IPM.ACTIVITY",
)

# 旧 Split-PstByDate.ps1（*_before_YYYYMMDD.pst）と本ツール（*_before_YYYYMMDD_{PC名}.pst）の出力PST
TOOL_OUTPUT_RE = re.compile(r"_(before|from)_\d{8}(_[^\\/]+)?\.pst$", re.IGNORECASE)

NEAR_CUTOFF = timedelta(days=1)

# フォルダー名に使えない文字・空白を含まないメールアドレス
EMAIL_RE = re.compile(r'^[^@\s\\/:*?"<>|]+@[^@\s\\/:*?"<>|]+\.[^@\s\\/:*?"<>|]+$')


class Names(NamedTuple):
    old_pst: str
    old_7z: str
    new_pst: str


def output_names(cutoff_local: datetime, computer: str) -> Names:
    """同じアカウントで複数PCを使う人がいるため、フォルダーから出してもPCが分かるようPC名を付ける。"""
    d = cutoff_local.strftime("%Y%m%d")
    return Names(
        f"old_mail_before_{d}_{computer}.pst",
        f"old_mail_before_{d}_{computer}.7z",
        f"new_mail_from_{d}_{computer}.pst",
    )


def pick_mailbox(outlook_addrs: list[str], ad_mail: Callable[[], str | None], username: str) -> tuple[str, str]:
    """(メールアドレス or ユーザー名, 'OUTLOOK'|'AD'|'USERNAME')。

    Outlook のアカウントが1つに定まればそれ（実際に使っているメールボックス）。
    複数・なしなら AD の mail 属性（DC への問い合わせは必要なときだけ）。どちらもだめならユーザー名。
    """
    addrs = {a.strip().lower() for a in outlook_addrs if a and EMAIL_RE.match(a.strip())}
    if len(addrs) == 1:
        return addrs.pop(), "OUTLOOK"
    mail = (ad_mail() or "").strip().lower()
    if EMAIL_RE.match(mail):
        return mail, "AD"
    return username, "USERNAME"


def norm_path(path: str) -> str:
    return os.path.normcase(os.path.abspath(path))


def pst_key(path: str) -> str:
    return hashlib.sha256(norm_path(path).encode("utf-8")).hexdigest()[:12]


def to_utc(value: object) -> datetime | None:
    """COMの日時をUTCに。

    MAPI proptag で取得した日時はUTCの壁時計値。pywin32 が付ける tzinfo は当てにせず付け直す。
    """
    if not isinstance(value, datetime):
        return None
    return datetime(value.year, value.month, value.day, value.hour, value.minute, value.second, tzinfo=UTC)


def _valid(d: datetime | None) -> bool:
    # 1601年（ゼロ値）や 4501年（OOM の「なし」）を除外
    return d is not None and 1980 <= d.year <= 2100


def pick_date(delivery: datetime | None, submit: datetime | None, received: bool) -> datetime | None:
    """受信: 配信日時(ReceivedTime)→送信日時(SentOn)。送信・下書き: 逆順。"""
    for d in (delivery, submit) if received else (submit, delivery):
        if _valid(d):
            return d
    return None


def classify(
    message_class: str | None,
    delivery: datetime | None,
    submit: datetime | None,
    received: bool,
    cutoff_utc: datetime,
) -> tuple[str, datetime | None]:
    """'old' / 'new' / 'unmigrated'（日付取得不能）/ 'non_mail' と判定に使った日時。"""
    if (message_class or "").upper().startswith(NON_MAIL_PREFIXES):
        return "non_mail", None
    d = pick_date(delivery, submit, received)
    if d is None:
        return "unmigrated", None
    return ("old" if d < cutoff_utc else "new"), d


def is_near_cutoff(d: datetime, cutoff_utc: datetime) -> bool:
    return abs(d - cutoff_utc) <= NEAR_CUTOFF


def pst_header(path: Path) -> tuple[str | None, str | None]:
    """('PST'|'OST'|None, 'ANSI'|'UNICODE'|None)。"""
    with open(path, "rb") as f:
        head = f.read(12)
    if len(head) < 12 or head[:4] != b"!BDN":
        return None, None
    kind = {b"SM": "PST", b"SO": "OST"}.get(head[8:10])
    ver = int.from_bytes(head[10:12], "little")
    fmt = "ANSI" if ver in (14, 15) else "UNICODE" if ver >= 23 else None
    return kind, fmt


def disk_needed(size: int, factor: float, reserve_gb: float) -> int:
    return int(size * factor + reserve_gb * 1024**3)


def totals(folders: list[dict]) -> dict:
    t = {k: sum(f[k] for f in folders) for k in ("old", "new", "unmigrated", "non_mail", "near_cutoff")}
    t["mail"] = t["old"] + t["new"] + t["unmigrated"]
    return t


def folder_counts(folders: list[dict]) -> dict[str, int]:
    """完了後の変更検出用：メールフォルダーごとの全アイテム数。"""
    return {f["path"]: f["old"] + f["new"] + f["unmigrated"] + f["non_mail"] for f in folders if f["kind"] == "mail"}
