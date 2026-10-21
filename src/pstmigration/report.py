"""report.json / status.json / 管理者向け集計（設計書 §11）。"""

from __future__ import annotations

import csv
import json
import logging
from collections import Counter
from pathlib import Path, PureWindowsPath

from pstmigration import __version__
from pstmigration.state import now_iso, write_json

log = logging.getLogger(__name__)

STATUS_KEYS = (
    "key",
    "source_path",
    "source_size",
    "is_default_store",
    "status",
    "checkpoint",
    "generation",
    "cutoff",
    "snapshot_at",
    "remote_dir",
    "counts",
    "outputs",
    "warnings",
    "attempts",
    "last_error",
)

PURVIEW_TARGET_ROOT = "/POP移行メール"


def build(entry: dict, status: str, data: dict, user: str, computer: str, folders: list[dict]) -> dict:
    counts = dict(entry.get("counts") or {})
    counts["moved_old"] = counts.get("old")
    counts["moved_new"] = counts.get("new")
    errors = list(entry.get("problems") or [])
    if entry.get("last_error"):
        errors.append(entry["last_error"])
    return {
        "tool_version": __version__,
        "mailbox": data.get("mailbox"),
        "mailbox_source": data.get("mailbox_source"),
        "user": user,
        "computer": computer,
        "status": status,
        "started_at": entry.get("snapshot_at"),
        "finished_at": now_iso(),
        "source": {
            "path": entry["source_path"],
            "size": entry.get("source_size"),
            "is_default_store": entry.get("is_default_store"),
            "pst_format": entry.get("pst_format"),
        },
        "counts": counts,
        "archive": entry.get("archive"),
        "outputs": entry.get("outputs"),
        "remote_dir": entry.get("remote_dir"),
        "durations": entry.get("durations"),
        "warnings": entry.get("warnings"),
        "errors": errors,
        "folders": folders,
    }


def publish_status(server_root: Path, data: dict, user: str, computer: str) -> None:
    """このPC・ユーザーの状態をサーバへ（失敗しても処理は続ける）。"""
    owner = data.get("mailbox") or user
    status = {
        "mailbox": data.get("mailbox"),
        "mailbox_source": data.get("mailbox_source"),
        "user": user,
        "computer": computer,
        "tool_version": __version__,
        "updated_at": now_iso(),
        "last_run_error": data.get("last_run_error"),
        "contacts": data.get("contacts"),
        "psts": [{k: e.get(k) for k in STATUS_KEYS} for e in data["psts"].values()],
    }
    try:
        write_json(server_root / "users" / owner / computer / "status.json", status)
    except OSError as e:
        log.warning("status.json をサーバへ出力できません: %s", e)


def _write_csv(path: Path, header: list[str], rows: list[list]) -> None:
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


def aggregate(root: Path) -> Path:
    """全 status.json を集約して _admin\\ に台帳と後続作業用CSVを出す（設計書 §11.3、§11.4）。"""
    statuses = [json.loads(p.read_text(encoding="utf-8")) for p in sorted(root.glob("users/*/*/status.json"))]

    ledger, uploaded = [], []
    for s in statuses:
        mailbox = s.get("mailbox") or ""
        for p in s["psts"]:
            c, o, err = p.get("counts") or {}, p.get("outputs") or {}, p.get("last_error") or {}
            ledger.append(
                [
                    mailbox,
                    s.get("mailbox_source"),
                    s["user"],
                    s["computer"],
                    p["key"],
                    p["source_path"],
                    round(p["source_size"] / 1024**3, 2) if p.get("source_size") else None,
                    p.get("is_default_store"),
                    p["status"],
                    p.get("checkpoint"),
                    err.get("code"),
                    c.get("mail"),
                    c.get("old"),
                    c.get("new"),
                    c.get("unmigrated"),
                    c.get("non_mail"),
                    (o.get("new_pst") or {}).get("size"),
                    (o.get("old_7z") or {}).get("size"),
                    p.get("remote_dir"),
                    "|".join(p.get("warnings") or []),
                    s["updated_at"],
                    s["tool_version"],
                ]
            )
            if p.get("checkpoint") == "TRANSFERRED":
                uploaded.append((mailbox or s["user"], s["computer"], p))

    # 0件の側はファイルが無い（outputs にキーが無い）ので、ある分だけ一覧にする
    with_new = [u for u in uploaded if "new_pst" in u[2]["outputs"]]
    per_owner = Counter(owner for owner, _, _ in with_new)
    purview, archives = [], []
    for owner, _, p in with_new:
        remote = PureWindowsPath(p["remote_dir"])
        notes = []
        if per_owner[owner] > 1:
            notes.append("MULTIPLE_UNITS")  # 1人に複数PST（複数PC・複数PST・再収集）
        if "@" not in owner:
            notes.append("MAILBOX_UNKNOWN")  # メールアドレスが取れなかった。手で記入
        purview.append(
            [
                "Exchange",
                remote.as_posix(),
                p["outputs"]["new_pst"]["name"],
                owner if "@" in owner else "",
                "FALSE",
                PURVIEW_TARGET_ROOT,
                "",
                "",
                "",
                "",
                "|".join(notes),
            ]
        )
    for owner, computer, p in uploaded:
        o, c = p["outputs"], p.get("counts") or {}
        if "old_7z" not in o:
            continue
        remote = PureWindowsPath(p["remote_dir"])
        archives.append(
            [
                owner,
                computer,
                str(remote / o["old_7z"]["name"]),
                o["old_7z"]["size"],
                o["old_7z"]["sha256"],
                c.get("old"),
                p["status"],
            ]
        )

    out = root / "_admin"
    out.mkdir(exist_ok=True)
    _write_csv(
        out / "ledger.csv",
        [
            "mailbox",
            "mailbox_source",
            "user",
            "computer",
            "pst_key",
            "source_path",
            "source_gb",
            "is_default_store",
            "status",
            "checkpoint",
            "last_error_code",
            "mail",
            "old",
            "new",
            "unmigrated",
            "non_mail",
            "new_pst_size",
            "archive_size",
            "remote_dir",
            "warnings",
            "updated_at",
            "tool_version",
        ],
        ledger,
    )
    _write_csv(
        out / "purview_mapping_draft.csv",
        [
            "Workload",
            "FilePath",
            "Name",
            "Mailbox",
            "IsArchive",
            "TargetRootFolder",
            "ContentCodePage",
            "SPFileContainer",
            "SPManifestContainer",
            "SPSiteUrl",
            "Note",
        ],
        purview,
    )
    _write_csv(
        out / "archive_list.csv",
        [
            "mailbox",
            "computer",
            "path",
            "size",
            "sha256",
            "old_mail_count",
            "status",
        ],
        archives,
    )
    return out
