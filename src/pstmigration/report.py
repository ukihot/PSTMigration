"""report.json / status.json / 管理者向け集計（設計書 §11）。"""

from __future__ import annotations

import csv
import json
import logging
import ntpath
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
    "split_mode",  # ライブ分割（§16）
    "chunks",
    "source_added",
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
        "split_mode": entry.get("split_mode"),
        "chunks": entry.get("chunks"),  # ライブ分割の束（名前・範囲・サイズ・SHA-256・サイズ差の分布）
        "source_added": entry.get("source_added"),  # 調査後に増えたメール（移行対象外。再回収の候補）
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


def _files(p: dict) -> list[tuple[str, dict, int | None]]:
    """収集単位の出力ファイル: (new_pst か old_7z, {name, size, sha256}, 旧メールの件数)。

    ライブ分割は束ごと（§16）、現行方式は1つずつ。0件の側はファイルが無い。
    照合済み（VERIFIED）でまだ圧縮・ハッシュ前の束は size が無いので数えない。
    """
    if p.get("chunks"):
        built = (c for c in p["chunks"] if "size" in c)
        ordered = sorted(built, key=lambda c: (c["bucket"] != "new", c["seq"]))
        return [("new_pst" if c["bucket"] == "new" else "old_7z", c, c["range"][1] - c["range"][0]) for c in ordered]
    o, c = p.get("outputs") or {}, p.get("counts") or {}
    return [(k, o[k], c.get("old") if k == "old_7z" else None) for k in ("new_pst", "old_7z") if k in o]


def _write_csv(path: Path, header: list[str], rows: list[list]) -> None:
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


def aggregate(root: Path) -> Path:
    """全 status.json を集約して _admin\\ に台帳と後続作業用CSVを出す（設計書 §11.3、§11.4）。"""
    statuses = [(p, json.loads(p.read_text(encoding="utf-8"))) for p in sorted(root.glob("users/*/*/status.json"))]
    identities = Counter(
        (s["computer"].casefold(), ntpath.normcase(ntpath.normpath(p["source_path"])))
        for _, s in statuses
        for p in s["psts"]
    )

    ledger, uploaded, recollect, errors = [], [], [], []
    for source, s in statuses:
        mailbox = s.get("mailbox") or ""
        source_status = source.relative_to(root).as_posix()
        for p in s["psts"]:
            c, err, files = p.get("counts") or {}, p.get("last_error") or {}, _files(p)
            added = sum((p.get("source_added") or {}).values())
            recollect_required = bool(added or "SOURCE_CHANGED" in (p.get("warnings") or []))
            duplicate = identities[(s["computer"].casefold(), ntpath.normcase(ntpath.normpath(p["source_path"])))] > 1
            ledger.append(
                [
                    mailbox,
                    s.get("mailbox_source"),
                    s["user"],
                    s["computer"],
                    p["key"],
                    p["source_path"],
                    round(p["source_size"] / 1024**3, 2) if p.get("source_size") is not None else None,
                    p.get("is_default_store"),
                    p["status"],
                    p.get("checkpoint"),
                    err.get("code"),
                    c.get("mail"),
                    c.get("old"),
                    c.get("new"),
                    c.get("unmigrated"),
                    c.get("non_mail"),
                    sum(f["size"] for k, f, _ in files if k == "new_pst") if files else None,
                    sum(f["size"] for k, f, _ in files if k == "old_7z") if files else None,
                    p.get("remote_dir"),
                    "|".join(p.get("warnings") or []),
                    s["updated_at"],
                    s["tool_version"],
                    p.get("split_mode"),
                    len(p.get("chunks") or []),
                    added,
                    p.get("generation"),
                    source_status,
                    duplicate,
                    recollect_required,
                    err.get("message"),
                    err.get("phase"),
                    err.get("checkpoint"),
                    err.get("exception_type"),
                    err.get("retryable"),
                    err.get("tool_version"),
                ]
            )
            if err:
                errors.append(
                    {
                        "mailbox": mailbox,
                        "computer": s["computer"],
                        "pst_key": p["key"],
                        "source_status": source_status,
                        "generation": p.get("generation"),
                        "status": p["status"],
                        "checkpoint": p.get("checkpoint"),
                        "error": err,
                    }
                )
            if p.get("checkpoint") == "TRANSFERRED":
                uploaded.append((mailbox or s["user"], s["computer"], p))
            if recollect_required:
                warnings = "|".join(p.get("warnings") or [])
                recollect.append([mailbox, s["computer"], p["key"], p["source_path"], p["status"], added, warnings])

    # 0件の側はファイルが無いので、ある分だけ一覧にする。ライブ分割は束ごとに1行
    with_new = [u for u in uploaded if any(k == "new_pst" for k, _, _ in _files(u[2]))]
    per_owner = Counter(owner for owner, _, _ in with_new)  # 収集単位で数える（束の数ではない）
    purview, archives = [], []
    for owner, _, p in with_new:
        remote = PureWindowsPath(p["remote_dir"])
        notes = []
        if per_owner[owner] > 1:
            notes.append("MULTIPLE_UNITS")  # 1人に複数PST（複数PC・複数PST・再収集）
        if "@" not in owner:
            notes.append("MAILBOX_UNKNOWN")  # メールアドレスが取れなかった。手で記入
        for kind, f, _ in _files(p):
            if kind != "new_pst":
                continue
            purview.append(
                [
                    "Exchange",
                    remote.as_posix(),
                    f["name"],
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
        remote = PureWindowsPath(p["remote_dir"])
        for kind, f, old_count in _files(p):
            if kind == "old_7z":
                archives.append(
                    [owner, computer, str(remote / f["name"]), f["size"], f["sha256"], old_count, p["status"]]
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
            "split_mode",
            "chunks",
            "source_added",
            "generation",
            "source_status",
            "duplicate_source",
            "recollect_required",
            "last_error_message",
            "last_error_phase",
            "last_error_checkpoint",
            "last_error_exception_type",
            "last_error_retryable",
            "last_error_tool_version",
        ],
        ledger,
    )
    write_json(out / "errors.json", errors)
    _write_csv(
        out / "recollect_candidates.csv",  # 調査後に増えたメールがある収集単位（移行直前の再回収の候補。§16.7）
        ["mailbox", "computer", "pst_key", "source_path", "status", "source_added", "warnings"],
        recollect,
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
