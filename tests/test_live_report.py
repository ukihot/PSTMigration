"""ライブ分割のレポートと管理台帳（設計書 §16.6、§16.7、§16.11）。"""

import csv

from pstmigration import chunks, report
from pstmigration.state import write_json


def row(f, k, p, s, h=None):
    return {"f": f, "k": k, "p": p, "s": s, "h": h}


def test_size_deltas_pair_copies_within_each_identity():
    """サイズ差（コピー先 − 元）の分布。合否には使わず記録する（実測 +4 バイト。仕様にはしない）。"""
    expected = [row("受信", "A", "P", 10), row("受信", "A", "P", 20), row("送信", "B", "Q", 7)]
    actual = [row("受信", "A", "P", 24), row("受信", "A", "P", 14), row("送信", "B", "Q", 7)]
    assert chunks.size_deltas(expected, actual, set()) == {"4": 2, "0": 1}


def test_size_deltas_ignore_rows_that_do_not_pair():
    # 照合で不一致になる行（欠落・余計）は差の分布に入れない
    assert chunks.size_deltas([row("受信", "A", "P", 10)], [row("受信", "Z", "P", 99)], set()) == {}


def live_unit(owner, key, *, source_added=None, status="COMPLETED"):
    def c(bucket, seq, rng, size):
        ext = ".pst" if bucket == "new" else ".7z"
        prefix = "new_mail_from_20160221_PC1" if bucket == "new" else "old_mail_before_20160221_PC1"
        return {
            "bucket": bucket,
            "seq": seq,
            "name": chunks.chunk_name(prefix, seq, ext),
            "range": rng,
            "status": "UPLOADED",
            "size": size,
            "sha256": f"{bucket}{seq}",
        }

    return {
        "key": key,
        "source_path": rf"C:\{key}.pst",
        "source_size": 47 * 1024**3,
        "status": status,
        "checkpoint": "TRANSFERRED",
        "split_mode": "live",
        "remote_dir": rf"users\{owner}\PC1\{key}_x",
        "counts": {"mail": 7, "old": 3, "new": 4, "unmigrated": 0, "non_mail": 0},
        "outputs": {},
        "chunks": [c("new", 1, [0, 3], 100), c("new", 2, [3, 4], 50), c("old", 1, [0, 2], 30), c("old", 2, [2, 3], 20)],
        "source_added": source_added,
        "warnings": ["SOURCE_CHANGED"] if source_added else [],
    }


def snapshot_unit(owner, key):
    return {
        "key": key,
        "source_path": rf"C:\{key}.pst",
        "status": "COMPLETED",
        "checkpoint": "TRANSFERRED",
        "remote_dir": rf"users\{owner}\PC1\{key}_x",
        "counts": {"mail": 3, "old": 2, "new": 1, "unmigrated": 0, "non_mail": 0},
        "outputs": {
            "old_7z": {"name": "old_mail_before_20160221_PC1.7z", "size": 10, "sha256": "aa"},
            "new_pst": {"name": "new_mail_from_20160221_PC1.pst", "size": 20, "sha256": "bb"},
        },
        "warnings": [],
    }


def write_status(root, owner, psts):
    write_json(
        root / "users" / owner / "PC1" / "status.json",
        {
            "mailbox": owner,
            "mailbox_source": "OUTLOOK",
            "user": owner.split("@")[0],
            "computer": "PC1",
            "tool_version": "1.2.0",
            "updated_at": "2026-10-22T10:00:00+09:00",
            "psts": psts,
        },
    )


def rows_of(path):
    with open(path, encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def test_aggregate_lists_each_live_chunk_and_recollect_candidates(tmp_path):
    write_status(tmp_path, "aoki@x.jp", [live_unit("aoki@x.jp", "a", source_added={"\\受信トレイ": 2})])
    write_status(tmp_path, "taro@x.jp", [snapshot_unit("taro@x.jp", "t")])
    out = report.aggregate(tmp_path)

    purview = rows_of(out / "purview_mapping_draft.csv")
    aoki = [r for r in purview if r["Mailbox"] == "aoki@x.jp"]
    assert [r["Name"] for r in aoki] == ["new_mail_from_20160221_PC1_001.pst", "new_mail_from_20160221_PC1_002.pst"]
    assert all("MULTIPLE_UNITS" not in r["Note"] for r in aoki)  # 同じ収集単位の束どうしでは付けない
    assert [r["Mailbox"] for r in purview].count("taro@x.jp") == 1

    archives = {r["path"]: r for r in rows_of(out / "archive_list.csv")}
    a1 = archives[r"users\aoki@x.jp\PC1\a_x\old_mail_before_20160221_PC1_001.7z"]
    assert (a1["size"], a1["sha256"], a1["old_mail_count"]) == ("30", "old1", "2")
    assert len(archives) == 3

    ledger = {r["pst_key"]: r for r in rows_of(out / "ledger.csv")}
    assert (ledger["a"]["new_pst_size"], ledger["a"]["archive_size"]) == ("150", "50")
    assert (ledger["a"]["split_mode"], ledger["a"]["chunks"], ledger["a"]["source_added"]) == ("live", "4", "2")
    assert (ledger["t"]["new_pst_size"], ledger["t"]["source_added"]) == ("20", "0")

    recollect = rows_of(out / "recollect_candidates.csv")
    assert [(r["mailbox"], r["pst_key"], r["source_added"]) for r in recollect] == [("aoki@x.jp", "a", "2")]


def test_status_and_report_carry_live_fields():
    e = live_unit("aoki@x.jp", "a", source_added={"\\受信トレイ": 1})
    assert {"split_mode", "chunks", "source_added"} <= set(report.STATUS_KEYS)
    rep = report.build(e, "COMPLETED", {"mailbox": "aoki@x.jp"}, "aoki", "PC1", [])
    assert rep["split_mode"] == "live" and len(rep["chunks"]) == 4 and rep["source_added"] == {"\\受信トレイ": 1}
