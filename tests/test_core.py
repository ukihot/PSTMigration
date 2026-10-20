"""Outlook 不要の単体テスト（設計書 §14.2 自動テスト）。"""

import csv
from datetime import UTC, datetime, timedelta, timezone

import pytest

from pstmigration import config, report, rules, transfer
from pstmigration.state import StepError, apply_error, new_entry, write_json

JST = timezone(timedelta(hours=9))
CUTOFF = config.Config().cutoff_utc


def utc(*a, tz=JST):
    return datetime(*a, tzinfo=tz).astimezone(UTC)


# ---- 日付判定（境界・送受信・取得不能） ---------------------------------------


def test_cutoff_is_jst_midnight():
    assert datetime(2020, 12, 31, 15, tzinfo=UTC) == CUTOFF


@pytest.mark.parametrize(
    "delivery, submit, received, expected",
    [
        (utc(2020, 12, 31, 23, 59, 59), None, True, "old"),
        (utc(2021, 1, 1, 0, 0, 0), None, True, "new"),
        # 受信は配信日時優先：送信は年内でも配信が年明けなら新
        (utc(2021, 1, 1, 0, 0, 5), utc(2020, 12, 31, 23, 59, 58), True, "new"),
        # 送信は送信日時優先
        (utc(2021, 1, 1, 0, 0, 5), utc(2020, 12, 31, 23, 59, 58), False, "old"),
        # 優先側が取れなければ代替
        (None, utc(2019, 5, 1), True, "old"),
        # 4501年（OOMの「なし」）は無効扱い
        (datetime(4501, 1, 1, tzinfo=UTC), utc(2022, 3, 3), True, "new"),
        (None, None, True, "unmigrated"),
    ],
)
def test_classify_dates(delivery, submit, received, expected):
    assert rules.classify("IPM.Note", delivery, submit, received, CUTOFF)[0] == expected


def test_classify_non_mail_and_meeting():
    assert rules.classify("IPM.Contact", None, None, True, CUTOFF)[0] == "non_mail"
    assert rules.classify("IPM.Schedule.Meeting.Request", utc(2020, 1, 1), None, True, CUTOFF)[0] == "old"


def test_to_utc_drops_foreign_tzinfo():
    assert rules.to_utc(datetime(2020, 12, 31, 15, 0, 0, tzinfo=JST)) == CUTOFF
    assert rules.to_utc("") is None


# ---- その他の判定 -------------------------------------------------------------


def test_pst_header(tmp_path):
    def make(client, ver):
        p = tmp_path / f"{client}{ver}.pst"
        p.write_bytes(b"!BDN" + b"\0" * 4 + client + ver.to_bytes(2, "little"))
        return p

    assert rules.pst_header(make(b"SM", 23)) == ("PST", "UNICODE")
    assert rules.pst_header(make(b"SM", 14)) == ("PST", "ANSI")
    assert rules.pst_header(make(b"SO", 23))[0] == "OST"
    (tmp_path / "x.pst").write_bytes(b"junk")
    assert rules.pst_header(tmp_path / "x.pst") == (None, None)


def test_output_names_include_computer():
    names = rules.output_names(config.Config().cutoff_local, "PC-0123")
    assert names == (
        "old_mail_before_20210101_PC-0123.pst",
        "old_mail_before_20210101_PC-0123.7z",
        "new_mail_from_20210101_PC-0123.pst",
    )


def test_tool_output_excluded_and_key_stable():
    assert rules.TOOL_OUTPUT_RE.search(r"c:\x\outlook_before_20210101.pst")  # 旧スクリプト
    assert rules.TOOL_OUTPUT_RE.search(r"c:\x\new_mail_from_20210101_pc-0123.pst")
    assert not rules.TOOL_OUTPUT_RE.search(r"c:\x\outlook.pst")
    assert rules.pst_key(r"C:\A\Outlook.pst") == rules.pst_key(r"c:\a\outlook.PST")


def test_totals_and_folder_counts():
    folders = [
        {"path": "\\受信トレイ", "kind": "mail", "old": 3, "new": 2, "unmigrated": 1, "non_mail": 1, "near_cutoff": 0},
        {"path": "\\予定表", "kind": "other", "old": 0, "new": 0, "unmigrated": 0, "non_mail": 9, "near_cutoff": 0},
    ]
    assert rules.totals(folders)["mail"] == 6
    assert rules.folder_counts(folders) == {"\\受信トレイ": 7}


def test_config_validation():
    with pytest.raises(ValueError):
        config.Config(cutoff="2021-01-01T00:00:00")
    with pytest.raises(TypeError):
        config.Config(unknown_key=1)


# ---- 状態遷移 -------------------------------------------------------------------


def test_apply_error_transitions():
    e = new_entry("k", "p", False)
    apply_error(e, StepError("NETWORK", transient=True), 3)
    assert (e["status"], e["attempts"]) == ("RETRY_PENDING", 0)
    apply_error(e, StepError("COUNT_MISMATCH"), 2)
    assert (e["status"], e["attempts"]) == ("RETRY_PENDING", 1)
    apply_error(e, StepError("COUNT_MISMATCH"), 2)
    assert e["status"] == "FAILED"
    e2 = new_entry("k2", "p", False)
    apply_error(e2, StepError("NOT_PST", final=True), 3)
    assert e2["status"] == "FAILED"


def test_classify_outlook_lost_and_file_lock():
    """現場で UNEXPECTED になり回数を消費していたもの（Outlook落ち・共有違反）は一時エラー扱い。"""
    from pstmigration.pipeline import Runner

    class FakeOutlook:
        def __init__(self, up):
            self.up = up

        def alive(self):
            return self.up

    r = object.__new__(Runner)
    e = new_entry("k", "p", False)
    r.ol = FakeOutlook(True)
    assert r._classify(e, PermissionError(13, "Permission denied")).code == "PST_LOCKED"
    assert r._classify(e, StepError("DISK_SPACE", transient=True)).code == "DISK_SPACE"
    assert r._classify(e, ValueError()).code == "UNEXPECTED"
    err = r._classify(e, AttributeError("GetNamespace.RemoveStore"))
    assert err.code == "UNEXPECTED"  # Outlook が生きているなら本当のバグ
    r.ol = FakeOutlook(False)
    err = r._classify(e, AttributeError("GetItemFromID.Move"))
    assert (err.code, err.transient, r.ol) == ("OUTLOOK_LOST", True, None)


# ---- 転送 ------------------------------------------------------------------------


def test_upload_resume_verify_and_no_overwrite(tmp_path):
    src = tmp_path / "a.pst"
    src.write_bytes(b"0123456789" * 100_000)
    sha = transfer.sha256_file(src)
    dst = tmp_path / "remote" / "a.pst"
    dst.parent.mkdir()
    # 途中まで転送済みの .partial から再開
    (dst.parent / "a.pst.partial").write_bytes(src.read_bytes()[:12345])
    transfer.upload(src, dst, sha, max_mbps=0)
    assert transfer.sha256_file(dst) == sha
    assert not (dst.parent / "a.pst.partial").exists()
    # 同一内容なら転送済み扱い
    transfer.upload(src, dst, sha, max_mbps=0)
    # 内容の異なる既存ファイルは上書きしない
    other = tmp_path / "b.pst"
    other.write_bytes(b"different")
    with pytest.raises(StepError) as ex:
        transfer.upload(other, dst, transfer.sha256_file(other), max_mbps=0)
    assert ex.value.code == "REMOTE_CONFLICT"
    assert transfer.sha256_file(dst) == sha


def test_upload_hash_mismatch_discards_partial(tmp_path):
    src = tmp_path / "a.pst"
    src.write_bytes(b"abc")
    dst = tmp_path / "r" / "a.pst"
    with pytest.raises(StepError) as ex:
        transfer.upload(src, dst, "0" * 64, max_mbps=0)
    assert ex.value.code == "HASH_MISMATCH"
    assert not dst.exists() and not (tmp_path / "r" / "a.pst.partial").exists()


def test_slot_without_slots_dir_is_unlimited(tmp_path):
    with transfer.slot(tmp_path, wait_min=0):
        pass


# ---- 集計 ------------------------------------------------------------------------


def test_pick_mailbox():
    def no_ad():
        raise AssertionError("Outlook で決まるなら AD に問い合わせない")

    assert rules.pick_mailbox(["Taro@Example.co.jp", ""], no_ad, "taro") == ("taro@example.co.jp", "OUTLOOK")
    # 複数アカウント → AD
    assert rules.pick_mailbox(["a@x.jp", "info@x.jp"], lambda: "Taro@x.jp", "taro") == ("taro@x.jp", "AD")
    # どちらも取れない → ユーザー名
    assert rules.pick_mailbox([], lambda: None, "taro") == ("taro", "USERNAME")
    assert rules.pick_mailbox(["not-an-address"], lambda: "", "taro") == ("taro", "USERNAME")


def test_aggregate(tmp_path):
    def unit(owner, key, checkpoint="TRANSFERRED", status="COMPLETED"):
        return {
            "key": key,
            "source_path": rf"C:\{key}.pst",
            "status": status,
            "checkpoint": checkpoint,
            "remote_dir": rf"users\{owner}\PC1\{key}_x",
            "counts": {"mail": 3, "old": 2, "new": 1, "unmigrated": 0, "non_mail": 0},
            "outputs": {
                "old_7z": {"name": "old_mail_before_20210101_PC1.7z", "size": 10, "sha256": "aa"},
                "new_pst": {"name": "new_mail_from_20210101_PC1.pst", "size": 20, "sha256": "bb"},
            },
            "warnings": [],
        }

    def status(owner, user, source, psts):
        write_json(
            tmp_path / "users" / owner / "PC1" / "status.json",
            {
                "mailbox": owner,
                "mailbox_source": source,
                "user": user,
                "computer": "PC1",
                "tool_version": "1.0.0",
                "updated_at": "2026-10-14T10:00:00+09:00",
                "psts": psts,
            },
        )

    taro = "taro@example.co.jp"
    status(taro, "taro", "OUTLOOK", [unit(taro, "a"), unit(taro, "b"), unit(taro, "c", None, "PENDING")])
    status("hanako", "hanako", "USERNAME", [unit("hanako", "d")])
    out = report.aggregate(tmp_path)

    def rows(name):
        with open(out / name, encoding="utf-8-sig") as f:
            return list(csv.DictReader(f))

    assert len(rows("ledger.csv")) == 4
    purview = {r["FilePath"]: r for r in rows("purview_mapping_draft.csv")}
    assert set(purview) == {f"users/{taro}/PC1/a_x", f"users/{taro}/PC1/b_x", "users/hanako/PC1/d_x"}
    a = purview[f"users/{taro}/PC1/a_x"]
    assert (a["Mailbox"], a["Note"], a["Name"]) == (taro, "MULTIPLE_UNITS", "new_mail_from_20210101_PC1.pst")
    d = purview["users/hanako/PC1/d_x"]
    assert (d["Mailbox"], d["Note"]) == ("", "MAILBOX_UNKNOWN")
    assert rows("archive_list.csv")[0]["sha256"] == "aa"


def test_aggregate_skips_missing_side(tmp_path):
    """旧0件（7zなし）・新0件（PSTなし）の収集単位は、それぞれの一覧に出さない。"""

    def unit(key, outputs):
        return {
            "key": key,
            "source_path": rf"C:\{key}.pst",
            "status": "COMPLETED",
            "checkpoint": "TRANSFERRED",
            "remote_dir": rf"users\a@x.jp\PC1\{key}_x",
            "counts": {},
            "outputs": outputs,
            "warnings": [],
        }

    new = {"new_pst": {"name": "new.pst", "size": 1, "sha256": "n"}}
    old = {"old_7z": {"name": "old.7z", "size": 1, "sha256": "o"}}
    write_json(
        tmp_path / "users" / "a@x.jp" / "PC1" / "status.json",
        {
            "mailbox": "a@x.jp",
            "user": "a",
            "computer": "PC1",
            "tool_version": "1.0.0",
            "updated_at": "2026-10-15T10:00:00+09:00",
            "psts": [unit("only_new", new), unit("only_old", old)],
        },
    )
    out = report.aggregate(tmp_path)

    def rows(name):
        with open(out / name, encoding="utf-8-sig") as f:
            return list(csv.DictReader(f))

    purview = rows("purview_mapping_draft.csv")
    assert [r["FilePath"] for r in purview] == ["users/a@x.jp/PC1/only_new_x"]
    assert purview[0]["Note"] == ""  # 新PSTは1つだけなので MULTIPLE_UNITS にならない
    assert [r["path"] for r in rows("archive_list.csv")] == [r"users\a@x.jp\PC1\only_old_x\old.7z"]
