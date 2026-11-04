"""Outlook 不要の単体テスト（設計書 §14.2 自動テスト）。"""

import csv
from datetime import UTC, datetime, timedelta, timezone

import pytest

from pstmigration import config, report, rules, transfer
from pstmigration.state import StepError, apply_error, new_entry, write_json

JST = timezone(timedelta(hours=9))
CUTOFF = config.Config().cutoff_utc


def test_write_json_preserves_unpaired_mapi_surrogates(tmp_path):
    from pstmigration.state import read_json

    path = tmp_path / "plan.json"
    data = {"subject": "日本語\ud800件名\udfff", "folder": "受信トレイ📨", "literal": r"\ud800"}
    write_json(path, data)
    assert read_json(path, None) == data
    assert "日本語" in path.read_text(encoding="utf-8")
    assert "\\ud800" in path.read_text(encoding="utf-8")


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
    r.ol_lost = False
    e = new_entry("k", "p", False)
    r.ol = None  # 裏起動の Outlook を自分で終了させた後のエラーは Outlook 落ちではない
    assert r._classify(e, PermissionError(13, "Permission denied")).code == "PST_LOCKED"
    r.ol = FakeOutlook(True)
    assert r._classify(e, PermissionError(13, "Permission denied")).code == "PST_LOCKED"
    assert r._classify(e, StepError("DISK_SPACE", transient=True)).code == "DISK_SPACE"
    assert r._classify(e, ValueError()).code == "UNEXPECTED"
    err = r._classify(e, AttributeError("GetNamespace.RemoveStore"))
    assert err.code == "UNEXPECTED"  # Outlook が生きているなら本当のバグ
    r.ol = FakeOutlook(False)
    err = r._classify(e, AttributeError("GetItemFromID.Move"))
    assert (err.code, err.transient, r.ol, r.ol_lost) == ("OUTLOOK_LOST", True, None, True)


def test_let_go_quits_only_headless_outlook():
    from pstmigration.pipeline import Runner

    class FakeOutlook:
        def __init__(self, headless):
            self.headless = headless

        def quit_if_headless(self):
            return self.headless

    r = object.__new__(Runner)
    r.ol = FakeOutlook(False)  # ユーザーが開いている Outlook はそのまま使い続ける
    r._let_go()
    assert r.ol is not None
    r.ol = FakeOutlook(True)
    r._let_go()
    assert r.ol is None  # 次に要るとき _outlook() で起動し直す


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
    assert ex.value.final and ex.value.details["local_matches_expected"] is False
    assert ex.value.details["stage"] == "transfer_verify"
    assert not dst.exists() and not (tmp_path / "r" / "a.pst.partial").exists()


def test_upload_corrupt_partial_distinguishes_transport_from_local_change(tmp_path):
    src, dst = tmp_path / "source.pst", tmp_path / "nas" / "out.pst"
    src.write_bytes(b"abcdefgh")
    dst.parent.mkdir()
    part = dst.with_name(dst.name + ".partial")
    part.write_bytes(b"bad")
    expected = transfer.sha256_file(src)
    with pytest.raises(StepError) as ex:
        transfer.upload(src, dst, expected, 0)
    err = ex.value
    assert err.code == "HASH_MISMATCH" and not err.final
    assert err.details["local_matches_expected"] is True
    assert err.details["expected_sha256"] == err.details["local_sha256"] == expected
    assert err.details["remote_sha256"] != expected
    assert src.read_bytes() == b"abcdefgh" and not dst.exists() and not part.exists()
    transfer.upload(src, dst, expected, 0)
    assert dst.read_bytes() == src.read_bytes()


def test_slot_without_slots_dir_is_unlimited(tmp_path):
    with transfer.slot(tmp_path, wait_min=0):
        pass


def test_slot_locks_are_released_after_failure(tmp_path):
    slots = tmp_path / "_slots"
    slots.mkdir()
    (slots / "slot1.lock").touch()
    with pytest.raises(RuntimeError), transfer.slot(tmp_path, wait_min=0):
        with pytest.raises(StepError) as ex, transfer.slot(tmp_path, wait_min=-1):
            pytest.fail("使用中の枠を取得できてしまった")
        assert ex.value.code == "SLOT_BUSY"
        raise RuntimeError("upload failed")
    with transfer.slot(tmp_path, wait_min=0):
        pass


def test_slot_network_failure_is_not_reported_as_busy(tmp_path, monkeypatch):
    import pywintypes

    slots = tmp_path / "_slots"
    slots.mkdir()
    (slots / "slot1.lock").touch()

    def disconnected(path):
        raise pywintypes.error(53, "CreateFile", "Network path was not found")

    monkeypatch.setattr(transfer, "open_exclusive", disconnected)
    with pytest.raises(StepError) as ex, transfer.slot(tmp_path, wait_min=0):
        pytest.fail("切断中に転送を始めた")
    assert ex.value.code == "NETWORK" and ex.value.transient


def test_slot_guest_policy_failure_during_directory_listing_is_retryable(tmp_path, monkeypatch):
    from pathlib import Path

    def blocked(path, pattern):
        raise OSError(22, "guest access blocked", str(path), 1272)

    monkeypatch.setattr(Path, "glob", blocked)
    with pytest.raises(StepError) as ex, transfer.slot(tmp_path, wait_min=0):
        pytest.fail("NAS一覧取得失敗後に転送を始めた")
    assert ex.value.code == "NETWORK" and ex.value.transient
    assert ex.value.__cause__.winerror == 1272


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
            "source_size": 20 * 1024**3 + 1024**3 // 2,
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
    assert rows("ledger.csv")[0]["source_gb"] == "20.5"
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


# ---- ログオフ時スナップショット ------------------------------------------------


def test_snapshot_over_budget_warning_is_saved(tmp_path):
    from pstmigration import pipeline
    from pstmigration.state import State

    src = tmp_path / "big.pst"
    src.write_bytes(b"x" * 2048)
    write_json(tmp_path / "state.json", {"psts": {"k": new_entry("k", str(src), False)}})
    cfg = config.Config(logoff_snapshot_budget_gb=1 / 1024**3, work_dir=str(tmp_path / "work"))
    pipeline.run_snapshot(cfg, tmp_path)
    e = State(tmp_path / "state.json").psts["k"]
    assert e["warnings"] == ["SNAPSHOT_OVER_BUDGET"] and e["snapshot_at"] is None


def test_upload_overwrite_replaces_backup(tmp_path):
    src, dst = tmp_path / "contacts.pst", tmp_path / "nas" / "contacts.pst"
    dst.parent.mkdir()
    dst.write_bytes(b"old")
    (dst.parent / "contacts.pst.partial").write_bytes(b"stale")
    src.write_bytes(b"new contacts")
    with pytest.raises(StepError):
        transfer.upload(src, dst, transfer.sha256_file(src), 0)  # 通常は上書きしない
    transfer.upload(src, dst, transfer.sha256_file(src), 0, overwrite=True)
    assert dst.read_bytes() == b"new contacts"


def test_cutoff_change_redoes_and_purges_old_generation(tmp_path):
    from pstmigration import pipeline

    nas, work = tmp_path / "nas", tmp_path / "work"
    old_dir = r"users\a@x.jp\PC1\Outlook_k1_20261013190540"
    (nas / old_dir).mkdir(parents=True)
    (work / "k1").mkdir(parents=True)
    done = new_entry("k1", r"C:\a.pst", False) | {
        "status": "COMPLETED",
        "checkpoint": "TRANSFERRED",
        "remote_dir": old_dir,
    }
    fresh = new_entry("k2", r"C:\b.pst", False)  # まだ何もしていない（1.0.1 の state なので cutoff なし）
    write_json(tmp_path / "state.json", {"psts": {"k1": done, "k2": fresh}})

    r = pipeline.Runner(
        config.Config(cutoff="2016-01-01T00:00:00+09:00", server_root=str(nas), work_dir=str(work)), tmp_path
    )
    r._apply_cutoff()
    k1, k2 = r.st.psts["k1"], r.st.psts["k2"]
    assert (k1["status"], k1["checkpoint"], k1["remote_dir"], k1["generation"]) == ("PENDING", None, None, 2)
    assert k1["stale_remote_dirs"] == [old_dir] and not (work / "k1").exists()
    assert (nas / old_dir).exists()  # 新世代が揃うまでは消さない
    assert (k2["generation"], k2["cutoff"]) == (1, "2016-01-01T00:00:00+09:00")

    r._apply_cutoff()  # 同じ cutoff ならもう何もしない
    assert k1["generation"] == 2

    r._purge_stale(k1)
    assert not (nas / old_dir).exists() and k1["stale_remote_dirs"] == []


# ---- 分割（MAPI 一括移動と、失敗時の1件ずつ移動） ---------------------------------


def test_split_moves_in_batches_and_falls_back(monkeypatch):
    import pywintypes

    from pstmigration import outlook

    class Folders:
        def __init__(self):
            self.d = {}

        def Item(self, name):
            if name not in self.d:
                raise pywintypes.com_error(-1, "not found", None, None)
            return self.d[name]

        def Add(self, name, _type):
            self.d[name] = Folder(name)
            return self.d[name]

    class Folder:
        def __init__(self, name):
            self.Name, self.EntryID, self.Folders = name, f"F-{name}-{id(self)}", Folders()

    class Store:
        def __init__(self, sid):
            self.StoreID, self.root = sid, Folder("root")

        def GetRootFolder(self):
            return self.root

    snap, old, new = Store("S"), Store("O"), Store("N")
    snap.root.Folders.Add("受信トレイ", 0)
    ids = {("受信トレイ",): {"old": [f"o{i}" for i in range(1200)], "new": ["n1", "n2"]}}

    class Mover:
        def __init__(self, fail_after):
            self.calls, self.fail_after, self.closed = [], fail_after, False

        def move(self, src_store, src_folder, dst_store, dst_folder, eids):
            if len(self.calls) == self.fail_after:
                raise pywintypes.com_error(-1, "MAPI error", None, None)
            self.calls.append((dst_store, len(eids)))

        def close(self):
            self.closed = True

    moved_each = []

    class NS:
        def GetItemFromID(self, eid, store_id):
            class Item:
                def Move(self, dest):
                    moved_each.append(eid)

            return Item()

    ol = object.__new__(outlook.Outlook)
    ol.ns = NS()

    mover = Mover(fail_after=None)
    monkeypatch.setattr(ol, "_mapi_mover", lambda: mover)
    assert ol.split(snap, old, new, None, 0, ids) == 1202
    assert mover.calls == [("O", 500), ("O", 500), ("O", 200), ("N", 2)] and mover.closed
    assert moved_each == []
    assert old.root.Folders.Item("受信トレイ") and new.root.Folders.Item("受信トレイ")  # 両方に同じ階層

    mover = Mover(fail_after=1)  # 2回目の一括移動で失敗 → 残りは1件ずつ
    monkeypatch.setattr(ol, "_mapi_mover", lambda: mover)
    assert ol.split(snap, old, new, None, 0, ids) == 1202
    assert mover.calls == [("O", 500)] and mover.closed
    assert len(moved_each) == 700 + 2


def test_rows_reads_table_in_row_major_batches():
    """Table.GetArray は a[行][列] で返る。転置せずに1行ずつ返すこと。"""
    from types import SimpleNamespace

    from pstmigration import outlook

    batches = iter([(("id1", "s1"), ("id2", "s2")), (("id3", "s3"),)])
    table = SimpleNamespace(EndOfTable=False, Columns=SimpleNamespace(RemoveAll=lambda: None, Add=lambda c: None))

    def get_array(n):
        assert n == outlook.TABLE_BATCH
        rows = next(batches, ())
        table.EndOfTable = rows == () or len(rows) < 2
        return rows

    table.GetArray = get_array
    folder = SimpleNamespace(GetTable=lambda *a: table)
    assert list(outlook._rows(folder)) == [("id1", "s1"), ("id2", "s2"), ("id3", "s3")]


@pytest.mark.parametrize(
    "started_here, windows, quits",
    [
        (True, 0, True),  # 本ツールが裏で起動し、画面も無い → 終了させる
        (True, 1, False),  # 本ツールが起動したが、ユーザーが画面を開いた → 終了させない
        (False, 0, False),  # ユーザーが起動した Outlook（画面が見えていなくても）→ 終了させない
        (False, 1, False),
    ],
)
def test_quit_only_outlook_started_by_this_tool(started_here, windows, quits):
    from types import SimpleNamespace

    from pstmigration import outlook

    called = []
    app = SimpleNamespace(
        Explorers=SimpleNamespace(Count=windows), Inspectors=SimpleNamespace(Count=0), Quit=lambda: called.append(1)
    )
    ol = object.__new__(outlook.Outlook)
    ol.app, ol.started_here = app, started_here
    assert ol.quit_if_headless() is quits
    assert bool(called) is quits


def test_code_defaults_contain_no_credentials():
    """NAS の資格情報は config.json で渡す。コードの既定値は空。"""
    c = config.Config()
    assert c.server_user == "" and c.server_password == ""
