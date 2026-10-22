"""ライブ分割方式の pipeline への組み込み（設計書 §16.3、§16.7〜§16.9）。

MAPI（mapilive）と 7z だけ偽物にする。NAS はローカルフォルダーで、転送は本物の transfer を使う。
"""

import contextlib
import json
import os
from collections import Counter
from pathlib import Path

import pytest

from pstmigration import chunks, config, mapilive, pipeline, rules
from pstmigration.state import StepError, new_entry, read_json

GB = 1024**3


class Crash(BaseException):
    pass


class FakeSource:
    """元PST。メールは {e, f, k, p, s, bucket}。テスト中に削除・追加できる。"""

    def __init__(self, n_new, n_old):
        self.mails = [self.mail(f"n{i}", "new") for i in range(n_new)] + [
            self.mail(f"o{i}", "old") for i in range(n_old)
        ]
        self.scans = 0
        self.on_copy = None  # コピーのたびに呼ぶ（処理中の変更を模す）

    @staticmethod
    def mail(e, bucket, f="\\受信トレイ"):
        return {"e": e, "f": f, "k": f"K{e}", "p": f"P{e}", "s": 300, "bucket": bucket}

    def scan(self, store, cutoff, log_subject=False):
        self.scans += 1
        row = lambda m: {k: m[k] for k in ("e", "f", "k", "p", "s")}  # noqa: E731
        items = {b: [row(m) for m in self.mails if m["bucket"] == b] for b in ("new", "old")}
        rec = {
            "path": "\\受信トレイ",
            "kind": "mail",
            "old": len(items["old"]),
            "new": len(items["new"]),
            "unmigrated": sum(m["bucket"] == "unmigrated" for m in self.mails),
            "non_mail": 0,
            "near_cutoff": 0,
            "old_bytes": 0,
            "new_bytes": 0,
        }
        return {
            "scan": {"folders": [rec], "problems": []},
            "items": items,
            "all_rows": [row(m) for m in self.mails],
            "mail_folders": [("\\受信トレイ", "IPF.Note")],
        }


class FakeWriter:
    def __init__(self, src, pst, source, folders, pause_ms=0):
        assert not pst.exists()
        self.src, self.pst = src, pst
        pst.write_text(json.dumps({"nonce": os.urandom(4).hex(), "rows": []}))

    def copy(self, items):
        if self.src.on_copy:
            self.src.on_copy()
        d = json.loads(self.pst.read_text())
        d["rows"] += [{**i, "pad": "x" * i["s"]} for i in items]  # 本物の PST と同じくメールのサイズ分増える
        self.pst.write_text(json.dumps(d))

    def close(self):
        pass


def fake_7z(args, cwd, code):
    if args[0] == "a":
        (cwd / args[-2]).write_bytes(b"7z:" + (cwd / args[-1]).read_bytes())


@pytest.fixture
def env(tmp_path, monkeypatch):
    monkeypatch.setenv("COMPUTERNAME", "PC1")
    src = FakeSource(5, 4)
    monkeypatch.setattr(mapilive, "Mapi", contextlib.nullcontext)
    monkeypatch.setattr(
        mapilive, "logon", lambda profile: type("S", (), {"OpenMsgStore": lambda *a: None, "Logoff": lambda *a: None})()
    )
    monkeypatch.setattr(mapilive, "open_user_store", lambda session, store_id: object())
    monkeypatch.setattr(mapilive, "scan", src.scan)
    monkeypatch.setattr(mapilive, "add_hashes", lambda store, items, need: None)
    monkeypatch.setattr(
        mapilive, "ChunkWriter", lambda pst, source, folders, pause_ms=0: FakeWriter(src, pst, source, folders)
    )
    monkeypatch.setattr(
        mapilive, "read_chunk", lambda pst, need: {"rows": json.loads(pst.read_text())["rows"], "folders": []}
    )
    monkeypatch.setattr(pipeline, "_run_7z", fake_7z)
    cfg = config.Config(
        cutoff="2016-02-21T00:00:00+09:00",
        server_root=str(tmp_path / "nas"),
        server_user="",
        work_dir=str(tmp_path / "work"),
        disk_reserve_gb=0,
        chunk_max_gb=1000 / GB,  # 束の目標 1000 バイト（メール3通ほど）
        chunk_min_gb=10 / GB,
        batch_target_mb=600 / 1024**2,
        split_pause_ms=0,
    )
    (tmp_path / "nas").mkdir()
    r = pipeline.Runner(cfg, tmp_path / "home")
    r.profile = "Outlook"
    e = new_entry("key1", str(tmp_path / "big.pst"), False)
    e.update(split_mode="live", store_id="00", source_size=47 * GB, cutoff=cfg.cutoff)
    r.st.psts["key1"] = e
    r.st.data["mailbox"] = "a@x.jp"
    return r, e, src, tmp_path


def nas_files(tmp_path):
    return sorted(p.name for p in (tmp_path / "nas").rglob("*") if p.is_file())


def test_live_split_completes_with_chunks_on_nas_and_report(env):
    r, e, src, tmp = env
    r._process_live(e)
    assert e["status"] == "COMPLETED", e.get("last_error")
    assert chunks.completion_problems(e["chunks"], {"new": 5, "old": 4}) == []
    names = nas_files(tmp)
    assert sum(n.startswith("new_mail_from_20160221_PC1_") and n.endswith(".pst") for n in names) >= 2
    assert sum(n.startswith("old_mail_before_20160221_PC1_") and n.endswith(".7z") for n in names) >= 2
    assert "report.json" in names
    assert not (r.cfg.work_path / "key1").exists()  # 完了後はローカルを片付ける
    assert e["counts"]["new"] == 5 and e["counts"]["old"] == 4


def test_mail_deleted_during_copy_fails_and_keeps_uploaded_chunks(env):
    r, e, src, tmp = env
    copies = Counter()

    def delete_planned_mail_on_second_copy():
        copies["n"] += 1
        if copies["n"] == 2:
            src.mails = [m for m in src.mails if m["e"] != "o3"]  # まだコピーしていない旧メールが消えた

    src.on_copy = delete_planned_mail_on_second_copy
    r._process_live(e)
    assert e["status"] == "FAILED" and e["last_error"]["code"] == "SOURCE_CHANGED_DURING_COPY"
    assert any(c["status"] == "UPLOADED" for c in e["chunks"])  # 送り終えた束の記録は残す


def test_mail_added_after_scan_is_reported_not_failed(env):
    r, e, src, tmp = env

    def add_once():
        if not any(m["e"] == "late" for m in src.mails):
            src.mails.append(FakeSource.mail("late", "new"))

    src.on_copy = add_once
    r._process_live(e)
    assert e["status"] == "COMPLETED"
    assert e["source_added"] == {"\\受信トレイ": 1} and "SOURCE_CHANGED" in e["warnings"]


def test_interrupted_live_split_resumes_from_saved_plan_without_duplicates(env):
    r, e, src, tmp = env
    copies = Counter()

    def crash_on_fourth_copy():
        copies["n"] += 1
        if copies["n"] == 4:
            raise Crash

    src.on_copy = crash_on_fourth_copy
    with pytest.raises(Crash):
        r._process_live(e)
    src.on_copy = None
    r2 = pipeline.Runner(r.cfg, tmp / "home")  # プロセスを作り直す（state.json から）
    r2.profile = "Outlook"
    e2 = r2.st.psts["key1"]
    r2._process_live(e2)
    assert e2["status"] == "COMPLETED"
    assert src.scans >= 2 and read_json(tmp / "home" / "state.json", {})  # 調査は計画作成の1回＋再調査
    assert len([n for n in nas_files(tmp) if n.endswith((".pst", ".7z"))]) == len(e2["chunks"])


def test_snapshot_at_logoff_skips_live_psts(env, monkeypatch):
    r, e, src, tmp = env
    Path(e["source_path"]).write_bytes(b"x")
    r.st.save()
    monkeypatch.setattr(pipeline, "copy_snapshot", lambda *a: pytest.fail("ライブ分割のPSTはログオフ時にコピーしない"))
    pipeline.run_snapshot(r.cfg, tmp / "home")


@pytest.mark.parametrize("mapi_ok, expected", [(True, "live"), (False, None)])
def test_mode_is_decided_once_before_any_progress(env, monkeypatch, mapi_ok, expected):
    r, e, src, tmp = env
    e.pop("split_mode")
    monkeypatch.setattr(r, "_mapi_ok", lambda: mapi_ok)
    assert r._wants_live(e) is (expected == "live")
    assert e.get("split_mode") == expected
    if not mapi_ok:
        assert e["status"] == "RETRY_PENDING" and e["last_error"]["code"] == "LIVE_SPLIT_UNAVAILABLE"


def test_snapshot_psts_in_progress_stay_on_snapshot_method(env, monkeypatch):
    r, e, src, tmp = env
    e.update(split_mode=None, snapshot_at="2026-10-22T00:00:00+09:00")
    monkeypatch.setattr(r, "_mapi_ok", lambda: pytest.fail("判定しない"))
    assert r._wants_live(e) is False


def test_reset_of_failed_live_pst_starts_new_generation(env):
    r, e, src, tmp = env
    e.update(status="FAILED", remote_dir="users\\a@x.jp\\PC1\\big_key1_1", chunks=[{"bucket": "new", "seq": 1}])
    r.st.save()
    assert pipeline.reset(tmp / "home", "key1")
    e2 = read_json(tmp / "home" / "state.json", {})["psts"]["key1"]
    assert (e2["status"], e2["chunks"], e2["remote_dir"], e2["generation"]) == ("PENDING", [], None, 2)
    assert e2["stale_remote_dirs"] == ["users\\a@x.jp\\PC1\\big_key1_1"]


def test_split_mode_rule_is_used_for_decision():
    assert (
        rules.split_mode("auto", 47 * GB, 200 * GB, True, budget_bytes=10 * GB, disk_factor=2.5, reserve_bytes=5 * GB)
        == "live"
    )


def test_step_error_is_not_swallowed_as_unexpected(env):
    r, e, src, tmp = env

    def boom():
        raise StepError("DISK_SPACE", "x", transient=True)

    src.on_copy = boom
    r._process_live(e)
    assert e["status"] == "RETRY_PENDING" and e["last_error"]["code"] == "DISK_SPACE"


def test_scan_only_never_runs_live_split(env, monkeypatch):
    """--scan-only（パイロットの調査）で束の作成・転送が走ってはいけない。"""
    r, e, src, tmp = env
    r.scan_only = True
    monkeypatch.setattr(r, "_process_live", lambda e: pytest.fail("調査のみなのにライブ分割した"))
    r._process_entry(e)


def test_entry_dispatch(env, monkeypatch):
    r, e, src, tmp = env
    called = []
    monkeypatch.setattr(r, "_process_live", lambda e: called.append("live"))
    monkeypatch.setattr(r, "_process", lambda e: called.append("snapshot"))
    r._process_entry(e)  # split_mode=live
    e.update(split_mode="snapshot", snapshot_at="2026-10-22T00:00:00+09:00")
    r._process_entry(e)
    e.update(status="COMPLETED")
    r._process_entry(e)  # 完了済みは何もしない
    assert called == ["live", "snapshot"]


def test_off_setting_holds_large_pst_with_a_message_that_says_so(env, monkeypatch):
    r, e, src, tmp = env
    e.pop("split_mode")
    r.cfg = config.Config(**{**r.cfg.__dict__, "live_split": "off"})
    monkeypatch.setattr(r, "_mapi_ok", lambda: True)
    assert r._wants_live(e) is False
    assert e["last_error"]["code"] == "LIVE_SPLIT_UNAVAILABLE" and "live_split" in e["last_error"]["message"]
