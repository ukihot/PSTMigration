"""ライブ分割の通し試験と、プロセス強制終了からの復旧（設計書 §16、統合テスト）。

PSTM_INTEGRATION=1 のときだけ実行する。実 MAPI・実 7z。NAS はローカルフォルダー（転送は本物の transfer）。
元PSTも束も、一時フォルダーに新規作成したテスト用PSTだけ。元PSTは名前付きの一時プロファイル
（PSTM-IT-source-*。本ツールの後片付けの対象外）に入れ、ユーザーの Outlook・IMAP・Legacy には触れない。
"""

import contextlib
import gc
import hashlib
import json
import os
import shutil
import subprocess
import sys
import time
from collections import Counter
from datetime import UTC, datetime
from pathlib import Path

import pytest

pytestmark = pytest.mark.skipif(os.environ.get("PSTM_INTEGRATION") != "1", reason="PSTM_INTEGRATION=1 で実行")

pythoncom = pytest.importorskip("pythoncom")
pywintypes = pytest.importorskip("pywintypes")

from pstmigration import config, mapilive, pipeline  # noqa: E402
from pstmigration.state import new_entry, read_json, write_json  # noqa: E402

GB = 1024**3
CUTOFF = "2016-02-21T00:00:00+09:00"
OLD = datetime(2015, 6, 1, tzinfo=UTC)
NEW = datetime(2020, 6, 1, tzinfo=UTC)
PR_BODY_W = 0x1000001F
PR_ATTACH_METHOD, PR_ATTACH_LONG_FILENAME_W, PR_ATTACH_DATA_BIN = 0x37050003, 0x3707001F, 0x37010102
WORKER = Path(__file__).with_name("live_worker.py")
SRC = Path(__file__).parents[2] / "src"


def add(folder, subject, when, body, attach=None, key=None):
    from win32comext.mapi import mapi

    m = folder.CreateMessage(None, 0)
    props = [
        (mapilive.PR_MESSAGE_CLASS, "IPM.Note"),
        (mapilive.PR_SUBJECT, subject),
        (PR_BODY_W, body),
        (mapilive.PR_DELIVERY_TIME, pywintypes.Time(when)),
        (mapilive.PR_HEADERS, "Received: test"),
    ]
    if key:
        props.append((mapilive.PR_SEARCH_KEY, key))
    m.SetProps(props)
    if attach:
        _, a = m.CreateAttach(None, 0)
        a.SetProps([(PR_ATTACH_METHOD, 1), (PR_ATTACH_LONG_FILENAME_W, "a.bin")])
        s = a.OpenProperty(PR_ATTACH_DATA_BIN, pythoncom.IID_IStream, 0, mapi.MAPI_MODIFY | mapi.MAPI_CREATE)
        s.Write(attach)
        s.Commit(0)
        a.SaveChanges(0)
    m.SaveChanges(0)


@pytest.fixture
def world(tmp_path):
    """テスト用の元PST（旧・新・添付・SearchKeyまで同じで本文違い）と、ライブ分割の設定・state。"""
    profile = f"PSTM-IT-source-{os.getpid()}"
    expected = {"old": Counter(), "new": Counter()}  # オラクル：作ったメールの (件名, 本文)
    with mapilive.Mapi():
        profiles_before = mapilive.profile_names()
        src_path = tmp_path / "source.pst"
        tp = mapilive.TempProfile(src_path, create=True, name=profile)
        w = tp.open_store(writable=True)
        inbox, sent = w.ensure_folder("\\受信トレイ", "IPF.Note"), w.ensure_folder("\\送信済みアイテム", "IPF.Note")
        w.ensure_folder("\\受信トレイ\\空", "IPF.Note")
        dup = os.urandom(16)
        for i in range(24):
            bucket, when = ("old", OLD) if i % 3 == 0 else ("new", NEW)
            folder = inbox if i % 2 else sent
            body = f"本文{i}"
            add(folder, f"m{i}", when, body, attach=os.urandom(120_000))  # 圧縮できない添付
            expected[bucket][(f"m{i}", body)] += 1
        for body in ("重複の中身1", "重複の中身2"):  # SearchKey・指紋まで同じで本文だけ違う
            add(inbox, "dup", NEW, body, key=dup)
            expected["new"][("dup", body)] += 1
        store_id = mapilive._query(tp.session.GetMsgStoresTable(0), (mapilive.PR_ENTRYID,))[0][0][1].hex()
        del w, inbox, sent, folder  # ループ変数 folder も参照を持つ（残すとセッションが閉じない）
        gc.collect()
        tp.session.Logoff(0, 0, 0)  # プロファイルは残し、セッションだけ閉じる
        tp = None
        gc.collect()
    # ここで MAPI を終了する（元PSTのファイルを手放させる。テスト中は各処理が自分で MAPI を始める）
    try:
        nas, home = tmp_path / "nas", tmp_path / "home"
        nas.mkdir()
        settings = {
            "cutoff": CUTOFF,
            "server_root": str(nas),
            "server_user": "",
            "work_dir": str(tmp_path / "work"),
            "disk_reserve_gb": 1,
            "chunk_max_gb": 1_000_000 / GB,  # 束の目標 約1MB（添付120KB × 8通ほど）
            "chunk_min_gb": 10_000 / GB,
            "batch_target_mb": 0.3,
            "split_pause_ms": 0,
        }
        cfg_file = tmp_path / "cfg.json"
        cfg_file.write_text(json.dumps(settings), encoding="utf-8")
        e = new_entry("srckey", str(src_path), False)
        e.update(split_mode="live", store_id=store_id, source_size=src_path.stat().st_size, cutoff=CUTOFF)
        write_json(home / "state.json", {"psts": {"srckey": e}, "mailbox": "it@example.jp"})
        yield {
            "cfg": config.Config(**settings),
            "cfg_file": cfg_file,
            "home": home,
            "nas": nas,
            "profile": profile,
            "src_path": src_path,
            "expected": expected,
            "tmp": tmp_path,
        }
    finally:
        with mapilive.Mapi():
            with contextlib.suppress(Exception):  # 元PSTの名前付きプロファイルを消す
                mapilive._mapi().MAPIAdminProfiles(0).DeleteProfile(profile, 0)
            mapilive.cleanup_temp_profiles()
            assert mapilive.profile_names() == profiles_before, "プロファイルが残っている"


def nas_contents(w) -> dict:
    """オラクル：NAS 上の束（PST・7z を展開したもの）を読み直し、側ごとに (件名, 本文) を数える。"""
    out = {"old": Counter(), "new": Counter()}
    names = sorted(p.name for p in w["nas"].rglob("*") if p.is_file())
    assert not [n for n in names if n.endswith(".partial")], ".partial が残っている"
    check = w["tmp"] / "check"
    shutil.rmtree(check, ignore_errors=True)
    check.mkdir()
    for p in sorted(w["nas"].rglob("*_mail_*")):
        bucket = "new" if p.name.startswith("new_mail_from_") else "old"
        if p.suffix == ".7z":
            subprocess.run(["7z", "x", f"-o{check}", str(p), "-y"], check=True, capture_output=True)
            pst = check / (p.stem + ".pst")
        else:
            pst = check / p.name
            shutil.copy(p, pst)
        tp = mapilive.TempProfile(pst, create=False)
        try:
            st = tp.open_store(writable=False)
            for _path, eid, _kind, _ in st.walk():
                for props in st.rows(st.open(eid)):
                    cols = dict(zip(mapilive.MSG_COLUMNS, props, strict=True))
                    msg = st.store.OpenEntry(cols[mapilive.PR_ENTRYID][1], None, 0)
                    subj = msg.GetProps((mapilive.PR_SUBJECT, PR_BODY_W), 0)[1]
                    out[bucket][(subj[0][1], subj[1][1])] += 1
                    del msg
            del st
        finally:
            gc.collect()
            tp.close()
    return out


def source_rows(w):
    """元PSTの論理的な中身（EntryID・SearchKey・サイズ・指紋）。"""
    with mapilive.Mapi():
        tp = mapilive.TempProfile(w["src_path"], create=False, name=f"PSTM-IT-read-{os.getpid()}")
        try:
            st = tp.open_store(writable=False)
            rows = sorted((r["e"], r["k"], r["s"], r["p"]) for r in mapilive.scan(st, w["cfg"].cutoff_utc)["all_rows"])
            del st
        finally:
            gc.collect()
            tp.close()
    return rows


def file_sha(path):
    """元PST全体の SHA-256 と、読めるようになるまで待った秒数（補助情報。完了条件にはしない §16.9）。

    処理の直後はウイルス対策などが一時的に開いていることがあるため、最大60秒待つ。
    """
    t0 = time.monotonic()
    while True:
        try:
            return hashlib.sha256(path.read_bytes()).hexdigest(), round(time.monotonic() - t0, 1)
        except PermissionError:
            if time.monotonic() - t0 > 60:
                return None, None
            time.sleep(0.5)


def test_live_split_end_to_end(world, capsys):
    w = world
    rows_before, sha_before = source_rows(w), file_sha(w["src_path"])
    r = pipeline.Runner(w["cfg"], w["home"])
    r.profile = w["profile"]
    e = r.st.psts["srckey"]
    r._process_live(e)
    assert e["status"] == "COMPLETED", e.get("last_error")
    assert sum(c["bucket"] == "new" for c in e["chunks"]) >= 2 and sum(c["bucket"] == "old" for c in e["chunks"]) >= 2
    with mapilive.Mapi():
        got = nas_contents(w)
    assert got == w["expected"], "NAS の中身が、作ったメールと一致しない"
    rows_after, sha_after = source_rows(w), file_sha(w["src_path"])
    assert rows_after == rows_before, "元PSTの中身（論理）が変わった"
    with capsys.disabled():  # 完了条件にはしない（§16.9）。補助情報として記録する
        (h0, w0), (h1, w1) = sha_before, sha_after
        verdict = "読めず" if h1 is None else ("一致" if h0 == h1 else "不一致")
        print(f"\n[補助] 元PST全体の SHA-256 は処理の前後で {verdict}（読めるまで 前 {w0}秒 / 後 {w1}秒）")
        print(f"[補助] サイズ差の分布: {Counter(k for c in e['chunks'] for k in c.get('size_delta', {}))}")


def run_worker(w):
    env = {**os.environ, "PYTHONPATH": str(SRC), "PYTHONIOENCODING": "utf-8"}
    args = [sys.executable, str(WORKER), str(w["cfg_file"]), str(w["home"]), w["profile"], "srckey"]
    return subprocess.Popen(args, env=env, stdout=subprocess.PIPE, stderr=subprocess.PIPE)


def kill_when(w, proc, predicate, timeout=300):
    """state.json が条件を満たしたらプロセスを強制終了する（TerminateProcess。finally は走らない）。"""
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if proc.poll() is not None:
            pytest.fail(f"止める前に終わった: {proc.stdout.read().decode()} {proc.stderr.read().decode()[-2000:]}")
        chunks_now = (read_json(w["home"] / "state.json", {}).get("psts", {}).get("srckey") or {}).get("chunks") or []
        if predicate(chunks_now):
            proc.kill()
            proc.wait()
            return chunks_now
        time.sleep(0.05)
    proc.kill()
    pytest.fail("条件に達しなかった")


def test_recovery_after_process_is_killed_twice(world):
    w = world
    first = kill_when(w, run_worker(w), lambda cs: any(c["status"] == "UPLOADED" for c in cs))
    n_uploaded = sum(c["status"] == "UPLOADED" for c in first)
    kill_when(
        w,
        run_worker(w),
        lambda cs: (
            sum(c["status"] == "UPLOADED" for c in cs) > n_uploaded
            or any(c["status"] in ("VERIFIED", "BUILT") for c in cs[n_uploaded:])
        ),
    )
    proc = run_worker(w)
    out, err = proc.communicate(timeout=600)
    assert proc.returncode == 0, err.decode()[-3000:]
    e = read_json(w["home"] / "state.json", {})["psts"]["srckey"]
    assert e["status"] == "COMPLETED", e.get("last_error")
    with mapilive.Mapi():
        assert not [n for n in mapilive.profile_names() if n.startswith(mapilive.TEMP_PREFIX)], (
            "一時プロファイルが残った"
        )
        got = nas_contents(w)
    assert got == w["expected"], "強制終了を挟むと NAS の中身が作ったメールと一致しない（重複・欠落）"
    names = [p.name for p in w["nas"].rglob("*_mail_*")]
    assert len(names) == len(set(names)) == len(e["chunks"])
