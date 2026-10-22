"""mapilive の統合テスト（実 MAPI。設計書 §16.4〜§16.6）。

PSTM_INTEGRATION=1 のときだけ実行する。使うのは一時フォルダーに新規作成したテスト用PSTだけで、
元・束とも一時プロファイル（PSTMigration-tmp-*。終了時に削除）で開く。ユーザーの Outlook・IMAP・
Legacy には接続も書き込みもしない（元PSTがユーザーのプロファイルにある場合は §16.12 #8〜#12 で実証済み）。
"""

import gc
import os
from collections import Counter
from datetime import UTC, datetime

import pytest

pytestmark = pytest.mark.skipif(os.environ.get("PSTM_INTEGRATION") != "1", reason="PSTM_INTEGRATION=1 で実行")

pythoncom = pytest.importorskip("pythoncom")
pywintypes = pytest.importorskip("pywintypes")

from pstmigration import chunks, config, mapilive, transfer  # noqa: E402

CUTOFF = config.Config(cutoff="2016-02-21T00:00:00+09:00").cutoff_utc
BEFORE = datetime(2016, 2, 20, 14, 59, 59, tzinfo=UTC)  # JST 2016-02-20 23:59:59
AFTER = datetime(2016, 2, 20, 15, 0, 0, tzinfo=UTC)
PR_BODY_W, PR_CONTAINER_CLASS_W = 0x1000001F, 0x3613001F
PR_ATTACH_METHOD, PR_ATTACH_LONG_FILENAME_W, PR_ATTACH_DATA_BIN = 0x37050003, 0x3707001F, 0x37010102


def add_message(
    folder,
    *,
    mclass="IPM.Note",
    subject="s",
    delivery=None,
    submit=None,
    headers=None,
    key=None,
    body="本文",
    attach=None,
):
    from win32comext.mapi import mapi

    m = folder.CreateMessage(None, 0)
    props = [(mapilive.PR_MESSAGE_CLASS, mclass), (mapilive.PR_SUBJECT, subject), (PR_BODY_W, body)]
    if delivery:
        props.append((mapilive.PR_DELIVERY_TIME, pywintypes.Time(delivery)))
    if submit:
        props.append((mapilive.PR_SUBMIT_TIME, pywintypes.Time(submit)))
    if headers:
        props.append((mapilive.PR_HEADERS, headers))
    if key:
        props.append((mapilive.PR_SEARCH_KEY, key))
    m.SetProps(props)
    if attach is not None:
        _, a = m.CreateAttach(None, 0)
        a.SetProps([(PR_ATTACH_METHOD, 1), (PR_ATTACH_LONG_FILENAME_W, "a.bin")])
        s = a.OpenProperty(PR_ATTACH_DATA_BIN, pythoncom.IID_IStream, 0, mapi.MAPI_MODIFY | mapi.MAPI_CREATE)
        s.Write(attach)
        s.Commit(0)
        a.SaveChanges(0)
    m.SaveChanges(0)


@pytest.fixture
def mapi_env(tmp_path):
    with mapilive.Mapi():
        before = mapilive.profile_names()
        yield tmp_path
        gc.collect()
        assert mapilive.profile_names() == before, "一時プロファイルが残っている"


@pytest.fixture
def source(mapi_env):
    """テスト用の元PST。受信（旧・新・日付不明・非メール・大きいヘッダー）、送信、空のサブフォルダー、予定表。"""
    path = mapi_env / "src.pst"
    dup = os.urandom(16)
    tp = mapilive.TempProfile(path, create=True)
    try:
        w = tp.open_store(writable=True)
        inbox = w.ensure_folder("\\受信トレイ", "IPF.Note")
        add_message(inbox, subject="旧", delivery=BEFORE, headers="Received: x")
        add_message(inbox, subject="新", delivery=AFTER, headers="Received: x")
        add_message(inbox, subject="大きいヘッダーの新", delivery=AFTER, headers="Received: " + "x" * 70_000)
        add_message(inbox, subject="日付不明")
        add_message(inbox, mclass="IPM.Contact", subject="非メール")
        add_message(inbox, subject="重複A", delivery=AFTER, headers="R", key=dup, body="中身1")
        add_message(
            inbox, subject="重複A", delivery=AFTER, headers="R", key=dup, body="中身2"
        )  # 指紋まで同じ・中身違い
        add_message(inbox, subject="添付", delivery=AFTER, headers="R", attach=os.urandom(300_000))
        sent = w.ensure_folder("\\送信済みアイテム", "IPF.Note")
        add_message(sent, subject="送信", delivery=AFTER, submit=BEFORE)  # 送信は送信日時で判定 → 旧
        w.ensure_folder("\\受信トレイ\\空のサブ", "IPF.Note")
        cal = w.ensure_folder("\\予定表", "IPF.Appointment")
        add_message(cal, mclass="IPM.Appointment", subject="予定")
        del w, inbox, sent, cal
    finally:
        tp.close()
    return path


def open_source(path):
    tp = mapilive.TempProfile(path, create=False)
    return tp, tp.open_store(writable=False)


def test_scan_classifies_like_current_method(source):
    tp, src = open_source(source)
    try:
        s = mapilive.scan(src, CUTOFF)
    finally:
        del src
        tp.close()
    by = {f["path"]: f for f in s["scan"]["folders"]}
    inbox = by["\\受信トレイ"]
    assert (inbox["old"], inbox["new"], inbox["unmigrated"], inbox["non_mail"]) == (1, 5, 1, 1)
    assert (by["\\送信済みアイテム"]["old"], by["\\送信済みアイテム"]["new"]) == (1, 0)
    assert by["\\予定表"]["kind"] == "other" and by["\\予定表"]["non_mail"] == 1
    assert "\\受信トレイ\\空のサブ" in by
    mine = {"\\受信トレイ", "\\受信トレイ\\空のサブ", "\\送信済みアイテム"}
    mail = {p for p, _ in s["mail_folders"]}
    assert mail >= mine
    # 余分は PST が自動で作る既定フォルダー（削除済みアイテム・送信トレイ等。名前は言語で変わる）で、空のもの
    assert all(by[p]["old"] + by[p]["new"] + by[p]["unmigrated"] + by[p]["non_mail"] == 0 for p in mail - mine)
    assert len(s["items"]["old"]) == 2 and len(s["items"]["new"]) == 5


def test_source_opened_read_only_rejects_writes(source):
    from win32comext.mapi import mapi

    tp, src = open_source(source)
    try:
        inbox = src.folder("\\受信トレイ")
        eid = bytes.fromhex(mapilive.scan(src, CUTOFF)["items"]["new"][0]["e"])
        with pytest.raises(pywintypes.com_error):
            src.store.OpenEntry(eid, None, mapi.MAPI_MODIFY)
        with pytest.raises(pywintypes.com_error):
            inbox.DeleteMessages([eid], 0, None, 0)
        assert len(mapilive.scan(src, CUTOFF)["items"]["new"]) == 5  # 消えていない
        del inbox
    finally:
        del src
        tp.close()


def test_copy_to_chunk_then_verify_with_hashes_and_release(source, mapi_env):
    tp, src = open_source(source)
    chunk = mapi_env / "new_mail_from_20160221_PC_001.pst"
    try:
        s = mapilive.scan(src, CUTOFF)
        need = chunks.hash_needed(s["all_rows"])
        assert len(need) == 1  # 「重複A」2通（SearchKey・指紋まで同じ）
        items = s["items"]["new"]
        mapilive.add_hashes(src, items, need)
        w = mapilive.ChunkWriter(chunk, src, s["mail_folders"])  # 最初の束：全階層を作る
        w.copy(items[:2])
        w.copy(items[2:])
        w.close()
    finally:
        del src
        tp.close()
    transfer.open_exclusive(chunk).Close()  # 閉じた直後に解放されている
    got = mapilive.read_chunk(chunk, need)
    assert chunks.mismatches(items, got["rows"], need) == []
    assert {f for f in got["folders"]} >= {"\\受信トレイ", "\\受信トレイ\\空のサブ", "\\送信済みアイテム"}
    transfer.open_exclusive(chunk).Close()
    assert Counter(r.get("h") is not None for r in got["rows"]) == Counter({False: 3, True: 2})


def test_wrong_copy_is_detected_by_hash(source, mapi_env):
    """「重複A」の中身1・中身2 のうち、中身1を2回コピーした束は照合で不一致になる。"""
    tp, src = open_source(source)
    chunk = mapi_env / "new_mail_from_20160221_PC_001.pst"
    try:
        s = mapilive.scan(src, CUTOFF)
        need = chunks.hash_needed(s["all_rows"])
        dups = [i for i in s["items"]["new"] if (i["k"], i["p"]) in need]
        mapilive.add_hashes(src, dups, need)
        w = mapilive.ChunkWriter(chunk, src, None)
        w.copy([dups[0], dups[0]])  # 取り違え：A・B のつもりが A・A
        w.close()
    finally:
        del src
        tp.close()
    assert chunks.mismatches(dups, mapilive.read_chunk(chunk, need)["rows"], need)


def test_leftover_temp_profiles_are_cleaned_up(mapi_env):
    before = mapilive.profile_names()
    tp = mapilive.TempProfile(mapi_env / "x.pst", create=True)
    tp.abandon()  # 異常終了を模す：ログオフもプロファイル削除もしない
    assert any(n.startswith(mapilive.TEMP_PREFIX) for n in mapilive.profile_names())
    removed = mapilive.cleanup_temp_profiles()
    assert removed and all(n.startswith(mapilive.TEMP_PREFIX) for n in removed)
    assert mapilive.profile_names() == before
