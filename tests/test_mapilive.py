"""mapilive の Outlook・MAPI に依存しない部分（設計書 §16.4、§16.6）。MAPI 本体は tests/integration で確認する。"""

from datetime import UTC, datetime

import pytest

from pstmigration import config, mapilive

CUTOFF = config.Config(cutoff="2016-02-21T00:00:00+09:00").cutoff_utc
PT_ERROR = 0x000A
E_OUTOFMEMORY = -2147024882  # MAPI_E_NOT_ENOUGH_MEMORY：値はあるが大きすぎてテーブルに載らない
E_NOT_FOUND = -2147221233  # MAPI_E_NOT_FOUND：値が無い


@pytest.mark.parametrize(
    "container_class, folder_type, kind",
    [
        (None, 1, "mail"),  # クラス無し（Outlook はメールフォルダー扱い）
        ("IPF.Note", 1, "mail"),
        ("IPF.Note.OutlookHomepage", 1, "mail"),
        ("ipf.note", 1, "mail"),
        ("IPF.Imap", 1, "mail"),
        ("IPF.Appointment", 1, "other"),
        ("IPF.Contact", 1, "other"),
        ("IPF.Task", 1, "other"),
        ("IPF.StickyNote", 1, "other"),
        ("IPF.Note", 2, "search"),  # 検索フォルダー
    ],
)
def test_folder_kind_matches_outlook_default_item_type(container_class, folder_type, kind):
    assert mapilive.folder_kind(container_class, folder_type) == kind


def test_fingerprint_is_stable_and_distinguishes_fields():
    t = datetime(2016, 2, 20, 15, 0, 0, tzinfo=UTC)
    a = mapilive.fingerprint("件名", t, t, "IPM.Note", True)
    assert a == mapilive.fingerprint("件名", t, t, "IPM.Note", True)
    for other in (
        mapilive.fingerprint("件名2", t, t, "IPM.Note", True),
        mapilive.fingerprint("件名", None, t, "IPM.Note", True),
        mapilive.fingerprint("件名", t, None, "IPM.Note", True),
        mapilive.fingerprint("件名", t, t, "IPM.Note.SMIME", True),
        mapilive.fingerprint("件名", t, t, "IPM.Note", False),
    ):
        assert other != a


def test_table_value_errors():
    """内容テーブルの値。無い列は PT_ERROR で返る。大きすぎる値は『ある』のにエラーで返る。"""
    tag = 0x007D001F  # PR_TRANSPORT_MESSAGE_HEADERS
    err_tag = (tag & 0xFFFF0000) | PT_ERROR
    assert mapilive.present((tag, "Received: ..."))
    assert mapilive.present((err_tag, E_OUTOFMEMORY))  # ヘッダーはある（受信メール）
    assert not mapilive.present((err_tag, E_NOT_FOUND))
    assert mapilive.value((err_tag, E_NOT_FOUND)) is None
    assert mapilive.value((tag, "x")) == "x"


def test_classify_folder_rows_into_plan_items_and_report():
    def raw(eid, mclass, delivery, submit, headers, size=10, key=b"K", subject="s", att=False):
        cols = {
            mapilive.PR_ENTRYID: eid,
            mapilive.PR_SEARCH_KEY: key,
            mapilive.PR_MESSAGE_CLASS: mclass,
            mapilive.PR_SUBJECT: subject,
            mapilive.PR_DELIVERY_TIME: delivery,
            mapilive.PR_SUBMIT_TIME: submit,
            mapilive.PR_HEADERS: headers,
            mapilive.PR_SIZE: size,
            mapilive.PR_HASATTACH: att,
        }
        return [(tag, v) if v is not None else ((tag & 0xFFFF0000) | PT_ERROR, E_NOT_FOUND) for tag, v in cols.items()]

    just_before = datetime(2016, 2, 20, 14, 59, 59, tzinfo=UTC)  # JST 2016-02-20 23:59:59
    just_after = datetime(2016, 2, 20, 15, 0, 0, tzinfo=UTC)
    rows = [
        raw(b"\x01", "IPM.Note", just_before, None, "Received: x"),
        raw(b"\x02", "IPM.Note", just_after, None, "Received: x"),
        raw(b"\x03", "IPM.Note", None, None, None),  # 日付不明
        raw(b"\x04", "IPM.Contact", None, None, None),  # メールフォルダー内の非メール
        raw(b"\x05", "IPM.Note", just_after, just_before, None, size=99),  # 送信：送信日時で判定 → 旧
    ]
    rec, items, all_rows, problems = mapilive.classify_rows("\\受信トレイ", rows, CUTOFF)
    assert (rec["old"], rec["new"], rec["unmigrated"], rec["non_mail"]) == (2, 1, 1, 1)
    assert rec["old_bytes"] == 10 + 99 and rec["new_bytes"] == 10
    assert [i["e"] for i in items["old"]] == ["01", "05"] and [i["e"] for i in items["new"]] == ["02"]
    assert {i["f"] for b in items.values() for i in b} == {"\\受信トレイ"}
    assert len(all_rows) == 5  # ハッシュ対象の判定（hash_needed）には計画外の行も入れる
    assert problems == [{"code": "NO_DATE", "folder": "\\受信トレイ", "message_class": "IPM.Note"}]


def test_digest_is_unambiguous_about_part_boundaries():
    assert mapilive.digest([b"ab", b"c"]) != mapilive.digest([b"a", b"bc"])
    assert mapilive.digest([b"ab", b"c"]) == mapilive.digest([b"ab", b"c"])


@pytest.mark.parametrize("name", ["⑴相談", "㈶テスト", "⓪連絡", "補助文字📨"])
def test_create_folder_uses_unicode_without_ansi_conversion(name):
    from win32comext.mapi import mapi

    calls = []

    class Folder:
        def CreateFolder(self, kind, folder_name, comment, iid, flags):
            assert flags & mapi.MAPI_UNICODE
            calls.append(folder_name)
            return self

        def SetProps(self, props):
            pass

        def SaveChanges(self, flags):
            pass

    store = object.__new__(mapilive.MapiStore)
    store._folders = {}
    store._root_eid = b"root"
    store.open = lambda eid: Folder()
    store.ensure_folder("\\" + name, "IPF.Note")
    assert calls == [name]


def test_folder_lookup_follows_same_named_siblings_and_reports_missing_as_source_change():
    # ルートに「受信トレイ」が2つ。探すフォルダーは2つ目の配下にだけある
    tree = {b"root": [b"in1", b"in2"], b"in1": [], b"in2": [b"t"], b"t": []}
    names = {b"in1": "受信トレイ", b"in2": "受信トレイ", b"t": "【　取引先　】"}
    store = object.__new__(mapilive.MapiStore)
    store._folders, store._root_eid = {}, b"root"
    store.open = lambda eid: eid
    store._children = lambda f: [(c, names[c], None, 1) for c in tree[f]]
    assert store.folder("\\受信トレイ\\【　取引先　】") == b"t"
    with pytest.raises(mapilive.StepError) as ei:
        store.folder("\\受信トレイ\\改名済み")
    assert ei.value.code == "SOURCE_CHANGED_DURING_COPY" and ei.value.final


# ---- 連絡先バックアップのコピー計画（どのフォルダーを、どこへ） --------------------------------


def test_contact_folders_are_copied_with_their_subfolders_under_one_box_per_store():
    walk = [  # MapiStore.walk() の形：(パス, EntryID, 種類, コンテナークラス)
        (r"\受信トレイ", b"1", "mail", "IPF.Note"),
        (r"\連絡先", b"2", "other", "IPF.Contact"),
        (r"\連絡先\取引先", b"3", "other", "IPF.Contact"),
        (r"\連絡先\取引先\古い", b"4", "other", "IPF.Contact"),
        (r"\予定表", b"5", "other", "IPF.Appointment"),
        (r"\受信トレイ\社内アドレス帳", b"6", "other", "IPF.Contact"),  # メールフォルダーの下の連絡先
        (r"\受信トレイ\社内アドレス帳\メモ", b"7", "mail", "IPF.Note"),  # 連絡先の下はメールでもコピー
    ]
    plan = mapilive.contact_copy_plan(walk, box="1. Legacy")
    assert plan == [
        (r"\連絡先", b"2", r"\1. Legacy\連絡先", "IPF.Contact"),
        (r"\連絡先\取引先", b"3", r"\1. Legacy\連絡先\取引先", "IPF.Contact"),
        (r"\連絡先\取引先\古い", b"4", r"\1. Legacy\連絡先\取引先\古い", "IPF.Contact"),
        (r"\受信トレイ\社内アドレス帳", b"6", r"\1. Legacy\社内アドレス帳", "IPF.Contact"),
        (r"\受信トレイ\社内アドレス帳\メモ", b"7", r"\1. Legacy\社内アドレス帳\メモ", "IPF.Note"),
    ]


def test_store_without_contact_folders_has_empty_plan():
    assert mapilive.contact_copy_plan([(r"\受信トレイ", b"1", "mail", "IPF.Note")], box="2. IMAP") == []
