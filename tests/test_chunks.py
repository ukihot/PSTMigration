"""ライブ分割方式の束管理（設計書 §16）。Outlook 不要。"""

from collections import Counter

import pytest
from hypothesis import given, settings
from hypothesis import strategies as st

from pstmigration import chunks

GB = 1024**3


def up(bucket, start, end, seq=1):
    return {"bucket": bucket, "seq": seq, "range": [start, end], "status": "UPLOADED"}


# ---- 段階1：束の範囲と完了判定（§16.9 の 1・2） ---------------------------------


@pytest.mark.parametrize(
    "records, counts, ok",
    [
        ([], {"old": 0, "new": 0}, True),  # 両側0件：束なしで完了
        ([up("new", 0, 5), up("new", 5, 9, 2)], {"old": 0, "new": 9}, True),
        ([up("new", 0, 5), up("new", 6, 9, 2)], {"old": 0, "new": 9}, False),  # 欠落（5番）
        ([up("new", 0, 6), up("new", 5, 9, 2)], {"old": 0, "new": 9}, False),  # 重複（5番）
        ([up("new", 0, 5)], {"old": 0, "new": 9}, False),  # 途中まで
        ([up("new", 0, 9), up("new", 9, 10, 2)], {"old": 0, "new": 9}, False),  # 計画外
        ([up("old", 0, 1)], {"old": 0, "new": 0}, False),  # 0件の側に束がある
        ([{**up("new", 0, 9), "status": "BUILT"}], {"old": 0, "new": 9}, False),  # 未転送
        ([up("new", 5, 9, 2), up("new", 0, 5)], {"old": 0, "new": 9}, True),  # 並び順は問わない
        ([up("new", 0, 0)], {"old": 0, "new": 0}, False),  # 空の束
    ],
)
def test_completion_requires_exact_cover_by_uploaded_chunks(records, counts, ok):
    assert (chunks.completion_problems(records, counts) == []) is ok


def test_resume_point_is_end_of_contiguous_uploaded_prefix():
    records = [up("new", 0, 5), up("new", 5, 9, 2), {**up("new", 9, 12, 3), "status": "BUILT"}]
    assert chunks.uploaded_until(records, "new") == 9
    assert chunks.uploaded_until(records, "old") == 0


@pytest.mark.parametrize(
    "sizes, start, target, max_items, end",
    [
        ([10, 10, 10, 10], 0, 25, 500, 2),  # 目標を超えない所まで
        ([10, 10, 10, 10], 0, 1000, 3, 3),  # 件数上限
        ([100, 1, 1], 0, 10, 500, 1),  # 1通で目標超え：その1通だけ
        ([1, 1, 1], 3, 10, 500, 3),  # 終端
        ([5, 5, 5], 1, 10, 500, 3),  # 途中から
    ],
)
def test_next_batch_by_message_size(sizes, start, target, max_items, end):
    items = [{"s": s} for s in sizes]
    assert chunks.next_batch(items, start, target, max_items) == end


@pytest.mark.parametrize(
    "free_gb, expected_gb",
    [
        (100, 5),  # 上限5GB
        (15, 4),  # (15-5)/2.5
        (7.5, 1),  # (7.5-5)/2.5 = 1GB（下限ちょうど）
        (7, None),  # 下限1GB未満：始めない
        (3, None),
    ],
)
def test_chunk_target_from_free_space(free_gb, expected_gb):
    got = chunks.chunk_target(int(free_gb * GB), max_bytes=5 * GB, min_bytes=1 * GB, reserve_bytes=5 * GB, factor=2.5)
    assert got == (None if expected_gb is None else int(expected_gb * GB))


def test_chunk_names_are_deterministic():
    assert chunks.chunk_name("new_mail_from_20160221_PC1", 1, ".pst") == "new_mail_from_20160221_PC1_001.pst"
    assert chunks.chunk_name("old_mail_before_20160221_PC1", 12, ".7z") == "old_mail_before_20160221_PC1_012.7z"


# ---- 段階2：束の照合（§16.6）。多重集合＋重複グループの本文・添付ハッシュ ------------------


def row(f, k, p, h=None):
    """f=フォルダー, k=SearchKey, p=指紋（件名・日時・クラス・添付有無）, h=本文・添付のハッシュ"""
    return {"f": f, "k": k, "p": p, "h": h}


def test_same_key_different_mail_is_detected_by_fingerprint():
    # 原本 A(key=X), B(key=X) → 出力 A, A。SearchKey だけの多重集合では見逃す例
    expected = [row("受信", "X", "A"), row("受信", "X", "B")]
    actual = [row("受信", "X", "A"), row("受信", "X", "A")]
    assert chunks.mismatches(expected, actual, chunks.hash_needed(expected))


def test_identical_fingerprints_are_told_apart_by_content_hash():
    expected = [row("受信", "X", "A", "c1"), row("受信", "X", "A", "c2")]
    need = chunks.hash_needed(expected)
    assert need == {("X", "A")}  # フォルダーを問わず「SearchKey＋指紋」で判定
    assert chunks.mismatches(expected, [row("受信", "X", "A", "c1")] * 2, need)
    assert not chunks.mismatches(expected, list(reversed(expected)), need)


def test_misplaced_copy_from_other_folder_with_same_key_and_fingerprint_is_detected():
    """別フォルダーにある「SearchKey・指紋が同じで中身が違う」メールが、計画したフォルダーへコピーされても検出する。

    PST の EntryID はストア全体で一意なので、別フォルダーのメールも計画したフォルダーへコピーされうる。
    """
    source = [row("受信", "X", "A"), row("受信", "X", "B"), row("送信", "X", "A")]
    need = chunks.hash_needed(source)
    expected = [row("受信", "X", "A", "c1")]
    copied = [row("受信", "X", "A", "c2")]  # 送信のメール（中身 c2）が受信へコピーされた
    assert chunks.mismatches(expected, copied, need)


def test_wrong_folder_missing_and_extra_are_detected():
    expected = [row("受信", "X", "A"), row("送信", "Y", "B")]
    need = chunks.hash_needed(expected)
    assert chunks.mismatches(expected, [row("受信", "X", "A"), row("受信", "Y", "B")], need)
    assert chunks.mismatches(expected, [row("受信", "X", "A")], need)
    assert chunks.mismatches(expected, [*expected, row("送信", "Y", "B")], need)
    assert not chunks.mismatches(expected, expected, need)


# 衝突しやすいよう、値の種類を少なくする
_mail = st.tuples(
    st.sampled_from(["受信", "送信"]),  # フォルダー
    st.sampled_from(["X", "Y"]),  # SearchKey（別のメールでも同じ値になりうる）
    st.sampled_from(["A", "B"]),  # 指紋
    st.sampled_from(["c1", "c2", "c3"]),  # 本文・添付の中身
)


@st.composite
def _copy_case(draw):
    """元フォルダーの全メール（計画外の日付不明・非メールも含む）、束の期待範囲、壊れたかもしれないコピー結果。"""
    source = draw(st.lists(_mail, min_size=1, max_size=10))
    planned = source[: draw(st.integers(1, len(source)))]  # 残りは計画外（日付不明など）
    start = draw(st.integers(0, len(planned) - 1))
    end = draw(st.integers(start + 1, len(planned)))
    copied = list(planned[start:end])
    for _ in range(draw(st.integers(0, 3))):  # CopyMessages の取り違え・欠落・重複・別フォルダー
        op = draw(st.sampled_from(["drop", "dup", "replace", "move"]))
        if not copied and op != "replace":
            continue
        i = draw(st.integers(0, max(0, len(copied) - 1)))
        if op == "drop":
            copied.pop(i)
        elif op == "dup":
            copied.append(copied[i])
        elif op == "replace":
            m = draw(st.sampled_from(source))  # コピーできるのは元にあるメールだけ
            copied[i : i + 1] = [m]
        else:
            f, k, p, c = copied[i]
            copied[i] = ("送信" if f == "受信" else "受信", k, p, c)
    return source, planned[start:end], copied


def test_mismatch_names_which_fingerprint_fields_differ_without_subject_value():
    fp = lambda subj, att: f"{subj}|2020-01-01T00:00:00+00:00|2020-01-01T00:00:00+00:00|IPM.Note|{att}"  # noqa: E731
    out = chunks.mismatches([row("受信", "X", fp("秘密|件名", "0"))], [row("受信", "X", fp("秘密|件名", "1"))], set())
    assert out[0] == "差分 受信 key=X: 添付 0→1" and "秘密" not in " ".join(out)
    out = chunks.mismatches([row("受信", "X", fp("a", "0"))], [row("受信", "X", fp("b", "0"))], set())
    assert out[0] == "差分 受信 key=X: 件名"


def test_missing_required_hash_fails():
    """ハッシュが要る組で、期待側・束側ともハッシュが無い行は一致にしない。"""
    r = {"f": "Inbox", "k": "KEY-001", "p": ("subject", "date")}
    need = {("KEY-001", ("subject", "date"))}
    assert chunks.mismatches([r], [r], need)


@settings(max_examples=2000)
@given(_copy_case(), st.data())
def test_verification_agrees_with_true_identity(case, data):
    """オラクル：メールの本当の中身（フォルダー・SearchKey・指紋・本文）の多重集合が一致し、
    かつハッシュが必要な行のハッシュが1つも欠けていないときだけ合格。

    照合関数は「フォルダー＋SearchKey＋指紋」の多重集合を使い、それでも重複する組だけ本文ハッシュを使う。
    中身が違うのに合格する（見逃し）も、中身が同じなのに不合格になる（誤検出）も無いこと。
    ハッシュの計算に失敗した行（h が欠けた行）があれば、中身が同じでも合格にしないこと。
    """
    source, expected, copied = case
    need = chunks.hash_needed([row(f, k, p) for f, k, p, _ in source])  # 元ストアの全行から決める
    lost = set()  # ハッシュの計算に失敗した行（期待側・実測側の位置）

    def as_rows(mails, side):  # ハッシュは必要な組の分だけ計算される（実装でも同じ）。ときどき失敗する
        out = []
        for i, (f, k, p, c) in enumerate(mails):
            if (k, p) in need and data.draw(st.booleans(), label=f"{side}[{i}] のハッシュ計算に失敗"):
                lost.add((side, i))
                out.append(row(f, k, p, None))
            else:
                out.append(row(f, k, p, c if (k, p) in need else None))
        return out

    exp_rows, got_rows = as_rows(expected, "exp"), as_rows(copied, "got")
    truth_ok = Counter(expected) == Counter(copied) and not lost
    assert (chunks.mismatches(exp_rows, got_rows, need) == []) == truth_ok


# ---- 設定の検証（§16.11 の追加項目）。0 だと束のコピーが進まず無限ループになる ----------------

from pstmigration import config  # noqa: E402


@pytest.mark.parametrize(
    "bad",
    [
        {"batch_target_mb": 0},
        {"batch_items": 0},
        {"chunk_max_gb": 0},
        {"chunk_min_gb": 0},
        {"chunk_min_gb": 6, "chunk_max_gb": 5},
        {"chunk_disk_factor": 0.9},  # 束PST＋7z が入らない比率
        {"plan_size_factor": 0.9},
        {"live_split": "yes"},
    ],
)
def test_invalid_live_split_settings_are_rejected(bad):
    with pytest.raises(ValueError):
        config.Config(**bad)


def test_live_split_defaults():
    c = config.Config()
    assert (c.live_split, c.chunk_max_gb, c.chunk_min_gb, c.batch_target_mb, c.batch_items) == ("auto", 5, 1, 100, 500)
    assert (c.chunk_disk_factor, c.plan_size_factor) == (2.5, 1.15)


# ---- 元PSTの変更検知（§16.7） ------------------------------------------------------------


def src_row(e, f="受信", k="K", p="P"):
    return {"e": e, "f": f, "k": k, "p": p}


def test_unchanged_source_has_no_failures_and_nothing_added():
    plan = [src_row("1"), src_row("2")]
    failures, added = chunks.source_changes(plan, {"1", "2", "x"}, [src_row("1"), src_row("2"), src_row("x")])
    assert failures == [] and added == {}


def test_planned_mail_deleted_moved_or_rekeyed_fails():
    plan = [src_row("1"), src_row("2"), src_row("3")]
    now = [src_row("2", f="削除済み"), src_row("3", k="K2")]  # 1:消えた 2:移された 3:SearchKey が変わった
    failures, _ = chunks.source_changes(plan, {"1", "2", "3"}, now)
    assert len(failures) == 3


def test_read_or_flag_change_is_allowed():
    """既読・フラグ・分類の変更（行の他の列が変わる）は FAILED にしない。"""
    changed = {**src_row("1"), "read": True, "flag": "follow-up", "s": 999}
    failures, added = chunks.source_changes([src_row("1")], {"1"}, [changed])
    assert failures == [] and added == {}


def test_mail_added_after_scan_is_counted_per_folder_not_failed():
    plan = [src_row("1")]
    now = [src_row("1"), src_row("n1"), src_row("n2", f="送信"), src_row("n3")]
    failures, added = chunks.source_changes(plan, {"1"}, now)
    assert failures == [] and added == {"受信": 2, "送信": 1}


# ---- 処理方式の選択（§16.3） ------------------------------------------------------------

from pstmigration import rules  # noqa: E402


@pytest.mark.parametrize(
    "live_split, size_gb, free_gb, mapi_ok, mode",
    [
        ("auto", 5, 100, True, "snapshot"),  # 現行方式で処理できる → 現行（新方式の検証が終わるまでは現行優先）
        ("auto", 5, 100, False, "snapshot"),
        ("auto", 47.5, 200, True, "live"),  # ログオフ時の上限（10GB）超え
        ("auto", 8, 20, True, "live"),  # 上限内だが空き不足（8×2.5+5=25GB 必要）
        ("auto", 47.5, 200, False, "wait"),  # MAPI が使えない：容量の足りない現行方式に戻さず保留
        ("always", 5, 100, True, "live"),
        ("always", 5, 100, False, "snapshot"),  # always でも MAPI が無ければ、現行で処理できるなら現行
        ("always", 47.5, 200, False, "wait"),
        ("off", 47.5, 200, True, "wait"),  # off：新方式を使わない。現行で処理できないものは保留
        ("off", 5, 100, True, "snapshot"),
    ],
)
def test_split_mode(live_split, size_gb, free_gb, mapi_ok, mode):
    got = rules.split_mode(
        live_split,
        int(size_gb * GB),
        int(free_gb * GB),
        mapi_ok,
        budget_bytes=10 * GB,
        disk_factor=2.5,
        reserve_bytes=5 * GB,
    )
    assert got == mode
