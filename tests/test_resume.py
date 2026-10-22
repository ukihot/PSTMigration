"""束の状態遷移・中断・再開（設計書 §16.8、§16.9）。

MAPI は偽物（束PSTを JSON で表す。作り直すたびに乱数が入り、SHA-256 が変わる＝本物の PST と同じ性質）。
NAS はローカルフォルダーで、転送は本物の transfer.upload を使う。
「中断」は BaseException を投げて表す（except Exception では捕まらない）。finally は走るので、
finally が走らない本当のプロセス強制終了は、隔離PSTの統合試験で別に確かめる。
"""

import json
import os
import shutil
import tempfile
from collections import Counter
from pathlib import Path

import pytest
from hypothesis import HealthCheck, settings
from hypothesis import strategies as st
from hypothesis.stateful import RuleBasedStateMachine, initialize, invariant, precondition, rule

from pstmigration import chunks, transfer
from pstmigration.state import StepError, read_json, write_json

LIMITS = {"max": 600, "min": 50, "reserve": 1000, "factor": 2.5, "batch": 120, "batch_items": 3, "plan_factor": 1.15}
PREFIX = {"new": "new_mail_from_20160221_PC", "old": "old_mail_before_20160221_PC"}
BIG_DISK = 10**9


class Crash(BaseException):
    """プロセスが落ちた。"""


def decode(path: Path) -> list[tuple]:
    """オラクル用：偽PST・偽7zに入っているメールの中身（実装の照合関数は使わない）。"""
    b = path.read_bytes()
    if b.startswith(b"7z:"):
        b = b[3:]
    return [(r["f"], r["k"], r["p"], r["c"]) for r in json.loads(b)["rows"]]


class World:
    """偽の MAPI・7z・NAS と state.json。ops として chunks.run_bucket に渡す。"""

    def __init__(self, root: Path, plan: dict[str, list[dict]]):
        self.work, self.nas, self.state = root / "work", root / "nas", root / "state.json"
        self.work.mkdir()
        self.nas.mkdir()
        self.plan = plan
        self.need = chunks.hash_needed(plan["new"] + plan["old"])
        for items in plan.values():
            for it in items:
                it["h"] = it["c"] if (it["k"], it["p"]) in self.need else None
        self.countdown = None  # 設定してから n 回目の操作で落ちる
        self.crash_on = None  # (操作名, n)：設定してからその操作の n 回目で落ちる
        self.capacity = BIG_DISK  # 空き容量 = 容量 − 作業フォルダーの使用量
        self.close_growth = 0  # 束PSTを閉じたときに増えるバイト数
        self.glitch = False
        self.glitched = False  # 取り違えを実際に起こしたか
        self.calls = Counter()
        self.opened = []  # (束PST名, 最初の束か)
        self.on_uploaded = None  # 束を送り終えるたびに呼ぶ

    def tick(self, what: str) -> None:
        self.calls[what] += 1
        if self.countdown is not None:
            self.countdown -= 1
            if self.countdown <= 0:
                self.countdown = None
                raise Crash(what)
        if self.crash_on and self.crash_on[0] == what:
            name, n = self.crash_on
            self.crash_on = (name, n - 1) if n > 1 else None
            if n <= 1:
                raise Crash(what)

    # ---- ops ----
    def free_bytes(self) -> int:
        return self.capacity - sum(p.stat().st_size for p in self.work.iterdir())

    def file_size(self, path: Path) -> int:
        return path.stat().st_size if path.exists() else 0

    def open_chunk(self, path: Path, first: bool):
        self.opened.append((path.name, first))
        return Writer(self, path)

    def rows(self, path: Path, need: set) -> list[dict]:
        self.tick("rows")
        data = json.loads(path.read_bytes())
        return [{**r, "h": r["c"] if (r["k"], r["p"]) in need else None} for r in data["rows"]]

    def pack(self, pst: Path) -> Path:
        self.tick("pack")
        arc = pst.with_suffix(".7z")
        arc.write_bytes(b"7z:" + pst.read_bytes())
        self.tick("pack-written")
        pst.unlink()
        return arc

    def remote_exists(self, name: str) -> bool:
        return (self.nas / name).exists()

    def remote_sha(self, name: str) -> str | None:
        p = self.nas / name
        return transfer.sha256_file(p) if p.exists() else None

    def remote_delete(self, name: str) -> None:
        self.tick("remote_delete")
        for p in (self.nas / name, self.nas / (name + ".partial")):
            p.unlink(missing_ok=True)

    def upload(self, local: Path, name: str, sha: str) -> None:
        self.tick("upload")
        part = self.nas / (name + ".partial")
        if not part.exists():  # 途中まで書いたところで落ちうる
            data = local.read_bytes()
            part.write_bytes(data[: len(data) // 2])
        self.tick("upload-partial")
        transfer.upload(local, self.nas / name, sha, 0)
        self.tick("upload-renamed")  # 正式名にした直後（state 保存前）

    def save(self, records: list[dict]) -> None:
        self.tick("save")
        write_json(self.state, {"chunks": records})
        self.tick("saved")

    def load(self) -> list[dict]:
        return read_json(self.state, {"chunks": []})["chunks"]

    def run(self) -> list[dict]:
        """プロセス1回分。state.json から作り直す（メモリ上の状態は持ち越さない）。"""
        records = self.load()
        for b in ("new", "old"):
            chunks.run_bucket(
                b,
                self.plan[b],
                records,
                self,
                lambda r=records: self.save(r),
                work=self.work,
                prefix=PREFIX[b],
                need=self.need,
                limits=LIMITS,
                on_uploaded=self.on_uploaded,
            )
        return records

    # ---- オラクル（実装の判定を使わない） ----
    def nas_contents(self, bucket: str) -> Counter:
        c = Counter()
        for p in self.nas.glob(PREFIX[bucket] + "_*"):
            if not p.name.endswith(".partial"):
                c.update(decode(p))
        return c

    def planned(self, bucket: str) -> Counter:
        return Counter((i["f"], i["k"], i["p"], i["c"]) for i in self.plan[bucket])

    def assert_completed_correctly(self, records: list[dict]) -> None:
        for b in ("new", "old"):
            assert self.nas_contents(b) == self.planned(b), f"{b}: NAS の中身が計画と違う"
        for r in records:
            assert transfer.sha256_file(self.nas / r["name"]) == r["sha256"]
        assert not list(self.work.glob("*_mail_*")), "完了後にローカルの束が残っている"


class Writer:
    def __init__(self, world: World, path: Path):
        self.w, self.path = world, path
        world.tick("open")
        if not path.exists():
            path.write_text(json.dumps({"nonce": os.urandom(8).hex(), "rows": []}))

    def copy(self, batch: list[dict]) -> None:
        w = self.w
        w.tick("copy")
        rows = [dict(i) for i in batch]
        if w.glitch and len(rows) >= 2:  # CopyMessages の取り違え：A・B のつもりが A・A
            rows[-1] = rows[0]
            w.glitch, w.glitched = False, True
        data = json.loads(self.path.read_text())
        half = len(rows) // 2
        data["rows"] += rows[:half]
        self.path.write_text(json.dumps(data))
        w.tick("copy-half")  # 一部だけコピーしたところで落ちうる
        data["rows"] += rows[half:]
        self.path.write_text(json.dumps(data))

    def close(self) -> None:
        if self.w.close_growth:  # 閉じたときに実ファイルサイズが増える
            data = json.loads(self.path.read_text())
            data["pad"] = data.get("pad", "") + "x" * self.w.close_growth
            self.path.write_text(json.dumps(data))


def make_plan(n_new: int, n_old: int) -> dict[str, list[dict]]:
    def items(b, n):
        return [{"e": f"{b}{i}", "f": "受信", "k": f"K{i}", "p": "P", "c": f"c{i}", "s": 40} for i in range(n)]

    return {"new": items("new", n_new), "old": items("old", n_old)}


@pytest.fixture
def world(tmp_path):
    return World(tmp_path, make_plan(7, 7))


# ---- §16.8 の判定表を1行ずつ（決まった位置で中断してから再開） ----------------------


def resume(world: World) -> list[dict]:
    world.calls.clear()
    records = world.run()
    assert chunks.completion_problems(records, {b: len(v) for b, v in world.plan.items()}) == []
    world.assert_completed_correctly(records)
    return records


def interrupt(world: World, op: str, n: int = 1) -> None:
    world.crash_on = (op, n)
    with pytest.raises(Crash):
        world.run()


def test_happy_path_makes_multiple_chunks_per_side(world):
    records = resume(world)
    assert sum(r["bucket"] == "new" for r in records) >= 2  # 束が複数に分かれている
    assert all(r["name"].endswith(".7z") for r in records if r["bucket"] == "old")


def test_only_first_chunk_of_each_side_is_opened_as_first(world):
    """空フォルダーを含む全階層は各側の最初の束（_001）に作る（§16.5）。作り直しても _001 は first。"""
    interrupt(world, "copy-half")  # new_001 の途中で落ちる → 作り直し
    resume(world)
    assert world.opened and all(first == name.endswith("_001.pst") for name, first in world.opened)
    assert {n for n, f in world.opened if f} == {chunks.chunk_name(PREFIX[b], 1, ".pst") for b in ("new", "old")}


def test_row1_uploaded_chunks_are_not_redone(world):
    resume(world)
    resume(world)
    assert world.calls["copy"] == 0 and world.calls["upload"] == 0


def test_row2_built_and_correct_file_on_nas_is_promoted_even_if_local_is_lost(world):
    interrupt(world, "upload-renamed")  # NAS に正式名で置いた直後、UPLOADED を記録する前に落ちた
    (rec,) = [r for r in world.load() if r["status"] == "BUILT"]
    (world.work / rec["name"]).unlink()  # ローカルの束も消えた
    resume(world)
    assert world.calls["upload"] == len(world.load()) - 1  # その束は送り直していない
    assert world.calls["open"] == len(world.load()) - 1  # 作り直してもいない


def test_row3_unconfirmed_file_on_nas_is_replaced(world):
    interrupt(world, "upload-renamed")
    (rec,) = [r for r in world.load() if r["status"] == "BUILT"]
    (world.nas / rec["name"]).write_bytes(b"not ours")  # 記録と違う同名ファイル
    resume(world)


def test_row4_built_with_local_file_is_uploaded_without_rebuilding(world):
    interrupt(world, "upload-partial")  # .partial を半分書いたところで落ちた
    resume(world)
    assert world.calls["open"] == len(world.load()) - 1


def uploaded(records):
    return [r for r in records if r["status"] == "UPLOADED"]


def test_row5_verified_old_chunk_is_reverified_then_packed(world):
    interrupt(world, "pack")  # 旧側：照合（VERIFIED）の後、7z の前に落ちた
    before = world.load()
    (rec,) = [r for r in before if r["status"] == "VERIFIED"]
    assert rec["bucket"] == "old"
    after = resume(world)
    assert world.calls["rows"] >= 1  # 前回の VERIFIED を信用せず再照合した
    # 作ったのは、中断時に UPLOADED でも VERIFIED でもなかった束だけ（VERIFIED の束は作り直していない）
    assert world.calls["open"] == len(after) - len(uploaded(before)) - 1
    assert any(r["seq"] == rec["seq"] and r["range"] == rec["range"] for r in after)


def test_row5_to_6_verified_chunk_that_fails_reverification_is_rebuilt(world):
    interrupt(world, "pack")
    before = world.load()
    (rec,) = [r for r in before if r["status"] == "VERIFIED"]
    pst = world.work / chunks.chunk_name(PREFIX["old"], rec["seq"], ".pst")
    data = json.loads(pst.read_text())
    data["rows"].pop()  # 照合後に壊れた
    pst.write_text(json.dumps(data))
    after = resume(world)
    # VERIFIED だった束も作り直した（中断時に UPLOADED でなかった束は全部作った）
    assert world.calls["open"] == len(after) - len(uploaded(before))


def test_row6_built_chunk_lost_locally_with_partial_on_nas_is_rebuilt_cleanly(world):
    """NAS に前の束の .partial が残ったまま作り直すと、次の転送が HASH_MISMATCH で回数を消費する（M3 で発覚）。"""
    interrupt(world, "upload-partial")  # BUILT、NAS には .partial が半分
    (rec,) = [r for r in world.load() if r["status"] == "BUILT"]
    (world.work / rec["name"]).unlink()  # ローカルの束が消えた → 作り直し
    resume(world)  # HASH_MISMATCH などのエラーなしで完了すること


def test_row6_verified_chunk_without_local_file_is_rebuilt(world):
    interrupt(world, "pack")
    for p in world.work.iterdir():
        p.unlink()
    resume(world)


def test_partial_copy_crash_leaves_no_duplicates(world):
    interrupt(world, "copy-half", 2)  # 束の途中、一部だけコピーしたところで落ちた
    resume(world)


def test_wrong_copy_is_failed_not_completed(world):
    world.glitch = True
    with pytest.raises(StepError) as e:
        world.run()
    assert e.value.code == "CHUNK_MISMATCH" and e.value.final
    assert not list(world.nas.glob("*_mail_*"))  # 取り違えた束は NAS に送らない


def test_disk_shortage_is_transient_and_resumable(world):
    world.capacity = LIMITS["reserve"] + 10
    with pytest.raises(StepError) as e:
        world.run()
    assert e.value.code == "DISK_SPACE" and e.value.transient
    world.capacity = BIG_DISK
    resume(world)


def test_final_pst_size_after_close_is_recorded(world):
    world.close_growth = 50
    records = resume(world)
    for r in records:  # 閉じた後（増えた後）の実サイズ。新側はそのまま送り、旧側は偽7z（3バイトの見出し付き）
        assert r["pst_size"] >= 50
        assert r["size"] == r["pst_size"] + (3 if r["bucket"] == "old" else 0)


@pytest.mark.parametrize("bucket", ["new", "old"])
def test_chunk_that_leaves_disk_below_reserve_after_close_is_discarded(tmp_path, bucket):
    # 束の目標 600B が取れる空き（1000 + 2.5×600）。閉じたときに 1000B 増えて、空きが予備（1000B）を割る
    w = World(tmp_path, make_plan(7, 0) if bucket == "new" else make_plan(0, 7))
    w.capacity, w.close_growth = LIMITS["reserve"] + int(LIMITS["factor"] * LIMITS["max"]), 1000
    with pytest.raises(StepError) as e:
        w.run()
    assert e.value.code == "DISK_SPACE" and e.value.transient
    assert not list(w.work.iterdir()), "空きを割った束はローカルから消して、領域を空ける"
    assert not [r for r in w.load() if r["status"] != "UPLOADED"]
    w.capacity, w.close_growth = BIG_DISK, 0
    resume(w)


def test_old_chunk_without_room_for_7z_is_discarded_before_packing(tmp_path):
    # 閉じた後の空きは予備を割らないが、7z（束PST 1個分）を作る空きは無い
    w = World(tmp_path, make_plan(0, 7))
    w.capacity, w.close_growth = LIMITS["reserve"] + int(LIMITS["factor"] * LIMITS["max"]), 300
    with pytest.raises(StepError) as e:
        w.run()
    assert e.value.code == "DISK_SPACE"
    assert w.calls["pack"] == 0, "7z を作り始めない"
    assert not list(w.work.iterdir())
    w.capacity, w.close_growth = BIG_DISK, 0
    resume(w)


def test_single_mail_larger_than_target_gets_its_own_chunk(tmp_path):
    plan = make_plan(3, 0)
    plan["new"][1]["s"] = 10_000  # 1通で束の目標を超える
    plan["new"][1]["c"] = "x" * 2000
    w = World(tmp_path, plan)
    records = resume(w)
    assert [r["range"] for r in records] == [[0, 1], [1, 2], [2, 3]]


# ---- Hypothesis：中断・消失・破損・容量不足・取り違えを任意の順で ----------------------

CRASH_POINTS = [
    "open",
    "copy",
    "copy-half",
    "rows",
    "pack",
    "pack-written",
    "remote_delete",
    "upload",
    "upload-partial",
    "upload-renamed",
    "save",
    "saved",
]

_items = st.lists(
    st.tuples(
        st.sampled_from(["受信", "送信"]),
        st.sampled_from(["X", "Y", "Z"]),
        st.sampled_from(["A", "B"]),
        st.sampled_from(["c1", "c2", "c3"]),
        st.integers(20, 80),
    ),
    max_size=16,
)


class ResumeMachine(RuleBasedStateMachine):
    @initialize(new=_items, old=_items)
    def setup(self, new, old):
        self.root = Path(tempfile.mkdtemp(prefix="pstm_resume_"))
        plan = {
            b: [{"e": f"{b}{i}", "f": f, "k": k, "p": p, "c": c, "s": s} for i, (f, k, p, c, s) in enumerate(lst)]
            for b, lst in (("new", new), ("old", old))
        }
        self.w = World(self.root, plan)
        self.counts = {b: len(v) for b, v in plan.items()}
        self.failed = False
        self.damaged = False  # NAS の .partial を壊した（次の転送で HASH_MISMATCH になりうる）

    def teardown(self):
        shutil.rmtree(getattr(self, "root", ""), ignore_errors=True)

    @precondition(lambda self: not self.failed)
    @rule(
        crash_after=st.none() | st.integers(1, 40),
        crash_on=st.none() | st.tuples(st.sampled_from(CRASH_POINTS), st.integers(1, 3)),
    )
    def run_process(self, crash_after, crash_on):
        self.w.countdown = crash_after
        self.w.crash_on = crash_on
        try:
            records = self.w.run()
        except Crash:
            return
        except StepError as e:
            # エラーは、それを起こす操作をした後でしか許さない（原因の無い失敗・回数の消費は不具合）
            allowed = set()
            if self.w.capacity < BIG_DISK or self.w.close_growth:
                allowed.add("DISK_SPACE")
            if self.damaged:
                allowed.add("HASH_MISMATCH")
            if self.w.glitched:
                allowed.add("CHUNK_MISMATCH")
            assert e.code in allowed, f"原因の無いエラー {e.code}: {e}"
            if e.code == "HASH_MISMATCH":
                self.damaged = False  # 壊れた .partial は消されている
            self.failed = e.final  # 取り違えは FAILED（以後は自動で再実行しない）
            return
        finally:
            self.w.countdown = self.w.crash_on = None
        # 例外なく戻ったなら完了している（進行性）。正しさはオラクルで確かめる
        assert chunks.completion_problems(records, self.counts) == []
        self.w.assert_completed_correctly(records)

    @rule(data=st.data())
    def lose_local_file(self, data):
        files = sorted(self.w.work.iterdir())
        if files:
            data.draw(st.sampled_from(files)).unlink()

    @rule(data=st.data())
    def damage_nas_partial(self, data):
        parts = sorted(self.w.nas.glob("*.partial"))
        if parts:
            data.draw(st.sampled_from(parts)).write_bytes(b"garbage")
            self.damaged = True

    @rule(data=st.data())
    def damage_local_chunk(self, data):
        """照合の後にローカルの束PSTが壊れた（1通欠ける）。"""
        psts = [p for p in sorted(self.w.work.glob("*.pst")) if json.loads(p.read_text())["rows"]]
        if psts:
            p = data.draw(st.sampled_from(psts))
            d = json.loads(p.read_text())
            d["rows"].pop()
            p.write_text(json.dumps(d))

    @rule(data=st.data())
    def tamper_unconfirmed_nas_file(self, data):
        """UPLOADED と記録されていない束と同名の、中身の違う正式ファイルが NAS にある。"""
        names = sorted(r["name"] for r in self.w.load() if r["status"] != "UPLOADED")
        if names:
            name = data.draw(st.sampled_from(names))
            (self.w.nas / name).write_text(json.dumps({"nonce": "other", "rows": []}))

    @rule(low=st.booleans())
    def disk(self, low):
        self.w.capacity = LIMITS["reserve"] + int(LIMITS["factor"] * LIMITS["max"]) if low else BIG_DISK

    @rule(growth=st.sampled_from([0, 0, 300, 1000]))
    def pst_grows_on_close(self, growth):
        self.w.close_growth = growth

    @rule()
    def glitch_next_copy(self):
        self.w.glitch = True

    @invariant()
    def nas_never_holds_unplanned_or_duplicated_mail(self):
        for b in ("new", "old"):
            extra = self.w.nas_contents(b) - self.w.planned(b)
            assert not extra, f"{b}: 計画に無い・計画より多いメールが NAS にある {extra}"

    @invariant()
    def uploaded_ranges_never_overlap_and_stay_in_plan(self):
        for b in ("new", "old"):
            covered = Counter(
                i
                for r in self.w.load()
                if r["bucket"] == b and r["status"] == "UPLOADED"
                for i in range(r["range"][0], r["range"][1])
            )
            assert all(n == 1 for n in covered.values())
            assert all(0 <= i < self.counts[b] for i in covered)

    @invariant()
    def completion_claim_is_true(self):
        records = self.w.load()
        if chunks.completion_problems(records, self.counts) == []:
            for b in ("new", "old"):
                assert self.w.nas_contents(b) == self.w.planned(b)
            for r in records:
                assert transfer.sha256_file(self.w.nas / r["name"]) == r["sha256"]


ResumeMachine.TestCase.settings = settings(
    max_examples=300,
    stateful_step_count=30,
    deadline=None,
    suppress_health_check=[HealthCheck.too_slow],
)
TestResumeMachine = ResumeMachine.TestCase


# ---- 束ごとの後処理（元PSTの変更検知に使う。§16.7） ----------------------------------------


def test_on_uploaded_is_called_after_each_chunk_is_saved_as_uploaded(world):
    seen = []

    def check(r):
        assert r["status"] == "UPLOADED"
        assert any(
            x["seq"] == r["seq"] and x["status"] == "UPLOADED" for x in world.load() if x["bucket"] == r["bucket"]
        )
        seen.append(r["name"])

    world.on_uploaded = check
    records = resume(world)
    assert seen == [r["name"] for r in records]


def test_failure_raised_after_a_chunk_stops_processing_and_keeps_its_record(world):
    def stop(r):
        raise StepError("SOURCE_CHANGED_DURING_COPY", "消えた", final=True)

    world.on_uploaded = stop
    with pytest.raises(StepError):
        world.run()
    assert [r["status"] for r in world.load()] == ["UPLOADED"]  # 1束目は送信済みとして残り、2束目は作らない


def test_each_chunk_records_size_delta_distribution(world):
    """§16.6：サイズ差（束 − 元）の分布を束ごとに記録する。偽の束はサイズをそのまま写すので差は 0。"""
    for r in resume(world):
        assert r["size_delta"] == {"0": r["range"][1] - r["range"][0]}


def test_row4_to_6_built_chunk_damaged_locally_is_rebuilt_not_sent(world):
    """BUILT の束が、SHA-256 を記録した後にローカルで壊れた。そのまま送らず作り直す（M9 で発覚した試験の穴）。"""
    interrupt(world, "upload")  # BUILT を保存した直後、送る前に落ちた
    before = world.load()
    (rec,) = [r for r in before if r["status"] == "BUILT"]
    (world.work / rec["name"]).write_bytes(b"damaged")
    after = resume(world)
    assert world.calls["open"] == len(after) - len(uploaded(before))  # BUILT だった束も作り直した
