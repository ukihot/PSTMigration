"""ライブ分割方式の束（チャンク）管理（設計書 §16）。Outlook に依存しない部分。"""

from __future__ import annotations

from collections import Counter, defaultdict
from collections.abc import Callable
from pathlib import Path

from pstmigration import transfer
from pstmigration.state import StepError

GB = 1024**3
BUCKETS = ("new", "old")
VERIFIED, BUILT, UPLOADED = "VERIFIED", "BUILT", "UPLOADED"
SEVENZIP_ROOM = 1.01  # 7z を作るのに要る空き（束PSTの何倍）


def chunk_name(prefix: str, seq: int, ext: str) -> str:
    return f"{prefix}_{seq:03d}{ext}"


def chunk_target(free: int, *, max_bytes: int, min_bytes: int, reserve_bytes: int, factor: float) -> int | None:
    """束の目標サイズ（§16.5）。空きが足りず下限を割るなら None（始めない）。"""
    target = min(max_bytes, int((free - reserve_bytes) / factor))
    return target if target >= min_bytes else None


def next_batch(items: list[dict], start: int, target_bytes: int, max_items: int) -> int:
    """1回の CopyMessages に入れる範囲の終端。PR_MESSAGE_SIZE の合計で目標を超えない所まで（最低1通）。"""
    end, total = start, 0
    while end < len(items) and end - start < max_items:
        size = items[end]["s"]
        if end > start and total + size > target_bytes:
            break
        total += size
        end += 1
    return end


def uploaded_until(records: list[dict], bucket: str) -> int:
    """先頭から隙間なく UPLOADED が続く所の終端（＝次の束の開始位置）。"""
    pos = 0
    for r in sorted((r for r in records if r["bucket"] == bucket and r["status"] == UPLOADED), key=_start):
        if r["range"][0] != pos:
            break
        pos = r["range"][1]
    return pos


def completion_problems(records: list[dict], counts: dict[str, int]) -> list[str]:
    """§16.9 の1・2：各側の束が計画を重複なく・漏れなく覆い、全部 UPLOADED か。空なら合格。"""
    problems = []
    for bucket in BUCKETS:
        rs = sorted((r for r in records if r["bucket"] == bucket), key=_start)
        pos = 0
        for r in rs:
            start, end = r["range"]
            if r["status"] != UPLOADED:
                problems.append(f"{bucket} #{r['seq']} が未転送（{r['status']}）")
            if end <= start:
                problems.append(f"{bucket} #{r['seq']} が空")
            if start != pos:
                problems.append(f"{bucket} {'欠落' if start > pos else '重複'}: {min(start, pos)}〜{max(start, pos)}")
            pos = max(pos, end)
        if pos != counts.get(bucket, 0):
            problems.append(f"{bucket} 範囲の終端 {pos} が計画の件数 {counts.get(bucket, 0)} と違う")
    return problems


def _start(r: dict) -> int:
    return r["range"][0]


# ---- 照合（§16.6） ----------------------------------------------------------------
# 行は {"f": フォルダー, "k": SearchKey, "p": 指紋（件名・日時・クラス・添付有無）, "h": 本文・添付のハッシュ}


def hash_needed(source_rows: list[dict]) -> set[tuple]:
    """元ストアの全行（全フォルダー・計画外も含む）で「SearchKey＋指紋」が重複する組。本文・添付のハッシュで区別する。

    フォルダーは見ない。PST の EntryID はストア全体で一意で、別フォルダーのメールも計画したフォルダーへ
    コピーされうるため（取り違えの相手はどのフォルダーからでも来る）。
    """
    counts = Counter((r["k"], r["p"]) for r in source_rows)
    return {i for i, n in counts.items() if n > 1}


def _match_key(r: dict, need: set[tuple]) -> tuple:
    return r["f"], r["k"], r["p"], r.get("h") if (r["k"], r["p"]) in need else None


def size_deltas(expected: list[dict], actual: list[dict], need: set[tuple]) -> dict[str, int]:
    """PR_MESSAGE_SIZE の差（束 − 元）の分布。照合キーが同じものどうしを小さい順に組にする。

    合否には使わず記録する（2026-10-22 の実測は全件 +4 バイト。固定の仕様にはしない。§16.6）。
    """
    want, got = defaultdict(list), defaultdict(list)
    for r in expected:
        want[_match_key(r, need)].append(r["s"])
    for r in actual:
        got[_match_key(r, need)].append(r["s"])
    c = Counter()
    for k, sizes in want.items():
        for a, b in zip(sorted(sizes), sorted(got.get(k, [])), strict=False):
            c[str(b - a)] += 1
    return dict(c)


def mismatches(expected: list[dict], actual: list[dict], need: set[tuple]) -> list[str]:
    """期待した行と束の中の行の「フォルダー＋SearchKey＋指紋（＋必要ならハッシュ）」の多重集合の差。空なら一致。

    ハッシュが要る組でハッシュが無い行は一致しない扱い（計算できなかったものを合格にしない）。
    """

    def key(r: dict) -> tuple:
        return _match_key(r, need)

    no_hash = [
        f"ハッシュ未取得（{side}） {r['f']} key={r['k']}"
        for side, rows in (("期待", expected), ("束", actual))
        for r in rows
        if (r["k"], r["p"]) in need and r.get("h") is None
    ]
    want, got = Counter(map(key, expected)), Counter(map(key, actual))
    return (
        no_hash
        + [f"欠落 {k[0]} key={k[1]} ×{n}" for k, n in (want - got).items()]
        + [f"余計 {k[0]} key={k[1]} ×{n}" for k, n in (got - want).items()]
    )


# ---- 束の作成・転送・再開（§16.5、§16.8） ------------------------------------------
#
# ops は次を持つオブジェクト（本番は MAPI・7z・NAS、テストは偽物）:
#   free_bytes() / file_size(path) / open_chunk(path, first) -> writer(.copy(items), .close())
#     first は各側の最初の束（_001）。空フォルダーを含む全階層をそこに作る
#   rows(path, need) -> 束の中の行 / pack(pst) -> 7z（7z t 合格後に pst を消す）
#   remote_exists(name) / remote_sha(name) / remote_delete(name)（正式名と .partial）/ upload(local, name, sha)
# limits は bytes: max / min / reserve、比率: factor / plan_factor、1回のコピー: batch（bytes）/ batch_items


def run_bucket(
    bucket: str,
    items: list[dict],
    records: list[dict],
    ops,
    save: Callable[[], None],
    *,
    work: Path,
    prefix: str,
    need: set[tuple],
    limits: dict,
    on_uploaded: Callable[[dict], None] | None = None,
) -> None:
    """1つの側（new / old）を、束ごとに 作成→照合→（旧側は7z）→転送 まで進める。

    束は1つずつ仕上げる（UPLOADED になるまで次を作らない）ので、未完了の束は各側に高々1つ。
    中断後は、その束を §16.8 の判定順で再開する。
    on_uploaded は束を UPLOADED として保存した後に呼ぶ（元PSTの変更検知。例外を投げれば以降を止める）。
    """
    side = _Side(bucket, items, records, ops, save, work, prefix, need, limits, on_uploaded)
    for r in side.mine():
        if r["status"] == UPLOADED:
            side.remove_local(r)  # 前回、ローカルを消す前に落ちた分
    while True:
        pending = [r for r in side.mine() if r["status"] != UPLOADED]
        if len(pending) > 1:
            raise StepError("STATE_CORRUPT", f"{bucket} に未完了の束が複数ある", final=True)
        if pending:
            side.resume(pending[0])
            continue
        start = uploaded_until(records, bucket)
        if start >= len(items):
            return
        side.build(start)


class _Side:
    def __init__(self, bucket, items, records, ops, save, work, prefix, need, limits, on_uploaded):
        self.bucket, self.items, self.records, self.ops, self.save = bucket, items, records, ops, save
        self.work, self.prefix, self.need, self.limits = work, prefix, need, limits
        self.on_uploaded = on_uploaded

    def mine(self) -> list[dict]:
        return [r for r in self.records if r["bucket"] == self.bucket]

    def pst(self, seq: int) -> Path:
        return self.work / chunk_name(self.prefix, seq, ".pst")

    def upload_name(self, seq: int) -> str:
        return chunk_name(self.prefix, seq, ".7z" if self.bucket == "old" else ".pst")

    def remove_local(self, r: dict) -> None:
        for p in (self.pst(r["seq"]), self.work / r["name"]):
            p.unlink(missing_ok=True)

    def check(self, r_range: list[int], pst: Path) -> tuple[list[str], dict]:
        """束PSTの中身を計画の範囲と照合する。戻り値: (不一致の説明, サイズ差の分布)。"""
        a, b = r_range
        rows = self.ops.rows(pst, self.need)
        return mismatches(self.items[a:b], rows, self.need), size_deltas(self.items[a:b], rows, self.need)

    # §16.8 の判定順
    def resume(self, r: dict) -> None:
        ops, name = self.ops, r["name"]
        if r.get("sha256") and ops.remote_sha(name) == r["sha256"]:  # 2
            self.done(r)
            return
        if ops.remote_exists(name):  # 3
            ops.remote_delete(name)
        local = self.work / name
        if r["status"] == BUILT and local.exists() and transfer.sha256_file(local) == r["sha256"]:  # 4
            self.upload(r)
            return
        pst = self.pst(r["seq"])
        if r["status"] == VERIFIED and pst.exists() and not self.check(r["range"], pst)[0]:  # 5
            self.finish(r)
            return
        self.discard(r)  # 6：捨てて作り直す（次のループで同じ位置から）

    def discard(self, r: dict) -> None:
        self.ops.remote_delete(r["name"])
        self.remove_local(r)
        self.records.remove(r)
        self.save()

    def build(self, start: int) -> None:
        ops, lim = self.ops, self.limits
        target = chunk_target(
            ops.free_bytes(),
            max_bytes=lim["max"],
            min_bytes=lim["min"],
            reserve_bytes=lim["reserve"],
            factor=lim["factor"],
        )
        if target is None:
            raise StepError("DISK_SPACE", f"空き {ops.free_bytes() / GB:.1f}GB では束を作れない", transient=True)
        seq = max((r["seq"] for r in self.mine()), default=0) + 1
        pst = self.pst(seq)
        for stale in (pst, pst.with_suffix(".7z")):  # 記録の無い作りかけ（前回この束の途中で落ちた）
            stale.unlink(missing_ok=True)
        end = self.copy_until_full(pst, start, target, first=seq == 1)
        # 閉じると実ファイルサイズが増えうる。閉じた後の最終サイズで空きを確かめ直す（Windows 自体を守る）
        pst_size = ops.file_size(pst)
        if ops.free_bytes() < lim["reserve"]:
            pst.unlink(missing_ok=True)
            raise StepError(
                "DISK_SPACE",
                f"{pst.name} を閉じた後の空き {ops.free_bytes() / GB:.1f}GB が予備を下回ったため捨てた",
                transient=True,
            )
        bad, deltas = self.check([start, end], pst)
        if bad:
            pst.unlink(missing_ok=True)
            raise StepError("CHUNK_MISMATCH", f"{pst.name}: " + " / ".join(bad[:5]), final=True)
        r = {
            "bucket": self.bucket,
            "seq": seq,
            "name": self.upload_name(seq),
            "range": [start, end],
            "status": VERIFIED,
            "pst_size": pst_size,
            "size_delta": deltas,
        }
        self.records.append(r)
        self.save()
        self.finish(r)

    def copy_until_full(self, pst: Path, start: int, target: int, *, first: bool) -> int:
        """束の実ファイルサイズと空き容量を見ながらコピーし、閉じる位置（終端）を返す。最低1回はコピーする。"""
        ops, lim, items = self.ops, self.limits, self.items
        end = start
        writer = ops.open_chunk(pst, first)
        try:
            while end < len(items):
                stop = next_batch(items, end, lim["batch"], lim["batch_items"])
                est = int(sum(i["s"] for i in items[end:stop]) * lim["plan_factor"])
                if end > start and ops.file_size(pst) + est > target:
                    break
                if ops.free_bytes() - est < lim["reserve"]:
                    if end == start:
                        raise StepError("DISK_SPACE", f"空き {ops.free_bytes() / GB:.1f}GB", transient=True)
                    break
                writer.copy(items[end:stop])
                end = stop
                if ops.file_size(pst) >= target:
                    break
        finally:
            writer.close()
        return end

    def finish(self, r: dict) -> None:
        """照合済み（VERIFIED）の束を、転送するファイルにして（旧側は7z）SHA-256 を記録し、転送する。"""
        ops, pst, old = self.ops, self.pst(r["seq"]), self.bucket == "old"
        # 旧側：7z（束PSTと同程度。圧縮できないデータでもわずかに大きい程度）を作る空きがなければ、束を捨てて空ける
        if old and ops.free_bytes() - int(ops.file_size(pst) * SEVENZIP_ROOM) < self.limits["reserve"]:
            self.discard(r)
            raise StepError(
                "DISK_SPACE",
                f"{pst.name} を 7z にする空きが無いため捨てた（次回小さい束で作り直す）",
                transient=True,
            )
        local = ops.pack(pst) if old else pst
        r.update(status=BUILT, sha256=transfer.sha256_file(local), size=local.stat().st_size)
        self.save()
        self.upload(r)

    def upload(self, r: dict) -> None:
        self.ops.upload(self.work / r["name"], r["name"], r["sha256"])
        self.done(r)

    def done(self, r: dict) -> None:
        r["status"] = UPLOADED
        self.save()
        self.remove_local(r)
        if self.on_uploaded:
            self.on_uploaded(r)


# ---- 元PSTの変更検知（§16.7） ------------------------------------------------------------


def source_changes(plan: list[dict], scanned_eids: set[str], current: list[dict]) -> tuple[list[str], dict[str, int]]:
    """計画（旧・新の全メール）と、いま元PSTを調べ直した全行を比べる。

    戻り値: (FAILED にする変化, 調査後に増えたメールのフォルダー別件数)
    計画したメールが計画したフォルダーに無い（削除・移動）、SearchKey が変わった → FAILED。
    既読・フラグ・分類の変更は SearchKey にも指紋にも入らないので、ここには出ない（許容）。
    調査時点に無かったメールは移行対象外。件数だけ返す（SOURCE_CHANGED 警告・再回収の候補）。
    """
    now = {r["e"]: r for r in current}
    failures = []
    for i in plan:
        r = now.get(i["e"])
        if r is None:
            failures.append(f"消えた {i['f']} e={i['e'][:16]}")
        elif r["f"] != i["f"]:
            failures.append(f"移された {i['f']} → {r['f']} e={i['e'][:16]}")
        elif r["k"] != i["k"]:
            failures.append(f"SearchKey が変わった {i['f']} e={i['e'][:16]}")
    added = Counter(r["f"] for r in current if r["e"] not in scanned_eids)
    return failures, dict(added)
