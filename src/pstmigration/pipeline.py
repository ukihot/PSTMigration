"""ログオン時・ログオフ時の処理本体（設計書 §2.3、§4、§6〜§9）。"""

from __future__ import annotations

import contextlib
import gc
import logging
import msvcrt
import os
import shutil
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path, PureWindowsPath

import pywintypes
import win32con
import win32file

from pstmigration import report, rules, transfer
from pstmigration.config import Config
from pstmigration.state import (
    TERMINAL,
    State,
    StepError,
    apply_error,
    new_entry,
    now_iso,
    read_json,
    warn,
    write_json,
)

log = logging.getLogger(__name__)

GB = 1024**3
SNAPSHOT = "snapshot.pst"
SCAN = "scan.json"
LEGACY_CUTOFF = "2021-01-01T00:00:00+09:00"  # 1.0.1 までは PST ごとに記録していない。当時配布した config の値
CONTACTS_DIR = "contacts"  # 作業フォルダー内。PSTごとのフォルダー名（12桁の16進）とは重ならない
SNAPSHOT_LABEL = "PSTMigration作業中（操作しないでください）"


def _user() -> str:
    return os.environ.get("USERNAME", "unknown")


def _computer() -> str:
    return os.environ.get("COMPUTERNAME", "unknown")


def _owner(data: dict) -> str:
    """NAS 上のユーザーフォルダー名。メールアドレスが分かればそれ、なければ Windows ユーザー名。"""
    return data.get("mailbox") or _user()


def _ad_mail() -> str | None:
    """ログオンユーザーの AD の mail 属性（DCに届かない・未設定なら None）。"""
    import win32com.client

    try:
        dn = win32com.client.Dispatch("ADSystemInfo").UserName
        return win32com.client.GetObject("LDAP://" + dn.replace("/", "\\/")).Get("mail")
    except (pywintypes.com_error, AttributeError):
        return None


def _check_disk(path: Path, need: int) -> None:
    free = shutil.disk_usage(path).free
    if free < need:
        raise StepError("DISK_SPACE", f"必要 {need / GB:.1f}GB / 空き {free / GB:.1f}GB", transient=True)


def _add_duration(entry: dict, phase: str, started: float) -> None:
    d = entry.setdefault("durations", {})
    d[f"{phase}_sec"] = round(d.get(f"{phase}_sec", 0) + time.monotonic() - started, 1)


def _reset_to_snapshot(entry: dict, work: Path, code: str) -> None:
    """作業ファイルを捨て、次回ログオフでスナップショットを取り直す。"""
    shutil.rmtree(work, ignore_errors=True)
    entry.update(checkpoint=None, snapshot_at=None, outputs={}, status="PENDING")
    warn(entry, code)
    log.warning("PST %s: %s のためスナップショットから再取得します", entry["key"], code)


def _restart(entry: dict, work: Path) -> None:
    """ログオフ時のスナップショットから全部やり直す（新しい世代として別フォルダーに出す）。"""
    shutil.rmtree(work, ignore_errors=True)
    entry.update(
        status="PENDING",
        checkpoint=None,
        snapshot_at=None,
        outputs={},
        counts={},
        attempts=0,
        last_error=None,
        warnings=[],
        generation=entry.get("generation", 1) + 1,
    )


def wait_released(path: Path, seconds: float) -> bool:
    """Outlook がファイルを手放すまで待つ（設計書 §5.5）。"""
    deadline = time.monotonic() + seconds
    while True:
        try:
            transfer.open_exclusive(path).Close()
            return True
        except pywintypes.error as e:
            if e.winerror in (2, 3):  # ファイルなし・フォルダーなし
                return True
            if time.monotonic() > deadline:
                return False
            time.sleep(2)


# ---------------------------------------------------------------------------
# ログオフ時：スナップショット（設計書 §4.2）
# ---------------------------------------------------------------------------


def copy_snapshot(src: Path, dst: Path) -> None:
    """書き込み拒否で開いてコピーする。開けない＝誰かが使用中。"""
    before = src.stat()
    try:
        handle = win32file.CreateFile(
            str(src),
            win32con.GENERIC_READ,
            win32con.FILE_SHARE_READ,
            None,
            win32con.OPEN_EXISTING,
            win32con.FILE_FLAG_SEQUENTIAL_SCAN,
            None,
        )
    except pywintypes.error as e:
        raise StepError("SOURCE_LOCKED", f"{src}: {e.strerror}", transient=True) from e
    partial = dst.with_name(dst.name + ".partial")
    fd = msvcrt.open_osfhandle(handle.Detach(), os.O_RDONLY)
    with os.fdopen(fd, "rb") as fi, open(partial, "wb") as fo:
        shutil.copyfileobj(fi, fo, transfer.CHUNK)
    after = src.stat()
    if (before.st_size, before.st_mtime_ns) != (after.st_size, after.st_mtime_ns):
        partial.unlink()
        raise StepError("SOURCE_CHANGED_DURING_COPY", str(src), transient=True)
    os.replace(partial, dst)


def run_snapshot(cfg: Config, home: Path) -> None:
    st = State(home / "state.json")
    budget_total = cfg.logoff_snapshot_budget_gb * GB
    budget = budget_total
    for e in st.psts.values():
        if e["status"] in TERMINAL or e.get("snapshot_at") or e.get("checkpoint"):
            continue
        src = Path(e["source_path"])
        try:
            size = src.stat().st_size
            if size > budget_total:
                warn(e, "SNAPSHOT_OVER_BUDGET")  # 管理者が予算を上げるまで取らない
                log.warning(
                    "PST %s: %.1fGB が上限 %gGB を超えるためコピーしません",
                    e["key"],
                    size / GB,
                    cfg.logoff_snapshot_budget_gb,
                )
                continue
            if size > budget:
                log.info("PST %s: 今回の上限を使い切ったため次回ログオフへ", e["key"])
                continue
            work = cfg.work_path / e["key"]
            shutil.rmtree(work, ignore_errors=True)  # 新しい世代の始まり。前の世代の scan.json 等を使わせない
            work.mkdir(parents=True, exist_ok=True)
            _check_disk(work, rules.disk_needed(size, cfg.disk_factor, cfg.disk_reserve_gb))
            t0 = time.monotonic()
            copy_snapshot(src, work / SNAPSHOT)
            budget -= size
            stamp = datetime.now().strftime("%Y%m%d%H%M%S")
            e.update(
                status="PENDING",
                snapshot_at=now_iso(),
                source_size=size,
                remote_dir=str(
                    PureWindowsPath("users", _owner(st.data), _computer(), f"{src.stem}_{e['key']}_{stamp}")
                ),
            )
            _add_duration(e, "snapshot", t0)
            log.info("スナップショット取得: %s (%.1fGB)", src, size / GB)
        except FileNotFoundError:
            apply_error(e, StepError("SOURCE_MISSING", str(src), transient=True), cfg.max_attempts)
        except StepError as err:
            apply_error(e, err, cfg.max_attempts)
        except OSError as ex:
            apply_error(e, StepError("SNAPSHOT_IO", str(ex), transient=True), cfg.max_attempts)
        finally:
            st.save()  # continue でも警告を残す


# ---------------------------------------------------------------------------
# ログオン時：検出〜転送
# ---------------------------------------------------------------------------


def _sevenzip() -> str:
    """EXE に同梱した 7z.exe（実行時に展開される _MEIPASS 内）。開発時は PATH 上の 7z。"""
    if getattr(sys, "frozen", False):
        bundled = Path(sys._MEIPASS) / "7z.exe"
        if bundled.exists():
            return str(bundled)
    found = shutil.which("7z")
    if not found:
        raise StepError("SEVENZIP_MISSING", "7z.exe が見つかりません", transient=True)
    return found


def _run_7z(args: list[str], cwd: Path, code: str) -> None:
    r = subprocess.run(
        [_sevenzip(), *args, "-bso0", "-bsp0"],
        cwd=cwd,
        capture_output=True,
        text=True,
        encoding="oem",
        errors="replace",
        creationflags=subprocess.CREATE_NO_WINDOW | subprocess.BELOW_NORMAL_PRIORITY_CLASS,
    )
    if r.returncode == 1:
        # 7z の「警告」終了＝ファイルを開けなかった（使用中）。PSTの問題ではないので回数に数えない
        raise StepError("PST_LOCKED", f"7z exit=1 {r.stderr.strip()[-500:]}", transient=True)
    if r.returncode != 0:
        raise StepError(code, f"7z exit={r.returncode} {r.stderr.strip()[-500:]}")


class Runner:
    def __init__(self, cfg: Config, home: Path, scan_only: bool = False):
        self.cfg = cfg
        self.home = home
        self.scan_only = scan_only
        self.names = rules.output_names(cfg.cutoff_local, _computer())
        self.st = State(home / "state.json")
        self.ol = None  # 裏起動の Outlook を終了させた後は None（必要になったら _outlook() で起動し直す）
        self.ol_lost = False  # 処理中に Outlook が落ちた。残りは次回ログオンで

    def _outlook(self):
        from pstmigration.outlook import Outlook  # COM はログオン時のみ

        if self.ol is None:
            self.ol = Outlook()
        return self.ol

    def _let_go(self) -> None:
        """切り離したPSTを Outlook に手放させる（裏起動の Outlook なら終了させる）。"""
        if self.ol is not None and self.ol.quit_if_headless():
            log.info("裏で起動した Outlook を終了しました（PSTを手放させるため）")
            self.ol = None
            gc.collect()  # 残っている COM 参照を解放し、Outlook のプロセスを終わらせる

    def _ol_died(self) -> bool:
        if self.ol is not None and not self.ol.alive():
            self.ol, self.ol_lost = None, True
        return self.ol_lost

    def run(self) -> None:
        transfer.connect(self.cfg.server_root, self.cfg.server_user, self.cfg.server_password)
        ready = False
        try:
            self._outlook().cleanup(self.cfg.work_path)
            self._resolve_mailbox()
            self._sync()
            self._apply_cutoff()
            self.st.data["last_run_error"] = None
            ready = True
            if not self.scan_only:
                self._backup_contacts()  # 小さいので、時間のかかる分割より先に
        except (StepError, pywintypes.com_error) as e:
            log.error("Outlook を利用できません: %s", e)
            self.st.data["last_run_error"] = {
                "code": getattr(e, "code", "OUTLOOK_COM"),
                "message": str(e)[:500],
                "at": now_iso(),
            }
        self.st.save()

        if ready:
            for e in list(self.st.psts.values()):
                if self.ol_lost:
                    break
                if e["status"] in TERMINAL:
                    continue
                if not e.get("checkpoint") and not e.get("snapshot_at"):
                    continue  # ログオフ時のスナップショット待ち
                self._process(e)
                self.st.save()

        report.publish_status(self.cfg.server_path, self.st.data, _user(), _computer())

    # ---- 検出 ------------------------------------------------------------

    def _resolve_mailbox(self) -> None:
        """NAS のフォルダー名・Purview の Mailbox に使うメールアドレスを決める（設計書 §8.2）。"""
        mailbox, source = rules.pick_mailbox(self.ol.smtp_addresses(), _ad_mail, _user())
        if source == "USERNAME" and "@" in (self.st.data.get("mailbox") or ""):
            return  # 今回たまたま取れなかっただけなら、前回のアドレスを使い続ける
        if mailbox != self.st.data.get("mailbox"):
            log.info("メールアドレス: %s（%s）", mailbox, source)
        self.st.data["mailbox"] = mailbox
        self.st.data["mailbox_source"] = source

    def _sync(self) -> None:
        for d in self.ol.detect(self.cfg.work_path):
            e = self.st.psts.get(d["key"])
            if e is None:
                e = new_entry(d["key"], d["path"], d["is_default"])
                if d["skip"]:
                    e["status"] = "SKIPPED"
                    e["skip_reason"] = d["skip"]
                self.st.psts[d["key"]] = e
                log.info("PST検出: %s (%s)", d["path"], d["skip"] or "対象")
            with contextlib.suppress(OSError):  # 毎回更新（上限超えで止まっているPSTを台帳で見つけるため）
                e["source_size"] = os.stat(d["path"]).st_size
            e["is_default_store"] = d["is_default"]
            if e["status"] == "COMPLETED" and self.ol.live_counts(d["store"]) != e.get("folder_counts"):
                if self.cfg.recollect_on_change:
                    log.info("PST %s: 完了後に変更あり。再収集します", e["key"])
                    _restart(e, self.cfg.work_path / e["key"])
                else:
                    warn(e, "SOURCE_CHANGED")

    def _apply_cutoff(self) -> None:
        """config の cutoff が変わったら、旧 cutoff で分けたPSTを作り直す。

        旧世代の NAS フォルダーは、新世代の転送が済んだ時点で消す（_finish）。それまでは残すので、
        NAS 上に何もない期間はできない。
        """
        for e in self.st.psts.values():
            if e["status"] == "SKIPPED":
                continue
            started = e.get("snapshot_at") or e.get("checkpoint") or e["status"] in TERMINAL
            same = datetime.fromisoformat(e.get("cutoff") or LEGACY_CUTOFF) == self.cfg.cutoff_local
            if started and not same:
                log.info("PST %s: cutoff が %s に変わったため作り直します", e["key"], self.cfg.cutoff)
                if e.get("remote_dir"):
                    e.setdefault("stale_remote_dirs", []).append(e["remote_dir"])
                _restart(e, self.cfg.work_path / e["key"])
                e["remote_dir"] = None
            e["cutoff"] = self.cfg.cutoff

    def _purge_stale(self, e: dict) -> None:
        for d in list(e.get("stale_remote_dirs") or []):
            try:
                shutil.rmtree(self.cfg.server_path / d)
            except FileNotFoundError:
                pass
            except OSError as ex:
                log.warning("PST %s: 旧 cutoff のフォルダーを消せません（手で削除）: %s", e["key"], ex)
                warn(e, "STALE_REMOTE_LEFT")
                continue
            log.info("PST %s: 旧 cutoff のフォルダーを削除: %s", e["key"], d)
            e["stale_remote_dirs"].remove(d)

    # ---- 連絡先バックアップ（日付で分けず、毎回作り直して NAS 上を置き換える） ----

    def _backup_contacts(self) -> None:
        c = self.st.data.setdefault("contacts", {})
        work = self.cfg.work_path / CONTACTS_DIR
        pst = work / f"contacts_{_computer()}.pst"
        try:
            # 前回の残りは cleanup() で切り離し済み。Outlook が手放すのを待って作り直す
            if not wait_released(pst, self.cfg.release_wait_sec):
                raise StepError("PST_LOCKED", f"Outlook がファイルを解放しません: {pst.name}", transient=True)
            shutil.rmtree(work, ignore_errors=True)
            work.mkdir(parents=True, exist_ok=True)
            # Outlook への参照をローカル変数に持たない（残っていると終了させても PST を手放さない）
            store = self._outlook().mount(pst, create=True, label="連絡先バックアップ")
            try:
                count = self.ol.export_contacts(store, self.cfg.work_path)
            finally:
                self.ol.unmount(store)
                del store
            self._let_go()
            if count:
                if not wait_released(pst, self.cfg.release_wait_sec):
                    raise StepError("PST_LOCKED", f"Outlook がファイルを解放しません: {pst.name}", transient=True)
                remote = PureWindowsPath("users", _owner(self.st.data), _computer(), pst.name)
                transfer.upload(
                    pst,
                    self.cfg.server_path / remote,
                    transfer.sha256_file(pst),
                    self.cfg.transfer_max_mbps,
                    overwrite=True,
                )
                c.update(remote_path=str(remote), size=pst.stat().st_size)
            c.update(backup_at=now_iso(), count=count, last_error=None)
            log.info("連絡先バックアップ: %d 件", count)
            shutil.rmtree(work, ignore_errors=True)
        except Exception as ex:  # 連絡先が失敗してもメール移行は続ける
            self._ol_died()
            code = getattr(ex, "code", "OUTLOOK_COM" if isinstance(ex, pywintypes.com_error) else "UNEXPECTED")
            log.warning("連絡先バックアップに失敗: %s", ex, exc_info=code == "UNEXPECTED")
            c["last_error"] = {"code": code, "message": str(ex)[:500], "at": now_iso()}
        self.st.save()

    # ---- PST 1件 ---------------------------------------------------------

    def _process(self, e: dict) -> None:
        work = self.cfg.work_path / e["key"]
        try:
            if not self._artifacts_ok(e, work):
                _reset_to_snapshot(e, work, "ARTIFACT_MISSING")
                return
            if self.scan_only:
                if e.get("checkpoint") is None:
                    self._split(e, work)
                return
            if e.get("checkpoint") is None and not self._split(e, work):
                return
            self._let_go()  # 以降（7z・転送）は Outlook を使わない。分割したPSTを手放させる
            if e["checkpoint"] == "SPLIT":
                self._archive(e, work)
                self.st.save()
            if e["checkpoint"] == "ARCHIVED":
                self._transfer(e, work)
                self.st.save()
            if e["checkpoint"] == "TRANSFERRED":
                self._finish(e, work)
        except Exception as ex:  # 想定外でも他のPSTの処理は続ける
            apply_error(e, self._classify(e, ex), self.cfg.max_attempts)

    def _classify(self, e: dict, ex: Exception) -> StepError:
        if self._ol_died():
            # 処理中に Outlook が終了・クラッシュした。PST側の問題ではないので回数に数えない
            return StepError("OUTLOOK_LOST", repr(ex), transient=True)
        if isinstance(ex, StepError):
            return ex
        if isinstance(ex, PermissionError):
            # 切り離し直後の Outlook やウイルス対策ソフトがファイルを掴んでいる（共有違反）
            return StepError("PST_LOCKED", repr(ex), transient=True)
        if isinstance(ex, pywintypes.com_error):
            return StepError("OUTLOOK_COM", str(ex), transient=True)
        log.exception("PST %s: 想定外のエラー", e["key"])
        return StepError("UNEXPECTED", repr(ex))

    def _artifacts_ok(self, e: dict, work: Path) -> bool:
        """出力ファイルが欠けていたら完了扱いにしない（FR-08）。"""
        cp = e.get("checkpoint")
        if cp == "SPLIT":
            return (work / self.names.old_pst).exists() and (work / self.names.new_pst).exists()
        if cp == "ARCHIVED":
            return all(
                (work / o["name"]).exists() and (work / o["name"]).stat().st_size == o["size"]
                for o in e["outputs"].values()
            )
        return True

    def _split(self, e: dict, work: Path) -> bool:
        """スナップショットの調査・分割・検証（設計書 §5）。"""
        snap = work / SNAPSHOT
        if not snap.exists():
            _reset_to_snapshot(e, work, "SNAPSHOT_MISSING")
            return False
        # 前回の異常終了で残った接続を cleanup() で外した直後は、まだ Outlook が掴んでいる
        if not wait_released(snap, self.cfg.release_wait_sec):
            raise StepError("PST_LOCKED", f"Outlook がファイルを解放しません: {snap.name}", transient=True)
        kind, fmt = rules.pst_header(snap)
        if kind != "PST":
            raise StepError("NOT_PST", f"ヘッダー判定={kind}", final=True)
        e["pst_format"] = fmt

        old_path, new_path = work / self.names.old_pst, work / self.names.new_pst
        existing = sum(p.stat().st_size for p in (old_path, new_path) if p.exists())
        size = snap.stat().st_size
        _check_disk(work, rules.disk_needed(size, self.cfg.disk_factor - 1, self.cfg.disk_reserve_gb) - existing)

        cutoff = self.cfg.cutoff_utc
        label_date = self.cfg.cutoff_local.strftime("%Y-%m-%d")
        e["status"] = "SCANNING"
        self.st.save()
        t0 = time.monotonic()
        ol = self._outlook()
        snap_store = old_store = new_store = None
        bad: list[str] = []
        try:
            snap_store = ol.mount(snap, label=SNAPSHOT_LABEL)
            scan = read_json(work / SCAN, None)
            ids = None  # 途中再開なら split() がスナップショットの残りを調べ直す
            if scan is None:  # 分割開始前の値を正解として固定する
                scan, ids = ol.scan(snap_store, cutoff, log_subject=self.cfg.log_subject)
                write_json(work / SCAN, scan)
            e["counts"] = rules.totals(scan["folders"])
            e["folder_counts"] = rules.folder_counts(scan["folders"])
            e["problems"] = scan["problems"]
            _add_duration(e, "scan", t0)
            if self.scan_only:
                write_json(
                    self.home / "reports" / f"{e['key']}.scan.json",
                    report.build(e, "SCANNED", self.st.data, _user(), _computer(), scan["folders"]),
                )
                return True

            e["status"] = "SPLITTING"
            self.st.save()
            t0 = time.monotonic()
            old_store = ol.mount(old_path, create=not old_path.exists(), label=f"旧メール（{label_date}より前）")
            new_store = ol.mount(new_path, create=not new_path.exists(), label=f"新メール（{label_date}以降）")
            ol.split(snap_store, old_store, new_store, cutoff, self.cfg.split_pause_ms, ids)
            del ids
            e["status"] = "VERIFYING"
            self.st.save()
            bad = ol.verify(snap_store, old_store, new_store, scan["folders"])
            _add_duration(e, "split", t0)
        finally:
            for s in (new_store, old_store, snap_store):
                if s is not None:
                    ol.unmount(s)

        del ol, s, snap_store, old_store, new_store  # 参照が残っていると Outlook が終了できない
        if bad:
            self._let_go()
            for p in (snap, old_path, new_path):
                wait_released(p, self.cfg.release_wait_sec)
            shutil.rmtree(work, ignore_errors=True)
            e.update(snapshot_at=None, outputs={})
            raise StepError("COUNT_MISMATCH", " / ".join(bad[:5]))
        e["checkpoint"] = "SPLIT"
        self.st.save()
        return True

    def _archive(self, e: dict, work: Path) -> None:
        """設計書 §6。"""
        e["status"] = "COMPRESSING"
        self.st.save()
        t0 = time.monotonic()
        old_pst, new_pst, archive = (work / n for n in (self.names.old_pst, self.names.new_pst, self.names.old_7z))
        for p in (old_pst, new_pst, work / SNAPSHOT):
            if not wait_released(p, self.cfg.release_wait_sec):
                raise StepError("PST_LOCKED", f"Outlook がファイルを解放しません: {p.name}", transient=True)
        (work / SNAPSHOT).unlink(missing_ok=True)
        archive.unlink(missing_ok=True)

        # 0件の側はファイルを作らない（空のPST・7zを置いても後続作業の邪魔になるだけ）
        c = self.cfg
        counts = e.get("counts") or {}
        e["outputs"] = {}
        if counts.get("new", 0):
            e["outputs"]["new_pst"] = {
                "name": new_pst.name,
                "size": new_pst.stat().st_size,
                "sha256": transfer.sha256_file(new_pst),
            }
            if e["outputs"]["new_pst"]["size"] > c.new_pst_warn_gb * GB:
                warn(e, "NEW_PST_OVER_LIMIT")
        else:
            new_pst.unlink(missing_ok=True)
        if not counts.get("old", 0):
            old_pst.unlink(missing_ok=True)
            e["checkpoint"] = "ARCHIVED"
            _add_duration(e, "compress", t0)
            self.st.save()
            return

        # 上の SHA-256 計算（数GBで数分）の間に、ウイルス対策ソフト等が旧PSTを開くことがあるので直前にもう一度待つ
        if not wait_released(old_pst, c.release_wait_sec):
            raise StepError("PST_LOCKED", f"ファイルが使用中のままです: {old_pst.name}", transient=True)
        _run_7z(
            [
                "a",
                "-t7z",
                "-m0=lzma2",
                f"-mx={c.sevenzip_level}",
                f"-mmt={c.sevenzip_threads}",
                "-y",
                archive.name,
                old_pst.name,
            ],
            work,
            "ARCHIVE_FAILED",
        )
        e["status"] = "VERIFYING"
        self.st.save()
        try:
            _run_7z(["t", archive.name], work, "ARCHIVE_TEST_FAILED")
        except StepError:
            archive.unlink(missing_ok=True)
            raise

        pst_size, zip_size = old_pst.stat().st_size, archive.stat().st_size
        e["archive"] = {
            "pst_size": pst_size,
            "size": zip_size,
            "ratio": round(zip_size / pst_size, 4) if pst_size else None,
        }
        e["outputs"]["old_7z"] = {"name": archive.name, "size": zip_size, "sha256": transfer.sha256_file(archive)}
        e["checkpoint"] = "ARCHIVED"
        _add_duration(e, "compress", t0)
        self.st.save()
        old_pst.unlink()  # 7z テスト合格後のみ削除（設計書 §7.3）

    def _transfer(self, e: dict, work: Path) -> None:
        """設計書 §8.3。"""
        e["status"] = "TRANSFERRING"
        self.st.save()
        t0 = time.monotonic()
        root = self.cfg.server_path
        remote = root / e["remote_dir"]
        with transfer.slot(root, self.cfg.slot_wait_min):
            for o in e["outputs"].values():
                transfer.upload(work / o["name"], remote / o["name"], o["sha256"], self.cfg.transfer_max_mbps)
        e["checkpoint"] = "TRANSFERRED"
        _add_duration(e, "transfer", t0)

    def _finish(self, e: dict, work: Path) -> None:
        unmigrated = e["counts"].get("unmigrated", 0)
        status = "COMPLETED" if unmigrated == 0 else "FAILED"
        if unmigrated:
            e["last_error"] = {
                "code": "UNMIGRATED_ITEMS",
                "message": f"未移行 {unmigrated} 件（元PSTに残存）",
                "at": now_iso(),
                "retryable": False,
            }
        folders = (read_json(work / SCAN, None) or {}).get("folders", [])
        rep = report.build(e, status, self.st.data, _user(), _computer(), folders)
        write_json(self.home / "reports" / f"{e['key']}.json", rep)
        try:
            write_json(self.cfg.server_path / e["remote_dir"] / "report.json", rep)
        except OSError as ex:
            raise StepError("NETWORK", str(ex), transient=True) from ex
        e["status"] = status
        self._purge_stale(e)  # 新世代が NAS に揃ったので、旧 cutoff の分を消す
        if status == "COMPLETED":
            e["last_error"] = None  # 途中で再試行した名残を台帳に残さない
        e["finished_at"] = rep["finished_at"]
        shutil.rmtree(work, ignore_errors=True)
        log.info("PST %s: %s", e["key"], status)


def reset(home: Path, key: str) -> bool:
    """管理者用：FAILED 等を PENDING に戻す（checkpoint は保持して続きから）。"""
    st = State(home / "state.json")
    e = st.psts.get(key)
    if e is None:
        return False
    e.update(status="PENDING", attempts=0, last_error=None)
    st.save()
    return True
