"""Outlook COM 操作：検出・マウント・調査・分割・検証（設計書 §3、§5）。

元PSTに対しては読み取り（Stores列挙・フォルダー件数）しか行わない。
Move などの書き込みはスナップショットと出力PSTに対してのみ行う。
"""

from __future__ import annotations

import contextlib
import logging
import time
from collections.abc import Iterator
from datetime import datetime
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client

from pstmigration import rules
from pstmigration.state import StepError

log = logging.getLogger(__name__)

OL_MAIL_ITEM = 0  # Folder.DefaultItemType
OL_CONTACT_ITEM = 2
OL_FOLDER_INBOX = 6  # Folders.Add の種別（メールフォルダー）
OL_FOLDER_CONTACTS = 10
OL_STORE_UNICODE = 2
OL_NOT_EXCHANGE = 3
OL_USER_ITEMS = 0

PR_FOLDER_TYPE = "http://schemas.microsoft.com/mapi/proptag/0x36010003"
FOLDER_SEARCH = 2

_PT = "http://schemas.microsoft.com/mapi/proptag/"
COLUMNS = [
    "EntryID",
    "MessageClass",
    "Subject",
    _PT + "0x0E060040",  # PR_MESSAGE_DELIVERY_TIME (= ReceivedTime)
    _PT + "0x00390040",  # PR_CLIENT_SUBMIT_TIME    (= SentOn)
    _PT + "0x007D001F",  # PR_TRANSPORT_MESSAGE_HEADERS（あれば受信メール）
    _PT + "0x0E080003",  # PR_MESSAGE_SIZE
]

# Outlook がモーダルダイアログ表示中などで呼び出しを拒否した
_RETRY_HRESULTS = {-2147418111, -2147417846}  # RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER

MAX_PROBLEMS = 1000

Names = tuple[str, ...]


def _call(fn, *args):
    for i in range(5):
        try:
            return fn(*args)
        except pywintypes.com_error as e:
            if e.args and e.args[0] in _RETRY_HRESULTS and i < 4:
                time.sleep(1 + i)
                continue
            raise


def _kind(folder) -> str:
    try:
        if int(folder.PropertyAccessor.GetProperty(PR_FOLDER_TYPE)) == FOLDER_SEARCH:
            return "search"
    except pywintypes.com_error:
        pass
    try:
        return "mail" if folder.DefaultItemType == OL_MAIL_ITEM else "other"
    except pywintypes.com_error:
        return "other"


def walk(folder, names: Names = ()) -> Iterator[tuple[Names, object, str]]:
    """(フォルダー名の並び, Folder, 'mail'|'other') を深さ優先で。検索フォルダーは配下ごと除外。"""
    subs = folder.Folders
    for i in range(1, subs.Count + 1):
        f = subs.Item(i)
        n = names + (f.Name,)
        kind = _kind(f)
        if kind == "search":
            continue
        yield n, f, kind
        yield from walk(f, n)


def _contact_tops(folder) -> Iterator[object]:
    """連絡先フォルダーのうち最上位のもの（CopyTo で配下ごとコピーされるため子は返さない）。"""
    subs = folder.Folders
    for i in range(1, subs.Count + 1):
        f = subs.Item(i)
        if _kind(f) == "search":
            continue
        try:
            is_contact = f.DefaultItemType == OL_CONTACT_ITEM
        except pywintypes.com_error:
            is_contact = False
        if is_contact:
            yield f
        else:
            yield from _contact_tops(f)


def _tree_count(folder) -> int:
    subs = folder.Folders
    return folder.Items.Count + sum(_tree_count(subs.Item(i)) for i in range(1, subs.Count + 1))


def rel_path(names: Names) -> str:
    return "\\" + "\\".join(names)


def _child(folder, names: Names, create: bool):
    for name in names:
        subs = folder.Folders
        try:
            folder = subs.Item(name)
        except pywintypes.com_error:
            if not create:
                return None
            folder = subs.Add(name, OL_FOLDER_INBOX)
    return folder


def _count(root, names: Names) -> int:
    f = _child(root, names, create=False)
    return -1 if f is None else f.Items.Count


def _rows(folder) -> Iterator[tuple]:
    table = folder.GetTable("", OL_USER_ITEMS)
    cols = table.Columns
    cols.RemoveAll()
    for c in COLUMNS:
        cols.Add(c)
    while not table.EndOfTable:
        yield table.GetNextRow().GetValues()


class Outlook:
    def __init__(self) -> None:
        pythoncom.CoInitialize()
        try:
            # 起動中ならそのインスタンス、未起動ならUIなしで起動される
            self.app = win32com.client.dynamic.Dispatch("Outlook.Application")
            self.ns = self.app.GetNamespace("MAPI")
            self.ns.Logon("", "", False, False)  # プロファイル選択ダイアログを出さない
        except pywintypes.com_error as e:
            raise StepError("OUTLOOK_UNAVAILABLE", str(e), transient=True) from e

    def quit_if_headless(self) -> bool:
        """本ツールが裏で起動した（画面のない）Outlook なら終了させ、True を返す。

        画面のない Outlook は、切り離したPSTも終了するまで掴み続ける（パイロットで確認）。
        ユーザーが開いている Outlook は閉じない。
        """
        try:
            if self.app.Explorers.Count or self.app.Inspectors.Count:
                return False
            self.app.Quit()
        except (pywintypes.com_error, AttributeError):
            pass  # もう終わっている
        return True

    def alive(self) -> bool:
        """Outlook のプロセスがまだいるか。落ちると dynamic Dispatch は AttributeError を出す。"""
        try:
            return self.ns.Stores.Count >= 0
        except (pywintypes.com_error, AttributeError):
            return False

    # ---- ストア ----------------------------------------------------------

    def _stores(self) -> Iterator[object]:
        stores = self.ns.Stores
        for i in range(1, stores.Count + 1):
            yield stores.Item(i)

    @staticmethod
    def _path(store) -> str:
        try:
            return store.FilePath or ""
        except pywintypes.com_error:
            return ""

    def smtp_addresses(self) -> list[str]:
        """プロファイル内のメールアカウントのアドレス（IMAP移行後の本人のアドレス）。"""
        out = []
        try:
            accounts = self.ns.Accounts
            for i in range(1, accounts.Count + 1):
                with contextlib.suppress(pywintypes.com_error):
                    out.append(accounts.Item(i).SmtpAddress or "")
        except pywintypes.com_error:
            pass
        return out

    def find_store(self, path: Path):
        target = rules.norm_path(str(path))
        for s in self._stores():
            p = self._path(s)
            if p and rules.norm_path(p) == target:
                return s
        return None

    def detect(self, work_dir: Path) -> list[dict]:
        """設計書 §3.2。作業フォルダー配下（本ツールのマウント）は無視する。"""
        try:
            default = rules.norm_path(self.ns.DefaultStore.FilePath or "")
        except pywintypes.com_error:
            default = ""
        work = rules.norm_path(str(work_dir)) + "\\"
        found: dict[str, dict] = {}
        for store in self._stores():
            path = self._path(store)
            if not path:
                continue
            n = rules.norm_path(path)
            if n in found or n.startswith(work):
                continue
            skip = None
            if n.endswith(".ost"):
                skip = "OST"
            elif not n.endswith(".pst"):
                skip = "NOT_PST"
            elif store.ExchangeStoreType != OL_NOT_EXCHANGE:
                skip = "EXCHANGE"
            elif rules.TOOL_OUTPUT_RE.search(n):
                skip = "TOOL_OUTPUT"
            found[n] = {
                "path": path,
                "key": rules.pst_key(path),
                "is_default": n == default,
                "skip": skip,
                "store": store,
            }
        return list(found.values())

    def mount(self, path: Path, *, create: bool = False, label: str | None = None):
        try:
            if create:
                self.ns.AddStoreEx(str(path), OL_STORE_UNICODE)
            else:
                self.ns.AddStore(str(path))
        except pywintypes.com_error as e:
            raise StepError("PST_OPEN_FAILED", f"{path}: {e}") from e
        for _ in range(40):
            store = self.find_store(path)
            if store is not None:
                if label:
                    # 表示名の変更は見た目だけなので失敗してもよい
                    with contextlib.suppress(pywintypes.com_error):
                        store.GetRootFolder().Name = label
                return store
            time.sleep(0.25)
        raise StepError("PST_OPEN_FAILED", f"Outlookがストアを認識しません: {path}")

    def unmount(self, store) -> None:
        try:
            self.ns.RemoveStore(store.GetRootFolder())
        except pywintypes.com_error as e:
            log.warning("ストアの切り離しに失敗: %s", e)

    def cleanup(self, work_dir: Path) -> None:
        """前回異常終了で残った作業用PSTの接続を外す（設計書 §3.3）。"""
        work = rules.norm_path(str(work_dir)) + "\\"
        for s in list(self._stores()):
            p = self._path(s)
            if p and rules.norm_path(p).startswith(work):
                log.info("残っていた作業用PSTを切り離します: %s", p)
                self.unmount(s)

    def export_contacts(self, dst, work_dir: Path) -> int:
        """全ストアの連絡先フォルダーを dst へコピーし、件数を返す（元ストアは読むだけ）。

        Exchange のストアはサーバ側にあるので対象外。IMAP の OST（このコンピューターのみ）は含める。
        """
        work = rules.norm_path(str(work_dir)) + "\\"
        root = dst.GetRootFolder()
        total = 0
        for i, store in enumerate(list(self._stores()), 1):
            path = self._path(store)
            if not path or rules.norm_path(path).startswith(work) or store.ExchangeStoreType != OL_NOT_EXCHANGE:
                continue
            tops = list(_contact_tops(store.GetRootFolder()))
            if not tops:
                continue
            # 同名の「連絡先」がストアごとにあるため、ストア単位のフォルダーに分ける
            box = root.Folders.Add(f"{i}. {store.DisplayName}", OL_FOLDER_CONTACTS)
            for f in tops:
                n = _tree_count(f)
                if _tree_count(_call(f.CopyTo, box)) != n:
                    raise StepError("CONTACTS_MISMATCH", f"{store.DisplayName}\\{f.Name}: 元 {n} 件")
                total += n
        return total

    # ---- 調査・分割・検証 --------------------------------------------------

    def scan(
        self, store, cutoff_utc: datetime, *, log_subject: bool = False
    ) -> tuple[dict, dict[Names, dict[str, list[str]]]]:
        """設計書 §5.3。戻り値: (scan.json の内容, フォルダー別の移動対象EntryID)。"""
        folders: list[dict] = []
        problems: list[dict] = []
        ids: dict[Names, dict[str, list[str]]] = {}
        for names, f, kind in walk(store.GetRootFolder()):
            rec = {
                "path": rel_path(names),
                "kind": kind,
                "old": 0,
                "new": 0,
                "unmigrated": 0,
                "non_mail": 0,
                "near_cutoff": 0,
                "old_bytes": 0,
                "new_bytes": 0,
            }
            if kind != "mail":
                rec["non_mail"] = f.Items.Count
                folders.append(rec)
                continue
            bucket_ids: dict[str, list[str]] = {"old": [], "new": []}
            for eid, mclass, subject, delivery, submit, headers, size in _rows(f):
                bucket, d = rules.classify(
                    mclass, rules.to_utc(delivery), rules.to_utc(submit), bool(headers), cutoff_utc
                )
                rec[bucket] += 1
                if bucket in bucket_ids:
                    bucket_ids[bucket].append(eid)
                    rec[bucket + "_bytes"] += int(size or 0)
                    if rules.is_near_cutoff(d, cutoff_utc):
                        rec["near_cutoff"] += 1
                elif bucket == "unmigrated" and len(problems) < MAX_PROBLEMS:
                    p = {"code": "NO_DATE", "folder": rec["path"], "message_class": mclass}
                    if log_subject:
                        p["subject"] = subject
                    problems.append(p)
            ids[names] = bucket_ids
            folders.append(rec)
        return {"folders": folders, "problems": problems}, ids

    def split(self, snap, old, new, cutoff_utc: datetime, pause_ms: int) -> int:
        """スナップショットから出力PSTへ Move する（設計書 §5.4）。

        途中再開に対応：スナップショットに残っているものだけを再調査して移動する。
        """
        _, ids = self.scan(snap, cutoff_utc)
        store_id = snap.StoreID
        roots = {"old": old.GetRootFolder(), "new": new.GetRootFolder()}
        moved = 0
        for names, buckets in ids.items():
            # 空フォルダーも含め、全メールフォルダーを両方のPSTに作る（階層の維持）
            dests = {b: _child(roots[b], names, create=True) for b in roots}
            for bucket, eids in buckets.items():
                for eid in eids:
                    item = _call(self.ns.GetItemFromID, eid, store_id)
                    _call(item.Move, dests[bucket])
                    del item
                    moved += 1
                    if moved % 100 == 0:
                        time.sleep(pause_ms / 1000)
        log.info("移動 %d 件", moved)
        return moved

    def verify(self, snap, old, new, folders: list[dict]) -> list[str]:
        """設計書 §5.5。不一致の説明のリスト（空なら合格）。"""
        r_snap, r_old, r_new = snap.GetRootFolder(), old.GetRootFolder(), new.GetRootFolder()
        bad = []
        for rec in folders:
            if rec["kind"] != "mail":
                continue
            names = tuple(rec["path"].split("\\")[1:])
            got = (_count(r_old, names), _count(r_new, names), _count(r_snap, names))
            exp = (rec["old"], rec["new"], rec["unmigrated"] + rec["non_mail"])
            if got != exp:
                bad.append(f"{rec['path']} 期待(旧,新,残)={exp} 実際={got}")
        return bad

    def live_counts(self, store) -> dict[str, int]:
        """元PSTのメールフォルダー別件数（読み取りのみ。設計書 §9.4）。"""
        return {rel_path(n): f.Items.Count for n, f, kind in walk(store.GetRootFolder()) if kind == "mail"}
