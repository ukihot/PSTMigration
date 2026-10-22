"""ライブ分割方式の Extended MAPI 操作（設計書 §16.4〜§16.6）。

元PSTはユーザーのプロファイルへの別セッションから読み取り専用で開く。束PSTは一時プロファイル
（束PSTだけを入れる）で開き、ユーザーの Outlook には接続しない（§16.5.1）。
"""

from __future__ import annotations

import contextlib
import gc
import hashlib
import itertools
import os
import time
from datetime import datetime
from pathlib import Path

from pstmigration import rules

# 内容テーブルの列
PR_ENTRYID = 0x0FFF0102
PR_SEARCH_KEY = 0x300B0102
PR_MESSAGE_CLASS = 0x001A001F
PR_SUBJECT = 0x0037001F
PR_DELIVERY_TIME = 0x0E060040  # PR_MESSAGE_DELIVERY_TIME
PR_SUBMIT_TIME = 0x00390040  # PR_CLIENT_SUBMIT_TIME
PR_HEADERS = 0x007D001F  # PR_TRANSPORT_MESSAGE_HEADERS（あれば受信メール）
PR_SIZE = 0x0E080003  # PR_MESSAGE_SIZE
PR_HASATTACH = 0x0E1B000B
MSG_COLUMNS = (
    PR_ENTRYID,
    PR_SEARCH_KEY,
    PR_MESSAGE_CLASS,
    PR_SUBJECT,
    PR_DELIVERY_TIME,
    PR_SUBMIT_TIME,
    PR_HEADERS,
    PR_SIZE,
    PR_HASATTACH,
)

PT_ERROR = 0x000A
E_NOT_ENOUGH_MEMORY = -2147024882  # 0x8007000E：値はあるが大きすぎてテーブルに載らない
FOLDER_SEARCH = 2
NON_MAIL_FOLDER_CLASSES = ("IPF.APPOINTMENT", "IPF.CONTACT", "IPF.TASK", "IPF.STICKYNOTE", "IPF.JOURNAL")
MAX_PROBLEMS = 1000


def _is_error(tag: int) -> bool:
    return tag & 0xFFFF == PT_ERROR


def value(prop: tuple):
    """内容テーブルの (タグ, 値)。値が無い列（PT_ERROR）は None。"""
    tag, v = prop
    return None if _is_error(tag) else v


def present(prop: tuple) -> bool:
    """値があるか。大きすぎて載らない値は PT_ERROR + MAPI_E_NOT_ENOUGH_MEMORY で返るが、値はある。"""
    tag, v = prop
    return not _is_error(tag) or v in (E_NOT_ENOUGH_MEMORY, E_NOT_ENOUGH_MEMORY & 0xFFFFFFFF)


def folder_kind(container_class: str | None, folder_type: int | None) -> str:
    """'mail' / 'other' / 'search'。既知の非メールのクラスだけ other（それ以外はメールとして中身を判定する）。"""
    if folder_type == FOLDER_SEARCH:
        return "search"
    return "other" if (container_class or "").upper().startswith(NON_MAIL_FOLDER_CLASSES) else "mail"


def _iso(d: datetime | None) -> str:
    u = rules.to_utc(d)
    return u.isoformat() if u else ""


def fingerprint(subject, delivery, submit, message_class, has_attach) -> str:
    """照合用の指紋（§16.6）。内容テーブルの列だけで作る（メールを1通ずつ開かない）。"""
    return "|".join((subject or "", _iso(delivery), _iso(submit), message_class or "", "1" if has_attach else "0"))


def _row(path: str, props) -> tuple[dict, dict]:
    """内容テーブルの1行 → (列ごとの (タグ, 値), 計画・照合用の行 {e, f, k, p, s})。"""
    c = dict(zip(MSG_COLUMNS, props, strict=True))  # 行は SetColumns の順に並ぶ
    eid, key = value(c[PR_ENTRYID]), value(c[PR_SEARCH_KEY])
    fp = fingerprint(
        value(c[PR_SUBJECT]),
        value(c[PR_DELIVERY_TIME]),
        value(c[PR_SUBMIT_TIME]),
        value(c[PR_MESSAGE_CLASS]),
        value(c[PR_HASATTACH]),
    )
    row = {
        "e": eid.hex() if eid else "",
        "f": path,
        "k": key.hex() if key else "",
        "p": fp,
        "s": int(value(c[PR_SIZE]) or 0),
    }
    return c, row


def classify_rows(path: str, raw_rows: list, cutoff_utc: datetime, *, log_subject: bool = False):
    """1フォルダー分の内容テーブルを、調査結果（現行 scan.json と同じ形）と計画のメール一覧に分ける。

    戻り値: (フォルダーの集計, {"old": [...], "new": [...]}, 全行（ハッシュ対象の判定用）, 問題の一覧)
    """
    rec = {
        "path": path,
        "kind": "mail",
        "old": 0,
        "new": 0,
        "unmigrated": 0,
        "non_mail": 0,
        "near_cutoff": 0,
        "old_bytes": 0,
        "new_bytes": 0,
    }
    items: dict[str, list[dict]] = {"old": [], "new": []}
    all_rows, problems = [], []
    for props in raw_rows:
        c, row = _row(path, props)
        mclass, subject = value(c[PR_MESSAGE_CLASS]), value(c[PR_SUBJECT])
        delivery, submit, size = value(c[PR_DELIVERY_TIME]), value(c[PR_SUBMIT_TIME]), row["s"]
        all_rows.append(row)
        bucket, d = rules.classify(
            mclass, rules.to_utc(delivery), rules.to_utc(submit), present(c[PR_HEADERS]), cutoff_utc
        )
        rec[bucket] += 1
        if bucket in items:
            items[bucket].append(row)
            rec[bucket + "_bytes"] += size
            if rules.is_near_cutoff(d, cutoff_utc):
                rec["near_cutoff"] += 1
        elif bucket == "unmigrated" and len(problems) < MAX_PROBLEMS:
            p = {"code": "NO_DATE", "folder": path, "message_class": mclass}
            if log_subject:
                p["subject"] = subject
            problems.append(p)
    return rec, items, all_rows, problems


def digest(parts) -> str:
    """本文・添付の中身のハッシュ。部分ごとに長さを付けてつなぐ（区切りの違いで同じ値にならない）。"""
    h = hashlib.sha256()
    for p in parts:
        h.update(len(p).to_bytes(8, "little"))
        h.update(p)
    return h.hexdigest()


# ---- MAPI ------------------------------------------------------------------------

TEMP_PREFIX = "PSTMigration-tmp-"
_seq = itertools.count(1)

PR_DISPLAY_NAME_A = 0x3001001E
PR_DISPLAY_NAME_W = 0x3001001F
PR_CONTAINER_CLASS_W = 0x3613001F
PR_FOLDER_TYPE = 0x36010003
PR_IPM_SUBTREE_ENTRYID = 0x35E00102
PR_SERVICE_UID = 0x3D0C0102
PR_SERVICE_NAME_A = 0x3D09001E
PR_PST_PATH_W = 0x6700001F
PR_BODY_W = 0x1000001F
PR_HTML = 0x10130102
PR_RTF_COMPRESSED = 0x10090102
PR_ATTACH_NUM = 0x0E210003
PR_ATTACH_METHOD = 0x37050003
PR_ATTACH_LONG_FILENAME_W = 0x3707001F
PR_ATTACH_DATA_BIN = 0x37010102
PR_ATTACH_DATA_OBJ = 0x3701000D
ATTACH_EMBEDDED_MSG = 5
FOLDER_GENERIC = 1
OPEN_IF_EXISTS = 1
QUERY_BATCH = 500


def _mapi():
    from win32comext.mapi import mapi

    return mapi


class Mapi:
    """MAPIInitialize / MAPIUninitialize（プロセス内で入れ子にしてよい）。"""

    def __enter__(self):
        _mapi().MAPIInitialize(None)
        return self

    def __exit__(self, *exc):
        gc.collect()  # MAPI オブジェクトを全部解放してから終了する
        _mapi().MAPIUninitialize()


def logon(profile: str):
    """プロファイルへの新しいセッション。MAPI_NEW_SESSION が無いと既存の共有セッションが返りうる。"""
    m = _mapi()
    flags = m.MAPI_EXTENDED | m.MAPI_NO_MAIL | m.MAPI_NEW_SESSION | m.MAPI_EXPLICIT_PROFILE
    return m.MAPILogonEx(0, profile, None, flags)


def _query(table, cols) -> list:
    table.SetColumns(cols, 0)
    out = []
    while rows := table.QueryRows(QUERY_BATCH, 0):
        out.extend(rows)
    return out


def profile_names() -> list[str]:
    admin = _mapi().MAPIAdminProfiles(0)
    rows = _query(admin.GetProfileTable(0), (PR_DISPLAY_NAME_A,))
    return sorted(r[0][1].decode("mbcs") for r in rows)


def cleanup_temp_profiles() -> list[str]:
    """異常終了で残った本ツールの一時プロファイル（PSTMigration-tmp-*）だけを消す。"""
    admin = _mapi().MAPIAdminProfiles(0)
    removed = [n for n in profile_names() if n.startswith(TEMP_PREFIX)]
    for n in removed:
        with contextlib.suppress(Exception):
            admin.DeleteProfile(n, 0)
    return removed


class TempProfile:
    """PST 1個だけを入れた一時プロファイルとセッション（§16.5.1）。既定プロファイルにはしない。"""

    def __init__(self, pst: Path, *, create: bool, name: str | None = None):
        if create and pst.exists():
            raise FileExistsError(pst)
        if not create and not pst.exists():
            raise FileNotFoundError(pst)
        import pywintypes

        m = _mapi()
        self.name = name or f"{TEMP_PREFIX}{os.getpid()}-{next(_seq)}"  # 省略時は PSTMigration-tmp-{PID}-{連番}
        self._admin = m.MAPIAdminProfiles(0)
        self._admin.CreateProfile(self.name, None, 0, 0)
        self.session = None
        try:
            svc = self._admin.AdminServices(self.name, None, 0, 0)
            svc.CreateMsgService("MSUPST MS", "PSTMigration", 0, 0)
            rows = _query(svc.GetMsgServiceTable(0), (PR_SERVICE_UID, PR_SERVICE_NAME_A))
            uid = next(r[0][1] for r in rows if r[1][1] == b"MSUPST MS")
            props = [(PR_PST_PATH_W, str(pst)), (PR_DISPLAY_NAME_W, pst.stem)]
            svc.ConfigureMsgService(pywintypes.IID(uid, True), 0, 0, props)
            del svc
            self.session = logon(self.name)
        except Exception:
            self.close()
            raise

    def open_store(self, *, writable: bool) -> MapiStore:
        m = _mapi()
        eid = _query(self.session.GetMsgStoresTable(0), (PR_ENTRYID,))[0][0][1]
        flags = m.MDB_NO_DIALOG | (m.MDB_WRITE if writable else 0)
        return MapiStore(self.session.OpenMsgStore(0, eid, None, flags), writable=writable)

    def close(self) -> None:
        """セッションをログオフしてプロファイルを消す。呼ぶ前に、このセッションの MapiStore 等は手放しておく。"""
        gc.collect()
        if self.session is not None:
            with contextlib.suppress(Exception):
                self.session.Logoff(0, 0, 0)
            self.session = None
        gc.collect()
        with contextlib.suppress(Exception):
            self._admin.DeleteProfile(self.name, 0)

    def abandon(self) -> None:
        """テスト用：異常終了を模す（ログオフもプロファイル削除もしない）。"""
        self.session = None
        gc.collect()


class MapiStore:
    """ストア1つ。フォルダーは IPM サブツリー（Outlook に見えるルート）からのパス（\\受信トレイ\\…）で扱う。"""

    def __init__(self, store, *, writable: bool):
        self.store, self.writable = store, writable
        self._root_eid = store.GetProps((PR_IPM_SUBTREE_ENTRYID,), 0)[1][0][1]
        self._folders: dict[str, object] = {}

    def open(self, eid: bytes):
        return self.store.OpenEntry(eid, None, _mapi().MAPI_MODIFY if self.writable else 0)

    def _children(self, folder) -> list[tuple]:
        cols = (PR_ENTRYID, PR_DISPLAY_NAME_W, PR_CONTAINER_CLASS_W, PR_FOLDER_TYPE)
        return [tuple(value(p) for p in r) for r in _query(folder.GetHierarchyTable(0), cols)]

    def walk(self) -> list[tuple[str, bytes, str, str | None]]:
        """(パス, EntryID, 'mail'|'other', コンテナークラス) を親から順に。検索フォルダーは配下ごと除外。"""
        out: list = []

        def visit(folder, names):
            for eid, name, cc, ft in self._children(folder):
                kind = folder_kind(cc, ft)
                if kind == "search":
                    continue
                n = (*names, name)
                out.append(("\\" + "\\".join(n), eid, kind, cc))
                visit(self.open(eid), n)

        visit(self.open(self._root_eid), ())
        return out

    def folder(self, path: str):
        f = self._folders.get(path)
        if f is None:
            f = self.open(self._root_eid)
            for name in path.strip("\\").split("\\"):
                eid = next((e for e, n, _, _ in self._children(f) if n == name), None)
                if eid is None:
                    raise KeyError(path)
                f = self.open(eid)
            self._folders[path] = f
        return f

    def ensure_folder(self, path: str, container_class: str | None):
        """書き込み用：無ければ作る（親も）。"""
        f = self._folders.get(path)
        if f is None:
            f, sofar = self.open(self._root_eid), ""
            for name in path.strip("\\").split("\\"):
                sofar += "\\" + name
                if sofar in self._folders:
                    f = self._folders[sofar]
                    continue
                f = f.CreateFolder(FOLDER_GENERIC, name, None, None, OPEN_IF_EXISTS)
                if container_class:
                    f.SetProps([(PR_CONTAINER_CLASS_W, container_class)])
                    f.SaveChanges(0)
                self._folders[sofar] = f
        return f

    def rows(self, folder) -> list:
        return _query(folder.GetContentsTable(0), MSG_COLUMNS)

    def digest(self, eid: bytes) -> str:
        return digest(self._parts(self.store.OpenEntry(eid, None, 0)))

    def _parts(self, msg) -> list[bytes]:
        import pythoncom

        def read(obj, tag) -> bytes:
            try:
                s = obj.OpenProperty(tag, pythoncom.IID_IStream, 0, 0)
            except Exception:  # 値が無い
                return b""
            out = []
            while chunk := s.Read(1 << 20):
                out.append(chunk)
            return b"".join(out)

        parts = [tag.to_bytes(4, "little") + read(msg, tag) for tag in (PR_BODY_W, PR_HTML, PR_RTF_COMPRESSED)]
        for n in sorted(r[0][1] for r in _query(msg.GetAttachmentTable(0), (PR_ATTACH_NUM,))):
            a = msg.OpenAttach(n, None, 0)
            method, name = (value(p) for p in a.GetProps((PR_ATTACH_METHOD, PR_ATTACH_LONG_FILENAME_W), 0)[1])
            parts.append(f"{method}|{name}".encode())
            if method == ATTACH_EMBEDDED_MSG:
                parts.extend(self._parts(a.OpenProperty(PR_ATTACH_DATA_OBJ, _mapi().IID_IMessage, 0, 0)))
            else:
                parts.append(read(a, PR_ATTACH_DATA_BIN))
        return parts


def open_user_store(session, store_id: str) -> MapiStore:
    """ユーザーのプロファイルにある元PSTを、別セッションから読み取り専用で開く（MDB_WRITE を付けない）。"""
    return MapiStore(session.OpenMsgStore(0, bytes.fromhex(store_id), None, _mapi().MDB_NO_DIALOG), writable=False)


def scan(store: MapiStore, cutoff_utc: datetime, *, log_subject: bool = False) -> dict:
    """元ストアの調査（§16.4）。読み取り専用で開いたストアを渡す。

    戻り値: scan（現行 scan.json と同じ形）、items（旧・新の計画）、all_rows（全フォルダーの全行）、
    mail_folders（(パス, コンテナークラス)。最初の束に全階層を作るため）
    """
    folders, problems, all_rows, mail_folders = [], [], [], []
    items: dict[str, list[dict]] = {"old": [], "new": []}
    for path, eid, kind, cc in store.walk():
        raw = store.rows(store.open(eid))
        rec, its, rows, probs = classify_rows(path, raw, cutoff_utc, log_subject=log_subject)
        all_rows.extend(rows)
        if kind == "other":  # 予定表・連絡先など：中身は移行しない（件数だけ記録。現行 §5.1 と同じ）
            rec = {**dict.fromkeys(rec, 0), "path": path, "kind": "other", "non_mail": len(raw)}
        else:
            mail_folders.append((path, cc))
            for b in items:
                items[b].extend(its[b])
            problems.extend(probs[: MAX_PROBLEMS - len(problems)])
        folders.append(rec)
    return {
        "scan": {"folders": folders, "problems": problems},
        "items": items,
        "all_rows": all_rows,
        "mail_folders": mail_folders,
    }


def add_hashes(store: MapiStore, items: list[dict], need: set[tuple]) -> None:
    """SearchKey＋指紋が重複する組のメールにだけ、本文・添付のハッシュ h を付ける（§16.6）。"""
    for i in items:
        if (i["k"], i["p"]) in need:
            i["h"] = store.digest(bytes.fromhex(i["e"]))


class ChunkWriter:
    """束PSTを一時プロファイルで新規作成し、元ストアから CopyMessages（移動なし）でコピーする。"""

    def __init__(self, pst: Path, source: MapiStore, folders: list[tuple[str, str | None]] | None, pause_ms: int = 0):
        self._pause = pause_ms / 1000
        self._tp = TempProfile(pst, create=True)
        try:
            self._dst = self._tp.open_store(writable=True)
            self._classes = dict(folders or [])
            for path, cc in folders or []:  # 最初の束：空フォルダーを含む全階層
                self._dst.ensure_folder(path, cc)
        except Exception:
            self._dst = None
            self._tp.close()
            raise
        self._source = source

    def copy(self, items: list[dict]) -> None:
        for path, group in itertools.groupby(items, key=lambda i: i["f"]):
            eids = [bytes.fromhex(i["e"]) for i in group]
            dst = self._dst.ensure_folder(path, self._classes.get(path, "IPF.Note"))
            self._source.folder(path).CopyMessages(eids, None, dst, 0, None, 0)
        time.sleep(self._pause)  # 元PSTを開いている利用者の Outlook の操作を妨げないよう少し譲る

    def close(self) -> None:
        self._dst = None
        self._tp.close()


def read_chunk(pst: Path, need: set[tuple]) -> dict:
    """束PSTの中身（照合用の行。ハッシュが要る組は本文・添付のハッシュ付き）とフォルダー一覧。読み取り専用。"""
    tp = TempProfile(pst, create=False)
    try:
        st = tp.open_store(writable=False)
        rows, folders = [], []
        for path, eid, _, _ in st.walk():
            folders.append(path)
            for props in st.rows(st.open(eid)):
                _, r = _row(path, props)
                if (r["k"], r["p"]) in need:
                    r["h"] = st.digest(bytes.fromhex(r["e"]))
                rows.append(r)
        del st
        return {"rows": rows, "folders": folders}
    finally:
        tp.close()
