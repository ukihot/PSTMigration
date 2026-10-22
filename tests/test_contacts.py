"""連絡先は利用中の Outlook に出力PSTを接続せず、閉じてから転送する。"""

from contextlib import nullcontext
from types import SimpleNamespace

import pytest

from pstmigration import config, mapilive, pipeline
from pstmigration.state import StepError


@pytest.mark.parametrize("fail", [False, True])
def test_backup_uses_isolated_export_and_keeps_previous_backup_on_failure(tmp_path, monkeypatch, fail):
    monkeypatch.setenv("COMPUTERNAME", "PC1")
    cfg = config.Config(work_dir=str(tmp_path / "work"), server_root=str(tmp_path / "nas"))
    r = pipeline.Runner(cfg, tmp_path / "home")
    r.profile = "Outlook"
    r.st.data["mailbox"] = "a@x.jp"
    sources = [{"store_id": "00", "name": "1. Legacy"}]
    r.ol = SimpleNamespace(contact_sources=lambda work: sources, alive=lambda: True)
    remote = cfg.server_path / "users" / "a@x.jp" / "PC1" / "contacts_PC1.pst"
    remote.parent.mkdir(parents=True)
    remote.write_bytes(b"previous")

    def export(pst, profile, stores):
        assert profile == "Outlook" and stores == sources
        pst.write_bytes(b"new contacts")
        if fail:
            raise StepError("CONTACTS_MISMATCH", "copy incomplete")
        return 3

    monkeypatch.setattr(mapilive, "export_contacts", export, raising=False)
    monkeypatch.setattr(r, "_let_go", lambda: pytest.fail("利用中の Outlook を終了させない"))
    r._backup_contacts()
    c = r.st.data["contacts"]
    if fail:
        assert c["last_error"]["code"] == "CONTACTS_MISMATCH"
        assert remote.read_bytes() == b"previous"
    else:
        assert c["count"] == 3 and c["last_error"] is None
        assert remote.read_bytes() == b"new contacts"
        assert not (cfg.work_path / "contacts").exists()


def test_contact_sources_include_imap_but_exclude_exchange_and_work_files(tmp_path, monkeypatch):
    from pstmigration.outlook import Outlook

    work = tmp_path / "work"
    stores = [
        SimpleNamespace(FilePath=str(tmp_path / "legacy.pst"), ExchangeStoreType=3, StoreID="01", DisplayName="PST"),
        SimpleNamespace(FilePath=str(tmp_path / "imap.ost"), ExchangeStoreType=3, StoreID="02", DisplayName="IMAP"),
        SimpleNamespace(FilePath=str(tmp_path / "exchange.ost"), ExchangeStoreType=0),
        SimpleNamespace(FilePath=str(work / "contacts" / "old.pst"), ExchangeStoreType=3),
        SimpleNamespace(FilePath="", ExchangeStoreType=3),
    ]
    ol = object.__new__(Outlook)
    monkeypatch.setattr(ol, "_stores", lambda: iter(stores))
    assert ol.contact_sources(work) == [{"store_id": "01", "name": "1. PST"}, {"store_id": "02", "name": "2. IMAP"}]


@pytest.mark.parametrize("copied, error", [(0, None), (1, None), (2, RuntimeError("copy failed"))])
def test_copy_checks_counts_and_uses_copy_flags(monkeypatch, copied, error):
    """部分成功・MAPI例外を成功扱いにせず、移動フラグも付けない。"""
    events = []
    rows = [[(mapilive.PR_ENTRYID, b"1")], [(mapilive.PR_ENTRYID, b"2")]]

    def copy(eids, iid, dst, ui, progress, flags):
        assert eids == [b"1", b"2"] and flags == 0
        if error:
            raise error

    sf = SimpleNamespace(GetContentsTable=lambda flags: rows, CopyMessages=copy)
    df = SimpleNamespace(GetContentsTable=lambda flags: rows[:copied])
    source = SimpleNamespace(walk=lambda: [("\\Contacts", b"f", "other", "IPF.Contact")], open=lambda eid: sf)
    dst = SimpleNamespace(ensure_folder=lambda path, cc: df)
    monkeypatch.setattr(mapilive, "_query", lambda table, cols: table)
    monkeypatch.setattr(mapilive, "open_user_store", lambda session, sid: source)
    monkeypatch.setattr(mapilive, "Mapi", nullcontext)
    monkeypatch.setattr(mapilive, "logon", lambda profile: SimpleNamespace(Logoff=lambda *a: events.append("logoff")))
    monkeypatch.setattr(
        mapilive,
        "TempProfile",
        lambda *a, **kw: SimpleNamespace(open_store=lambda **kw: dst, close=lambda: events.append("close")),
    )
    with pytest.raises(RuntimeError if error else StepError) as exc:
        mapilive.export_contacts(None, "Outlook", [{"store_id": "00", "name": "1. PST"}])
    if not error:
        assert exc.value.code == "CONTACTS_MISMATCH"
    assert events == ["close", "logoff"]
