"""ARCHIVED / TRANSFERRED で中断したスナップショットの復旧。NAS は一時フォルダーを使う。"""

import csv
import json

import pytest

from pstmigration import config, pipeline, report, transfer
from pstmigration.state import StepError, new_entry, read_json


class Crash(BaseException):
    pass


@pytest.fixture
def archived(tmp_path, monkeypatch):
    monkeypatch.setenv("COMPUTERNAME", "PC1")
    cfg = config.Config(server_root=str(tmp_path / "nas"), work_dir=str(tmp_path / "work"), release_wait_sec=0)
    r = pipeline.Runner(cfg, tmp_path / "home")
    e = new_entry("key1", str(tmp_path / "original.pst"), False)
    e.update(checkpoint="ARCHIVED", remote_dir=r"users\a@x.jp\PC1\unit", counts={"old": 1, "new": 2, "unmigrated": 0})
    r.st.psts[e["key"]] = e
    r.st.data["mailbox"] = "a@x.jp"
    work = cfg.work_path / e["key"]
    work.mkdir(parents=True)
    for kind, name, content in (("new_pst", r.names.new_pst, b"new mail"), ("old_7z", r.names.old_7z, b"old mail")):
        path = work / name
        path.write_bytes(content)
        e["outputs"][kind] = {"name": name, "size": len(content), "sha256": transfer.sha256_file(path)}
    r.st.save()
    monkeypatch.setattr(r, "_split", lambda *a: pytest.fail("検証済みの分割をやり直した"))
    monkeypatch.setattr(r, "_archive", lambda *a: pytest.fail("検証済みの7zを再生成した"))
    return r, e, work


def reload_runner(r):
    resumed = pipeline.Runner(r.cfg, r.home)
    return resumed, resumed.st.psts["key1"]


def test_archived_crash_resumes_without_rebuilding(archived, monkeypatch):
    r, e, work = archived

    def crash(*a):
        raise Crash

    monkeypatch.setattr(r, "_transfer", crash)
    with pytest.raises(Crash):
        r._process(e)
    resumed, e = reload_runner(r)
    resumed._process(e)
    assert e["status"] == "COMPLETED"
    remote = r.cfg.server_path / e["remote_dir"]
    assert sorted(p.suffix for p in remote.iterdir()) == [".7z", ".json", ".pst"]
    assert not work.exists()


def test_transfer_disconnect_reuses_already_uploaded_files(archived, monkeypatch):
    r, e, work = archived
    upload = transfer.upload

    def disconnect(src, dst, *args):
        upload(src, dst, *args)
        raise StepError("NETWORK", "転送後に切断", transient=True)

    monkeypatch.setattr(transfer, "upload", disconnect)
    r._process(e)
    assert (e["status"], e["checkpoint"]) == ("RETRY_PENDING", "ARCHIVED")
    r.st.save()
    remote = r.cfg.server_path / e["remote_dir"] / e["outputs"]["new_pst"]["name"]
    before = remote.stat().st_mtime_ns
    monkeypatch.setattr(transfer, "upload", upload)
    resumed, e = reload_runner(r)
    resumed._process(e)
    assert e["status"] == "COMPLETED" and remote.stat().st_mtime_ns == before
    assert read_json(remote.parent / "report.json", {})["errors"] == []


@pytest.mark.parametrize("damage", ["same_size", "missing", "metadata"])
def test_invalid_archived_artifacts_are_preserved_and_not_completed(archived, damage):
    r, e, work = archived
    local = work / e["outputs"]["new_pst"]["name"]
    if damage == "same_size":
        local.write_bytes(b"bad mail")
    elif damage == "missing":
        local.unlink()
    else:
        e["outputs"] = {}
    r._process(e)
    assert e["status"] == "FAILED" and e["checkpoint"] == "ARCHIVED"
    assert e["last_error"]["code"].startswith("ARTIFACT_")
    assert work.exists() and (work / r.names.old_7z).read_bytes() == b"old mail"
    assert not list(r.cfg.server_path.rglob("report.json"))


def test_archived_missing_local_copy_can_use_verified_nas_copy(archived):
    r, e, work = archived
    o = e["outputs"]["new_pst"]
    local, remote = work / o["name"], r.cfg.server_path / e["remote_dir"] / o["name"]
    transfer.upload(local, remote, o["sha256"], 0)
    local.unlink()
    r._process(e)
    assert e["status"] == "COMPLETED" and remote.read_bytes() == b"new mail"


@pytest.mark.parametrize("checkpoint", ["ARCHIVED", "TRANSFERRED"])
@pytest.mark.parametrize(
    "counts",
    [
        {"old": -1, "new": 2},
        {"old": 1, "new": -1},
        {"old": True, "new": 2},
        {"old": 1.0, "new": 2},
        {"old": "1", "new": 2},
        {"old": None, "new": 2},
        {"old": 1, "new": 2, "unmigrated": -1},
        {"old": 1},
        [1, 2],
        None,
    ],
)
def test_resume_rejects_invalid_counts_without_removing_artifacts(archived, checkpoint, counts):
    r, e, work = archived
    if checkpoint == "TRANSFERRED":
        r._transfer(e, work)
    e["counts"] = counts
    r.st.save()
    resumed, e = reload_runner(r)
    resumed._process(e)
    assert (e["status"], e["checkpoint"]) == ("FAILED", checkpoint)
    assert e["last_error"]["code"] == "ARTIFACT_METADATA" and not e["last_error"]["retryable"]
    assert work.exists() and (work / r.names.old_7z).read_bytes() == b"old mail"
    assert not list(r.cfg.server_path.rglob("report.json"))


@pytest.mark.parametrize("old,new", [(0, 2), (1, 0), (0, 0)])
def test_resume_accepts_zero_counts_without_an_output_for_that_side(archived, old, new):
    r, e, work = archived
    e["counts"].update(old=old, new=new)
    for kind, count in (("old_7z", old), ("new_pst", new)):
        if not count:
            del e["outputs"][kind]
    r._process(e)
    assert e["status"] == "COMPLETED" and not work.exists()


def test_failed_report_write_resumes_without_outlook_or_retransfer(archived, monkeypatch):
    r, e, work = archived
    write = pipeline.write_json

    def disconnect(path, data):
        if path.name == "report.json":
            raise OSError("NAS unavailable")
        write(path, data)

    monkeypatch.setattr(pipeline, "write_json", disconnect)
    r._process(e)
    assert (e["status"], e["checkpoint"]) == ("RETRY_PENDING", "TRANSFERRED")
    assert e["last_error"]["phase"] == "FINALIZING"
    r.st.save()
    report.publish_status(r.cfg.server_path, r.st.data, "a", "PC1")
    out = report.aggregate(r.cfg.server_path)
    assert read_json(out / "errors.json", [])[0]["error"]["code"] == "NETWORK"
    monkeypatch.setattr(pipeline, "write_json", write)
    monkeypatch.setattr(transfer, "upload", lambda *a: pytest.fail("完了登録の再開で再転送した"))
    resumed, e = reload_runner(r)
    resumed._process(e)
    assert e["status"] == "COMPLETED" and not work.exists()
    assert read_json(r.cfg.server_path / e["remote_dir"] / "report.json", {})["errors"] == []
    resumed.st.save()
    persisted = read_json(resumed.st.path, {})["psts"][e["key"]]
    assert persisted["status"] == "COMPLETED" and persisted["last_error"] is None
    report.publish_status(resumed.cfg.server_path, resumed.st.data, "a", "PC1")
    status = read_json(resumed.cfg.server_path / "users" / "a@x.jp" / "PC1" / "status.json", {})
    assert status["psts"][0]["status"] == "COMPLETED" and status["psts"][0]["last_error"] is None
    assert read_json(report.aggregate(resumed.cfg.server_path) / "errors.json", None) == []


@pytest.mark.parametrize("damage", ["same_size", "different_size", "missing"])
def test_transferred_remote_corruption_is_not_completed(archived, monkeypatch, damage):
    r, e, work = archived
    r._transfer(e, work)
    o = e["outputs"]["new_pst"]
    remote = r.cfg.server_path / e["remote_dir"] / o["name"]
    if damage == "missing":
        remote.unlink()
    else:
        remote.write_bytes(b"bad mail" if damage == "same_size" else b"bad")
    r.st.save()
    monkeypatch.setattr(transfer, "upload", lambda *a: pytest.fail("不一致を再転送で隠した"))
    resumed, e = reload_runner(r)
    resumed._process(e)
    assert e["status"] == "FAILED" and e["last_error"]["code"] == "ARTIFACT_MISMATCH"
    assert work.exists() and not list(r.cfg.server_path.rglob("report.json"))


def test_unknown_exception_is_terminal_and_published_with_cause(archived, monkeypatch):
    r, e, work = archived

    def boom(*a):
        try:
            raise ValueError("inner failure")
        except ValueError as ex:
            raise RuntimeError("outer failure") from ex

    monkeypatch.setattr(r, "_transfer", boom)
    r._process(e)
    err = e["last_error"]
    assert e["status"] == "FAILED" and e["attempts"] == 0 and not err["retryable"]
    assert err["exception_type"] == "builtins.RuntimeError" and err["checkpoint"] == "ARCHIVED"
    assert "ValueError: inner failure" in err["traceback"] and "RuntimeError: outer failure" in err["traceback"]
    report.publish_status(r.cfg.server_path, r.st.data, "a", "PC1")
    out = report.aggregate(r.cfg.server_path)
    assert read_json(out / "errors.json", [])[0]["error"] == err
    with (out / "ledger.csv").open(encoding="utf-8-sig", newline="") as f:
        row = next(csv.DictReader(f))
    assert row["last_error_message"] == "RuntimeError('outer failure')"
    assert row["last_error_exception_type"] == "builtins.RuntimeError"


def test_reset_json_on_nas_resets_only_this_pcs_old_failures_once(archived, tmp_path):
    r, e, work = archived
    e.update(status="FAILED", last_error={"code": "UNEXPECTED", "at": "2026-10-28T08:50:29+09:00"})
    other = {"status": "FAILED", "last_error": {"at": "2026-10-28T00:00:00+09:00"}}
    later = {"status": "FAILED", "last_error": {"at": "2026-11-13T00:00:00+09:00"}}  # 一覧を書いた後の失敗
    psts = {"key1": e, "key2": other, "key3": later}
    req = tmp_path / "reset.json"
    req.write_text(
        json.dumps({"before": "2026-11-12T18:00:00+09:00", "targets": ["pc1/key1", "PC2/key2", "PC1/key3"]}),
        encoding="utf-8",
    )
    assert pipeline.apply_resets(psts, req, "PC1") == ["key1"]
    assert (e["status"], e["last_error"], other["status"], later["status"]) == ("PENDING", None, "FAILED", "FAILED")
    r._process(e)  # ARCHIVED から転送だけやり直す
    assert e["status"] == "COMPLETED" and e["generation"] == 1
    e.update(status="FAILED", last_error={"at": "2026-11-13T09:00:00+09:00"})  # 戻した後の再失敗は戻さない
    assert pipeline.apply_resets(psts, req, "PC1") == []
    assert pipeline.apply_resets(psts, tmp_path / "none.json", "PC1") == []
    req.write_text("{broken", encoding="utf-8")
    assert pipeline.apply_resets(psts, req, "PC1") == []


def test_snapshot_reset_keeps_checkpoint_and_artifacts(archived):
    r, e, work = archived
    e["status"] = "FAILED"
    r.st.save()
    assert pipeline.reset(r.home, "key1")
    resumed, e = reload_runner(r)
    assert e["checkpoint"] == "ARCHIVED" and work.exists()
    resumed._process(e)
    assert e["status"] == "COMPLETED" and e["generation"] == 1
