"""v1.2.4 の診断情報と 32ビット版 Outlook への委譲（設計書 §18.5、§18.6）。"""

import csv
import json
import subprocess
import sys

import pytest

from pstmigration import __main__ as entry
from pstmigration import pipeline, report, transfer, winenv
from pstmigration.state import StepError


def test_holders_names_the_process_that_locks_the_file(tmp_path):
    f = tmp_path / "a.pst"
    f.write_bytes(b"x")
    h = transfer.open_exclusive(f)
    try:
        who = winenv.holders(f)
        err = pipeline.locked_error(f, "解放しません")
    finally:
        h.Close()
    assert any(w.lower().startswith("python") for w in who)
    assert err.code == "PST_LOCKED" and err.transient and "使用中: python" in err.message.lower()
    assert winenv.holders(f) == [] and winenv.holders(tmp_path / "none.pst") == []


def test_snapshot_reports_bitness_and_disk(tmp_path):
    (tmp_path / "w").mkdir()
    (tmp_path / "w" / "b.bin").write_bytes(b"0" * 1024)
    env = winenv.snapshot(tmp_path / "w")
    assert env["tool_bits"] in (32, 64) and env["free_gb"] > 0 and env["work_gb"] >= 0
    assert {"outlook_bitness", "office_version", "windows"} <= env.keys()


@pytest.fixture
def frozen64(tmp_path, monkeypatch):
    exe = tmp_path / "PSTMigration.exe"
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "executable", str(exe))
    monkeypatch.setattr(winenv, "TOOL_BITS", 64)
    monkeypatch.setattr(winenv, "outlook", lambda: {"bitness": "x86", "version": "16.0"})
    calls = []

    def run(args, **kw):
        calls.append(args)
        return subprocess.CompletedProcess(args, 7)

    monkeypatch.setattr(subprocess, "run", run)
    return tmp_path, calls


def test_64bit_exe_hands_over_to_32bit_exe_for_32bit_outlook(frozen64, monkeypatch):
    tmp, calls = frozen64
    assert entry.delegate_to_32bit(["--scan-only"]) is None  # 32ビット版が無ければ自分で続ける
    (tmp / "PSTMigration32.exe").write_bytes(b"")
    assert entry.delegate_to_32bit(["--scan-only"]) == 7
    assert calls == [[str(tmp / "PSTMigration32.exe"), "--scan-only"]]
    monkeypatch.setattr(winenv, "outlook", lambda: {"bitness": "x64", "version": "16.0"})
    assert entry.delegate_to_32bit([]) is None
    monkeypatch.setattr(winenv, "TOOL_BITS", 32)  # 32ビット版自身は再委譲しない
    monkeypatch.setattr(winenv, "outlook", lambda: {"bitness": "x86", "version": "16.0"})
    assert entry.delegate_to_32bit([]) is None


def test_transfer_network_error_records_stage_and_progress(tmp_path, monkeypatch):
    src, dst = tmp_path / "a.pst", tmp_path / "nas" / "a.pst"
    src.write_bytes(b"abc" * 1000)
    real = transfer.sha256_file

    def sha(p):
        if p != src:  # NAS 上の .partial の読み直しで切断
            raise OSError(22, "Invalid argument")
        return real(p)

    monkeypatch.setattr(transfer, "sha256_file", sha)
    with pytest.raises(StepError) as ex:
        transfer.upload(src, dst, real(src), 0)
    d = ex.value.details
    assert ex.value.code == "NETWORK" and ex.value.transient
    assert (d["stage"], d["sent"], d["size"], d["offset"]) == ("verify", 3000, 3000, 0)


def test_status_carries_env_and_recent_logs_and_aggregate_lists_pcs(tmp_path):
    logs = tmp_path / "logs"
    logs.mkdir()
    for i in range(report.LOGS_TO_PUBLISH + 2):
        (logs / f"run-202611{i:02d}-000000-1.jsonl").write_text("{}", encoding="utf-8")
    data = {
        "mailbox": "a@x.jp",
        "env": {"tool_bits": 32, "outlook_bitness": "x86", "free_gb": 4.7, "mapi_error": None},
        "contacts": {"last_error": {"code": "OUTLOOK_COM", "message": "E_FAIL"}},
        "psts": {},
    }
    nas = tmp_path / "nas"
    report.publish_status(nas, data, "a", "PC1", logs)
    st = json.loads((nas / "users" / "a@x.jp" / "PC1" / "status.json").read_text(encoding="utf-8"))
    assert st["env"]["outlook_bitness"] == "x86"
    published = sorted(p.name for p in (nas / "users" / "a@x.jp" / "PC1" / "logs").iterdir())
    assert published == sorted(p.name for p in logs.iterdir())[-report.LOGS_TO_PUBLISH :]
    out = report.aggregate(nas)
    with open(out / "pcs.csv", encoding="utf-8-sig") as f:
        row = next(csv.DictReader(f))
    assert (row["computer"], row["outlook_bitness"], row["free_gb"], row["contacts_error_code"]) == (
        "PC1",
        "x86",
        "4.7",
        "OUTLOOK_COM",
    )


def test_orphan_work_folders_of_completed_or_unknown_psts_are_removed(tmp_path):
    for name in ("done", "failed", "gone", "contacts"):
        (tmp_path / name).mkdir()
        (tmp_path / name / "x.pst").write_bytes(b"x")
    psts = {"done": {"status": "COMPLETED"}, "failed": {"status": "FAILED"}}
    pipeline.purge_orphan_work(tmp_path, psts)
    assert sorted(p.name for p in tmp_path.iterdir()) == ["contacts", "failed"]
    pipeline.purge_orphan_work(tmp_path / "none", psts)  # 作業フォルダーがまだ無い
