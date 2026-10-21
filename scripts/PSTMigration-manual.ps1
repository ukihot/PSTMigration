# 手動実行：GPO のタイミングを待たずに ①検出 → ②コピー → ③分割・7z・転送 を続けて行う。
# Outlook を閉じた状態で、同じフォルダーの PSTMigration-manual.cmd をダブルクリックする。
# ログオフ時のコピー上限（logoff_snapshot_budget_gb）は無視するので、大きいPSTもこれで処理できる。

$ErrorActionPreference = "Stop"
$exe = Join-Path $PSScriptRoot "PSTMigration.exe"
$home_ = Join-Path $env:LOCALAPPDATA "PSTMigration"
$cfg = Join-Path $home_ "config-manual.json"

function Step($msg) { Write-Host "`n[$(Get-Date -Format HH:mm:ss)] $msg" -ForegroundColor Cyan }
function Fail($msg) { Write-Host "`n$msg" -ForegroundColor Red; exit 1 }
function Run([string[]]$a) {
    $p = Start-Process $exe -ArgumentList $a -Wait -PassThru
    if ($p.ExitCode -ne 0) { Fail "PSTMigration.exe が異常終了しました（exit=$($p.ExitCode)）。$home_\logs の最新ログを確認してください。" }
}

function Wait-OutlookExit($minutes, $msg) {
    $deadline = (Get-Date).AddMinutes($minutes)
    while (Get-Process OUTLOOK -ErrorAction SilentlyContinue) {
        if ((Get-Date) -gt $deadline) { Fail $msg }
        Start-Sleep 5
    }
}

if (Get-Process PSTMigration -ErrorAction SilentlyContinue) { Fail "PSTMigration が実行中です。終わってからもう一度実行してください。" }
$ol = Get-Process OUTLOOK -ErrorAction SilentlyContinue
if ($ol | Where-Object MainWindowHandle -ne 0) { Fail "Outlook を閉じてから実行してください。" }
if ($ol) {
    # 画面のない Outlook ＝ 途中で止めた PSTMigration が裏で起動した残り。普通に終了させる
    Step "裏で動いている Outlook を終了します"
    try { [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application").Quit() } catch {}
    Wait-OutlookExit 2 "Outlook が終了しません。タスクマネージャーで OUTLOOK.EXE を終了してから、もう一度実行してください。"
}

# bin\config.json は GPO で上書きされるので、上限を外したコピーを別に作る
$t = [IO.File]::ReadAllText((Join-Path $PSScriptRoot "config.json"))
$t = $t -replace '"logoff_snapshot_budget_gb"\s*:\s*[\d.]+', '"logoff_snapshot_budget_gb": 100000'
[IO.File]::WriteAllText($cfg, $t)

Step "① PSTを検出（Outlook を裏で起動します）"
Run @("--config", $cfg)

# ①で裏起動した Outlook が終わるのを待つ（起動中だとコピーできない）
Wait-OutlookExit 3 "Outlook が終了しません。開いていたら閉じてください。タスクマネージャーで OUTLOOK.EXE が消えてから、もう一度実行してください。"

Step "② PSTをコピー（20GBで数分。Outlook はまだ開かないでください）"
Run @("--snapshot", "--config", $cfg)

Step "③ 分割 → 7z → NAS転送（数時間かかることがあります。Outlook は開いて使ってよい。PCはスリープ・シャットダウンしない）"
Run @("--config", $cfg)

Step "結果"
$st = [IO.File]::ReadAllText((Join-Path $home_ "state.json")) | ConvertFrom-Json
$st.psts.PSObject.Properties.Value | ForEach-Object {
    [pscustomobject]@{ 状態 = $_.status; 段階 = $_.checkpoint; エラー = $_.last_error.code; PST = $_.source_path }
} | Format-Table -AutoSize -Wrap
if ($st.contacts) { Write-Host "連絡先: $($st.contacts.count) 件（$($st.contacts.backup_at)）  エラー: $($st.contacts.last_error.code)" }
Write-Host "`n状態が COMPLETED なら完了。PENDING / RETRY_PENDING ならもう一度実行、FAILED ならエラーを確認。"
