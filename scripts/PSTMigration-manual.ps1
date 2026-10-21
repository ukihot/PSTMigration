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

# 画面のない Outlook ＝ PSTMigration が裏で起動したもの。普通に終了させる
# （強制終了はしない。ユーザーの元PSTも開いているので、壊さないため）
function Stop-HeadlessOutlook {
    $ol = Get-Process OUTLOOK -ErrorAction SilentlyContinue
    if (-not $ol -or ($ol | Where-Object MainWindowHandle -ne 0)) { return }
    Step "裏で動いている Outlook を終了します"
    try { [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application").Quit() } catch {}
    Wait-OutlookExit 2 "Outlook が終了しません。タスクマネージャーで OUTLOOK.EXE を終了してから、もう一度実行してください。"
}

if (Get-Process PSTMigration -ErrorAction SilentlyContinue) { Fail "PSTMigration が実行中です。終わってからもう一度実行してください。" }
if (Get-Process OUTLOOK -ErrorAction SilentlyContinue | Where-Object MainWindowHandle -ne 0) { Fail "Outlook を閉じてから実行してください。" }
Stop-HeadlessOutlook

$finished = $false
try {

# bin\config.json は GPO で上書きされるので、上限を外したコピーを別に作る
$t = [IO.File]::ReadAllText((Join-Path $PSScriptRoot "config.json"))
$t = $t -replace '"logoff_snapshot_budget_gb"\s*:\s*[\d.]+', '"logoff_snapshot_budget_gb": 100000'
[IO.File]::WriteAllText($cfg, $t)

Step "① PSTを検出（途中で止まっていたPSTがあれば、その続きの分割・7z・転送もここで行うので数時間かかることがあります）"
Run @("--config", $cfg)

# ①で裏起動した Outlook が終わるのを待つ（起動中だとコピーできない）
Wait-OutlookExit 3 "Outlook が終了しません。開いていたら閉じてください。タスクマネージャーで OUTLOOK.EXE が消えてから、もう一度実行してください。"

Step "② PSTをコピー（20GBで数分。Outlook はまだ開かないでください）"
Run @("--snapshot", "--config", $cfg)

Step "③ 分割 → 7z → NAS転送（数時間かかることがあります。Outlook は開いて使ってよい。PCはスリープ・シャットダウンしない）"
Run @("--config", $cfg)
$finished = $true
} finally {
    # Ctrl+C・途中のエラーで止まったとき：EXE と 7z を止め、裏の Outlook を終了させる（次回は続きから再開）
    if (-not $finished) {
        Write-Host "`n中断しました。後片付けをしています…" -ForegroundColor Yellow
        Stop-Process -Name PSTMigration, 7z -Force -ErrorAction SilentlyContinue
        Stop-HeadlessOutlook
        Write-Host "後片付けが終わりました。もう一度実行すると、止めたところから再開します。" -ForegroundColor Yellow
    }
}

Step "結果"
$st = [IO.File]::ReadAllText((Join-Path $home_ "state.json")) | ConvertFrom-Json
$st.psts.PSObject.Properties.Value | ForEach-Object {
    [pscustomobject]@{ 状態 = $_.status; 段階 = $_.checkpoint; エラー = $_.last_error.code; PST = $_.source_path }
} | Format-Table -AutoSize -Wrap
if ($st.contacts) { Write-Host "連絡先: $($st.contacts.count) 件（$($st.contacts.backup_at)）  エラー: $($st.contacts.last_error.code)" }
Write-Host "`n状態が COMPLETED なら完了。PENDING / RETRY_PENDING ならもう一度実行、FAILED ならエラーを確認。"
