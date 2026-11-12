# PSTMigration

社員PCに残っている **POP時代のPST** を、自動で回収し、NAS（`\\192.168.8.120\PSTBackup`）に集めるツール。

PSTの中のメールは、config の `cutoff`（今の配布値は **2016-02-21**）を境に2つに分けて置く。

| | 中身 | NASに置く形 | その後（手作業） |
|---|---|---|---|
| 旧メール | `cutoff` より前 | 7z に圧縮 | SharePoint に保管 |
| 新メール | `cutoff` 以降 | PST のまま | Purview で Exchange Online に取り込む |

元のPSTのメールは **変更・削除しない**。

---

## どんな動きをするか

社員は何もしなくてよい。画面も出ない。GPO で配られ、次の流れで勝手に進む。PSTの大きさと端末の空き容量で、2つの方式を自動で使い分ける。

### ふつうのPST（ログオフ時のコピー上限 `logoff_snapshot_budget_gb` 以内で、空きも足りる）

```
 ① サインイン  ──30分後──▶  PST を見つけて記録する                      （数秒）
 ② シャットダウン / サインアウト ──▶  PST をコピーする                      （2GBで4秒）
 ③ 次のサインイン ──30分後──▶  分割 → 7z圧縮 → NASへ転送 → 完了           （2GBで数十分）
```

- **なぜコピーがシャットダウン時なのか**：Outlook が起動している間は PST がロックされていて、ファイルとして読めない。Windows がシャットダウン時に Outlook を閉じたあとなら読める。社員が Outlook を開いたままシャットダウンしても問題ない。
- ローカルの空きは「PSTのサイズ × 2.5 ＋ 5GB」必要（コピーと分割後のPSTを同時に置くため）。

### 大きいPST（上限を超える、または空きが足りない）：ライブ分割

```
 ① サインイン ──30分後──▶  元PSTを読み取り専用で開き、約5GBずつの「束」にコピー
                          → 1束ずつ 照合 → （旧メールは7z）→ NASへ転送 → ローカルから削除 → 次の束
```

- シャットダウン時のコピーは使わない。サインイン中（Outlook 使用中でもよい）に裏で進む。
- ローカルに置くのは常に作業中の1束だけなので、**PSTが47GBでも必要な空きは十数GB**（束の目標5GB × 2.5 ＋ 5GB）。空きが少なければ束を小さくする。
- NAS には `new_mail_from_…_001.pst, _002.pst …` / `old_mail_before_…_001.7z, _002.7z …` のように束ごとに置く。
- 処理中に元PSTのメールが消された・移されたら `FAILED`（取りこぼしを防ぐため）。既読・フラグの変更は問題ない。調査後に届いたメールは対象外として台帳に残す（`recollect_candidates.csv`）。
- 使える条件：Outlook と同じビット数（64bit）の MAPI。使えない端末では保留（`LIVE_SPLIT_UNAVAILABLE`）。
- 詳しい仕様は設計書 §16。

### 共通

- **処理中も Outlook は普通に使える**。処理は低い優先度で、裏で動く。
- **1回やればおしまい**。完了したPSTは、以後のサインイン・シャットダウンでは何もしない（数秒で終わる）。
- 途中で電源が切れても、次回のサインインで続きから再開する。
- **連絡先は別扱い**：①のたびに、Outlook に見えている連絡先フォルダー（PST・IMAP の「このコンピューターのみ」。Exchange は除く）を `contacts_{PC名}.pst` にまとめ、NAS 上のものを置き換える。日付では分けない、常に最新のバックアップ。GPO を止める（[4-8](#4-8-回収が終わったらgpo-は残したまま止める)）と止まる。
  出力PSTは一時MAPIプロファイルで作成するため、利用中の Outlook に接続されず、Outlook を閉じずにバックアップできる。コピー件数の照合とPSTの解放を確認してから NAS 上を置き換え、失敗した場合は前回のバックアップを残す。

### 待ち時間の目安（パイロット実測）

| 場面 | 社員から見て |
|---|---|
| シャットダウン時のコピー | 2GBで **4秒** 延びるだけ（SSD）。初回だけ |
| サインイン後 | 何も変わらない（30分後から裏で動く） |

---

## 進み具合の見方

### NAS で見る（いちばん簡単）

`\\192.168.8.120\PSTBackup\users\{メールアドレス}\{PC名}\`

| 見えるもの | 状態 |
|---|---|
| フォルダー自体がない | まだ一度もツールが動いていない（①の前） |
| `status.json` だけある | PSTは見つけた。コピーや転送はこれから |
| `contacts_{PC名}.pst` | 連絡先のバックアップ（サインインのたびに置き換え）。件数・日時・エラーは `status.json` の `contacts` |
| `{PST名}_..._{日時}\` フォルダーに 7z・PST が増えていく | ライブ分割の処理中（束ごとに届く） |
| `{PST名}_..._{日時}\` フォルダーに `report.json` がある | **完了**（旧メールが0件の人は 7z が無い。新メールが0件なら PST が無い。件数は `report.json` の `counts`） |

> NAS の `status.json` は、サインイン時の処理が **全部終わったときに** 更新される。処理の途中（数十分〜数時間）は古いままなので、「変わっていない＝失敗」ではない。途中経過は下の「そのPCで見る」で。

`status.json` の `status` の意味：

| status | 意味 |
|---|---|
| `PENDING` | 順番待ち（②のコピー待ち、または次のサインイン待ち） |
| `COMPLETED` | 完了 |
| `RETRY_PENDING` | 失敗したが、次回自動でやり直す。放っておいてよい |
| `FAILED` | 自動ではやり直さない。[エラーコード](#エラーコード) を見て対応 |

全員分まとめて見るときは [5章](#5-運用) の `--aggregate`（Excel で開ける一覧ができる）。

### そのPCで見る（処理中の様子）

```powershell
Get-Content $env:LOCALAPPDATA\PSTMigration\state.json -Encoding utf8 | Select-String '"status"|split_mode|snapshot_at|"code"'
```

| 見えるもの | 意味 |
|---|---|
| `split_mode` が `snapshot` / `live` | ふつうの方式 / ライブ分割 |
| `snapshot_at` が `null`（ふつうの方式） | ②のコピーがまだ |
| `status` が `SCANNING` `SPLITTING` `VERIFYING` `COMPRESSING` `TRANSFERRING` `FINALIZING` | 処理中（`FINALIZING` は完了レポート保存中） |
| `chunks` の各 `status`（ライブ分割） | `VERIFIED`（照合済み）→ `BUILT`（転送待ち）→ `UPLOADED`（NAS に届いた） |
| `status` が `COMPLETED` | 完了 |

詳しいログは `%LOCALAPPDATA%\PSTMigration\logs\` の一番新しいファイル。

---

## 導入状況（2026-10-22 時点）

| | 項目 | 状態 |
|---|---|---|
| ✅ | NAS の共有・アカウント | 済み |
| ✅ | 配布物を NETLOGON に置く | 済み（`\\tsuchiyoshi.private\NETLOGON\PSTMigration\`） |
| ✅ | GPO `PSTMigration` | 適用中（68台で検出済み） |
| ✅ | ふつうの方式 | 稼働中。TYCO248（20GB、手動実行）で完了を確認 |
| ⏳ | ライブ分割 | 1.2.0 で `live_split: auto` のまま全社配布（2026-10-22）。社員PCでの稼働を確認中（[4-9](#4-9-ライブ分割の様子を見る止める)） |

10GB を超えるPSTが23個ある（2026-10-22 の集計）。1.2.0 が届いたPCから、ライブ分割で処理される。

ここから下は、作り直し・設定変更・トラブル時のための手順書。

---

## 1. ビルド

ビルドPCに必要：[uv](https://docs.astral.sh/uv/)、7-Zip（EXEに内蔵するため。**配布先PCには不要**）

```bash
scripts/build.sh          # Git Bash
scripts\build.bat         # cmd
```

lint（Ruff）→ テスト → EXE化 の順に走り、`dist/PSTMigration/` に配布一式ができる。

| ファイル | 内容 |
|---|---|
| `PSTMigration.exe` | 本体（7-Zip CLI 内蔵、コンソールなし） |
| `config.json` | 設定（[設定](#6-設定-configjson)） |
| `PSTMigration-manual.cmd` / `.ps1` | 手動実行用（[3章](#3-手動で動かすgpo-を待たずに試すやり直す)） |
| `7-Zip-License.txt` | 7-Zip のライセンス（LGPL。同梱が必要） |
| `SHA256SUMS.txt` | 配布物のハッシュ（改ざん確認用。リリースごとに控えておく） |

7-Zip が自動検出できない場合：`SEVENZIP_DIR=/c/path/to/7-Zip scripts/build.sh`

**NAS のパスワードは `dist\PSTMigration\config.json` に書く**。プロジェクト直下の `config.json`（git 管理）には書かない。
再ビルドすると `dist\` は直下の `config.json` で上書きされるので、**ビルドのたびに入れ直してから** NETLOGON に置く。直下の `config.json` は配布値の雛形（`cutoff` は `2016-02-21T00:00:00+09:00`、`live_split` は `auto`）。

## 2. NAS と配布場所の準備（初回のみ）

NAS（`192.168.8.120`）は AD に参加していないので、**ツール専用のローカルアカウントを1つ** 作り、全PCがそのアカウントで書き込む。ユーザーごとのアクセス権は設定しない。

### NAS

1. NAS の管理画面で共有フォルダー `PSTBackup` を作る。
2. ローカルユーザー（例 `pstmigration`）を作り、`PSTBackup` に **読み書き** 権限を付ける。
3. 配布する `config.json` の `server_user` / `server_password` にそのアカウントを書く。

```json
  "server_root": "\\\\192.168.8.120\\PSTBackup",
  "server_user": "pstmigration",
  "server_password": "（NASで設定したパスワード）",
```

ツールは起動時にこのアカウントで共有へ接続する（ドライブ文字は割り当てない。ユーザーからは見えない）。フォルダー `users\` `_slots\` `_admin\` は必要に応じて自動で作られる。

> **割り切り**：パスワードは `config.json` に平文で入り、各PCの `%LOCALAPPDATA%\PSTMigration\bin\` と NETLOGON に置かれる。読み取れる人は NAS 上の全員分のPSTにアクセスできる。
> 回収期間中だけの運用とし、**回収が終わったら NAS 側でアカウントを無効化するかパスワードを変える**。

### NAS 上のフォルダー構成

何百人分になるので、**メールアドレスのフォルダー** の中に結果が入るようにしている。

```
\\192.168.8.120\PSTBackup\
    users\
        taro.yamada@example.co.jp\
            PC-0123\
                status.json                                   このPCの処理状況
                contacts_PC-0123.pst                          連絡先（サインインのたびに置き換え）
                Outlook_3f9a1c0d2b7e_20261013190540\          元PST 1個分（ふつうの方式）
                    old_mail_before_20160221_PC-0123.7z
                    new_mail_from_20160221_PC-0123.pst
                    report.json
                Archive_5be0d14a7c21_20261022101500\          元PST 1個分（ライブ分割）
                    new_mail_from_20160221_PC-0123_001.pst
                    new_mail_from_20160221_PC-0123_002.pst
                    old_mail_before_20160221_PC-0123_001.7z
                    report.json
            PC-0456\                                          同じ人の別PC
    _slots\
    _admin\
```

### メールアドレスの決め方

各PCで、ログオン時に次の順で決める。決め方は `ledger.csv` の `mailbox_source` 列で分かる。

| 順 | 取得元 | `mailbox_source` |
|---|---|---|
| 1 | Outlook に設定されているメールアカウント（1つだけの場合） | `OUTLOOK` |
| 2 | AD のユーザーの `mail` 属性（Outlook のアカウントが複数・なしの場合） | `AD` |
| 3 | 取れなければ Windows ユーザー名（フォルダー名がユーザー名になる） | `USERNAME` |

- 一度アドレスが取れたPCは、次回取れなくても前回のアドレスを使い続ける。
- 当社の AD は `mail` 属性が空のアカウントがある（少なくとも検証OU）。Outlook に共有アドレス（info@ など）も設定している人は `USERNAME` になりやすい。その人は Purview の CSV で手で記入する。

### 配布場所（NETLOGON）

各PCはドメインアカウントでログオンしており、ツールが接続する前は NAS を読めない。そのため配布物は NAS ではなく、全ドメインユーザーが読める NETLOGON に置く。

```
\\<ドメイン名>\NETLOGON\PSTMigration\    ← dist/PSTMigration/ の中身をコピー
```

### 同時転送枠

VPN・サーバの負荷を抑えるため、同時に転送するPC台数をスロットファイルの数で制限する。5台にするなら：

```powershell
New-Item -ItemType Directory \\192.168.8.120\PSTBackup\_slots -Force
1..5 | % { New-Item ("\\192.168.8.120\PSTBackup\_slots\slot{0:D2}.lock" -f $_) -ItemType File }
```

台数を変えるにはファイルを増減するだけでよい。`_slots` が無い・空なら制限なし。ライブ分割は束を1つ送るたびに枠を取り直す。

### 繋がらないとき

- ドメインのGPOで「NTLM の制限：リモートサーバーへの送信 NTLM トラフィック」を拒否にしていると NAS に接続できない。
- Windows 11 24H2 以降はクライアント側で SMB 署名が必須。古い NAS で署名に対応していないと接続できない（NAS の設定で SMB 署名を有効にする）。
- ユーザーが同じ NAS に別のアカウントで手動接続していると、その接続のまま使う（ログに警告が出る）。

## 3. 手動で動かす（GPO を待たずに試す・やり直す）

**ふつうはこれだけ**：Outlook を閉じて、`%LOCALAPPDATA%\PSTMigration\bin\PSTMigration-manual.cmd` をダブルクリックする（GPO で EXE と一緒に配られている）。①検出 → ②コピー → ③分割・7z・転送 を続けて行い、最後に結果を表示する。

- ログオフ時のコピー上限（`logoff_snapshot_budget_gb`）を無視するので、上限を超える大きいPSTも、空きが「PSTのサイズ × 2.5 ＋ 5GB」あればふつうの方式で処理できる。
- 空きが足りなければ、`live_split` の設定に従ってライブ分割（または保留）になる。
- ③の間は Outlook を開いて使ってよいが、PC はスリープ・シャットダウンしない。
- Ctrl+C で止めると、EXE・7z と、ツールが裏で起動した Outlook を後片付けする。次回は続きから。

各段階を個別に動かすときは次のとおり（上の①②③と対応）。

```powershell
$exe = "$env:LOCALAPPDATA\PSTMigration\bin\PSTMigration.exe"

# ① PSTを見つける（Outlook 起動中でOK）。ライブ分割のPSTはここで処理まで進む
Start-Process $exe -Wait

# ② PSTをコピーする（シャットダウンの代わり。Outlook を閉じてから）
Start-Process $exe -ArgumentList "--snapshot" -Wait

# ③ 分割 → 7z → NAS転送（Outlook を起動してから）
Start-Process $exe -Wait
```

- EXE は画面を持たないので、`& $exe` と書くと終了を待たずに戻ってくる。**必ず `Start-Process -Wait`** で実行する。
- 分割せずに件数だけ見たいときは、③の代わりに `--scan-only`。結果は `%LOCALAPPDATA%\PSTMigration\reports\{key}.scan.json`。`--scan-only` ではライブ分割も動かない。
- 別の設定で動かすときは `-ArgumentList "--config","<設定ファイル>"`。`bin\config.json` は GPO で上書きされるので、別の場所に置く。

### パイロットで確認すること

- [ ] NAS の `users\{メールアドレス}\{PC名}\` ができる（フォルダー名がユーザー名なら、メールアドレスが取れていない → [メールアドレスの決め方](#メールアドレスの決め方)）
- [ ] 完了後、その中の `{PST名}_..._{日時}\` に 7z・PST・`report.json` がある（0件の側のファイルは作られない）
- [ ] `report.json` の `counts` で `mail = old + new + unmigrated`、かつ `unmigrated = 0`
- [ ] 7z を展開した旧PSTと、新PSTを Outlook で開き、フォルダー・件数・添付・日時がそろっている
- [ ] 元のPSTの件数が変わっていない（削除済みアイテムにゴミが増えていない）
- [ ] Outlook のセキュリティ警告が出ない、ESET が EXE を検知しない（設計書 §14.1）
- [ ] ライブ分割：束（`_001` …）が NAS に順に届き、ローカルの `work\` に束がたまらない。`report.json` の `chunks` がそろい、`split_mode = live`

## 4. GPO設定

ドメインコントローラー（Windows Server 2012、`192.168.100.1`）の「グループ ポリシーの管理」で作業する。ドメインは `tsuchiyoshi.private`。

方針：**新しい GPO を1つ作り、まず `検証` OU にだけリンクして試す**。問題なければ `全社基本` にもリンクする。既存の GPO には手を付けない（新しい GPO がリストに1行増えるだけ。他の GPO と設定が重ならないので、リンクの順番も気にしなくてよい）。

### 4-1. 配布物を NETLOGON に置く

DC 上で（またはエクスプローラーから）フォルダーを作り、`dist\PSTMigration\` の中身（パスワード入りの `config.json` を含む）をコピーする。

```
\\tsuchiyoshi.private\NETLOGON\PSTMigration\
    PSTMigration.exe
    config.json
    PSTMigration-manual.cmd
    PSTMigration-manual.ps1
    7-Zip-License.txt
    SHA256SUMS.txt
```

DC 上の実体は通常 `C:\Windows\SYSVOL\sysvol\tsuchiyoshi.private\scripts\PSTMigration\`。

### 4-2. GPO を作ってリンクする

1. 「グループ ポリシーの管理」で `tsuchiyoshi.private` → `全社基本` → **`検証`** を右クリック
2. 「**このドメインに GPO を作成し、このコンテナーにリンクする**」
3. 名前：`PSTMigration` → OK
4. `検証` の下にできた `PSTMigration` を右クリック → 「**編集**」（グループ ポリシー管理エディターが開く）

以下 4-3〜4-5 は、すべてエディターの **「ユーザーの構成」** 側で設定する（「コンピューターの構成」ではない）。

### 4-3. EXE の配布（フォルダー → ファイル）

ファイル配布は `*` で複数ファイルをコピーするとき **コピー先フォルダーを作らない**（無いと `0x80070003 指定されたパスが見つかりません` で失敗する）。先にフォルダーを作る項目を足す。

**(1) ユーザーの構成 → 基本設定 → Windows の設定 → フォルダー** を右クリック → 新規作成 → フォルダー

| タブ | 項目 | 値 |
|---|---|---|
| 全般 | アクション | 更新 |
| 全般 | パス | `%LOCALAPPDATA%\PSTMigration\bin` |
| 共通 | ログオンしたユーザーのセキュリティ コンテキストで実行する | ✅ チェックする |

**(2) ユーザーの構成 → 基本設定 → Windows の設定 → ファイル** を右クリック → 新規作成 → ファイル（フォルダーの項目はファイルより先に処理される）

| タブ | 項目 | 値 |
|---|---|---|
| 全般 | 操作 | **置換**（「更新」だと既存ファイルを上書きしない） |
| 全般 | ソース ファイル | `\\tsuchiyoshi.private\NETLOGON\PSTMigration\*` |
| 全般 | ターゲット フォルダー | `%LOCALAPPDATA%\PSTMigration\bin` |
| 共通 | ログオンしたユーザーのセキュリティ コンテキストで実行する | ✅ チェックする |

パスはすべて **`%LOCALAPPDATA%`** で書く（各ユーザーの `C:\Users\<ユーザー>\AppData\Local` になる）。基本設定の変数 `%LocalAppDataDir%` は、この環境ではフォルダー作成が `0x80070057 パラメーターが間違っています` で失敗したので使わない。

ソース ファイルの **先頭は `\\`、末尾は `\*`**。`\*` が無いと「`PSTMigration` という名前のファイル」とみなされ、エラーも出ずに何もコピーされない。正しく入ると下の欄の表示が「ターゲット フォルダー」に変わる。

うまく配られないときは、PCのイベント ビューアー → Windows ログ → Application で、ソース「Group Policy Files」（ID 4098 / 4117）にコピー元・コピー先と失敗理由が出る。

### 4-4. ログオン時の自動実行（スケジュールされたタスク）

**ユーザーの構成 → 基本設定 → コントロール パネルの設定 → タスク スケジューラ** を右クリック → 新規作成 → **スケジュールされたタスク (Windows 7 以降)**

| タブ | 項目 | 値 |
|---|---|---|
| 全般 | 操作 | **作成**（「更新」「置換」にしない。下の注意） |
| 全般 | 名前 | `PSTMigration` |
| 全般 | タスクの実行時に使うユーザー アカウント | `%LogonDomain%\%LogonUser%`（「ユーザーまたはグループの変更」で入力） |
| 全般 | セキュリティ オプション | 「ユーザーがログオンしているときのみ実行する」 |
| トリガー | 新規 → タスクの開始 | 「ログオン時」、「特定のユーザー」で `%LogonDomain%\%LogonUser%` |
| トリガー | 遅延時間を指定する | ✅ 30 分（選択肢に15分が無いため） |
| 操作 | 新規 → 操作 | プログラムの開始 |
| 操作 | プログラム/スクリプト | `%LOCALAPPDATA%\PSTMigration\bin\PSTMigration.exe` |
| 条件 | コンピューターを AC 電源で使用している場合のみタスクを開始する | ☐ 外す |
| 設定 | タスクを停止するまでの時間 | ☐ 外す |
| 設定 | タスクが既に実行中の場合に適用される規則 | 新しいインスタンスを開始しない |

**操作は必ず「作成」**。「更新」だと、グループポリシーが処理されるたび（サインイン直後・`gpupdate`・約90分ごとの自動更新）にタスクが登録し直される。登録し直されると、待機中の「ログオン30分後」の実行予定が消えて、**タスクが一度も動かない**（パイロットで発生：10:02 サインイン → 10:03 にタスクが再登録 → 10:32 に動かず）。
「作成」はタスクが無いときだけ作り、以後は触らない。

タスクの設定（遅延時間など）を後から変えたいときは、PC側のタスクが更新されないので、一度「置換」にして `gpupdate` で配り直し、すぐ「作成」に戻す。

### 4-5. ログオフ時のスナップショット（スクリプト）

**ユーザーの構成 → ポリシー → Windows の設定 → スクリプト (ログオン/ログオフ) → ログオフ** をダブルクリック → 追加

| 項目 | 値 |
|---|---|
| スクリプト名 | `cmd.exe` |
| スクリプトのパラメーター | `/c cd /d "%LOCALAPPDATA%\PSTMigration\bin" && if exist PSTMigration.exe start "" /wait PSTMigration.exe --snapshot` |

**手入力せずコピー＆ペーストする**（`\` が抜けやすい。抜けていても止まりはしないが、スナップショットが取られない）。

EXE を直接指定せず `cmd.exe` を挟むのは、(1) `%LOCALAPPDATA%` を確実に展開させるため、(2) `start /wait` で EXE の終了まで待たせるため（ウィンドウのない EXE は、待たずに終わったとみなされてログオフが進み、コピー途中で止められてしまう）。ログオフスクリプトは既定で非表示実行なので、黒い画面は出ない。

**`if exist` は必須**。EXE がまだ配られていないPCで `start` すると「見つかりません」のエラーダイアログが **見えない画面に** 出て、誰も閉じられないままサインアウトが最大10分止まる（実際に起きた）。

### 4-6. 効いているか確認する（検証 OU のPCで）

```powershell
gpupdate /force          # その後、一度サインアウトしてサインイン
gpresult /r /scope user  # 「適用されたグループ ポリシー オブジェクト」に PSTMigration があること
Get-ChildItem "$env:LOCALAPPDATA\PSTMigration\bin"      # EXE と config.json が配られている
Get-ScheduledTask -TaskName PSTMigration                 # タスクがある
```

その後の流れ：ログオン30分後に検出 → サインアウト時にスナップショット → 次回ログオン30分後に本処理。結果は NAS の `users\{メールアドレス}\{PC}\` と `ledger.csv` で確認する。

### 4-7. 全社に広げる

`全社基本` を右クリック → 「**既存の GPO をリンクする**」 → `PSTMigration` を選ぶ。`検証` OU は `全社基本` の下にあるので、`検証` 側のリンクは削除してよい（残しても二重には動かない）。

**社員への周知**：初回だけシャットダウン・サインアウトが少し長くなることがある、と伝えておく（SSD なら数秒〜数十秒）。

### 4-8. 回収が終わったら（GPO は残したまま止める）

1. タスクの項目（4-4）を開き、操作を **「削除」** にする → 各PCのタスクが次回のポリシー処理で消える
2. ログオフスクリプト（4-5）の行を **削除** する
3. NAS のアカウントを無効化する（各PCの `config.json` にパスワードが残っているため）

フォルダーとファイルの項目（4-3）は残してもよい（EXE がPCに置かれるだけで、動かない）。

### 4-9. ライブ分割の様子を見る・止める

`live_split` は `auto`（ふつうの方式で処理できないPSTを、自動でライブ分割する）で配っている。

様子を見る：

- `--aggregate` の `ledger.csv` で `split_mode = live` の行の `status`・`chunks`・`last_error_code`。
- NAS の収集単位フォルダーに束（`_001` …）が順に届いているか。
- そのPCの `%LOCALAPPDATA%\PSTMigration\work\{key}\` に束がたまっていないか（作業中の1束だけのはず）。
- [パイロットで確認すること](#パイロットで確認すること)のライブ分割の項目。

止める：NETLOGON の `config.json` を `"live_split": "off"` にして配る。次のポリシー処理で各PCに届き、まだ始まっていない大きいPSTは `LIVE_SPLIT_UNAVAILABLE` で保留になる。処理中・処理済みのPST（`split_mode = live`）はそのまま続きを進めるので、完全に止めるには GPO のタスクを止める（[4-8](#4-8-回収が終わったらgpo-は残したまま止める)）。

### バージョンアップ

1.2.4（2026-11-12）：**32ビット版 Outlook の PC に対応**（`PSTMigration32.exe` を同梱。`PSTMigration.exe` が Outlook のビット数を見て任せる。従来は 64ビットの EXE から MAPI を使えず、連絡先バックアップ `OUTLOOK_COM` と `LIVE_SPLIT_UNAVAILABLE` になっていた）。NAS の `_admin\reset.json` で FAILED をPCごとに戻せるようにした（[FAILED を再試行させる](#failed-を再試行させる)）。診断の追加：PCの環境（ツール/Outlook のビット数・Office の版・空き容量・作業フォルダー容量・NAS 接続可否・MAPI の失敗理由）を status.json と `_admin\pcs.csv` に、`PST_LOCKED` に掴んでいるプロセス名、転送失敗に段階・転送済みバイト・経過秒、直近20回分の実行ログを NAS の `users\<所有者>\<PC>\logs\` に置く。ライブ分割の転送で、途切れた `.partial` の末尾64MBを書き直してから再開する。転送後の不一致でローカルの束を読み直せないとき（別プロセスが使用中）を `ACCESS_DENIED` ではなく `HASH_MISMATCH`（再試行）にする。計画後に元PSTのフォルダーが見つからないときは `SOURCE_CHANGED_DURING_COPY` にし、同じ階層の同名フォルダーもたどる。`CHUNK_MISMATCH` に指紋のどの項目が違うか（件名は値なし）を出す。完了済み・台帳に無いPSTの作業フォルダーを消す。設計書 §18。

1.2.3（2026-11-04）：障害解析と検証済み成果物からの再開を改善。`UNEXPECTED` は即 `FAILED` にして、例外クラス・内部例外付きスタックトレース・発生工程・発生時バージョンを `last_error` に保存する。`--aggregate` で詳細を `_admin/errors.json` に集約する。`ARCHIVED` の再開時はサイズとSHA-256を確認し、不正な成果物は消さずに停止する。`TRANSFERRED` からの完了登録再開もNASを照合する。ライブ分割は調査前にも最低限の空き容量を確認する。詳細と復旧手順は設計書 §17。

`ledger.csv` の末尾に、世代・出典 `status.json`・同一PC/同一PSTパスの重複候補・再収集要否・例外情報を追加。重複行は履歴か別対象かを判断できないため削除しない。`source_gb` は実サイズ0と未取得（空欄）を区別する。

NAS上の詳細も確認したところ、CSVの39件に対し、現在も残る `ARCHIVED / UNEXPECTED` 37件は同じSMBゲストアクセス拒否だった。代表端末のログで、v1.2.1の `transfer.slot` 内のNAS一覧取得時にWinError 1272が出たことを確認。この経路を `NETWORK / RETRY_PENDING` に修正したが、NASの認証設定自体は別途確認が必要。ライブ分割4件の文字コード例外については、フォルダー作成のUnicode指定とJSON保存も修正した。

v1.2.2以前のCSVの `tool_version` は状態を最後に出力したツールの版であり、例外が発生した版を証明しない。既存の `FAILED` は自動で解除しない。

1.2.2（2026-10-30）：`--aggregate` が、照合済みで圧縮前の束（ライブ分割）を含む status.json で異常終了し、`_admin\` に何も出ない不具合を修正。

1.2.1（2026-10-22）：連絡先バックアップを一時MAPIプロファイルで作る方式へ変更。ライブ分割では Windows Search 等による出力PSTのロック解放を待ってから照合・ハッシュ計算・圧縮・転送を行い、待機上限を超えた場合も作成済みの束から再開する。

`scripts/build.sh` → `dist/PSTMigration/` の中身で NETLOGON の `PSTMigration\` を上書きするだけ（`config.json` のパスワード・`cutoff`・`live_split` を入れ直すのを忘れない）。ファイルの項目が「置換」なので、次回のポリシー処理（ログオン時・約90分ごと・`gpupdate`）で各PCへ配られる。ツールが実行中のPCでは EXE が使用中で差し替えに失敗するが、次のポリシー処理で入る。処理途中のPSTは state.json の続きから再開する。

## 5. 運用

### 進み具合を見る

管理者PCで、先にエクスプローラーから `\\192.168.8.120\PSTBackup` を NAS のアカウントで開いておき：

```powershell
Start-Process "\\tsuchiyoshi.private\NETLOGON\PSTMigration\PSTMigration.exe" -ArgumentList "--aggregate","\\192.168.8.120\PSTBackup" -Wait
```

`_admin\` に4つのCSVが出る（Excelでそのまま開ける）。

| ファイル | 中身 |
|---|---|
| `ledger.csv` | 移行台帳。1行＝1PST。status・件数・エラーコード・方式・束の数 |
| `purview_mapping_draft.csv` | Purview インポート用マッピングCSVのたたき台（新メールPST1つにつき1行。ライブ分割は束ごと） |
| `archive_list.csv` | 旧メール7zの一覧（サイズ・SHA-256。ライブ分割は束ごと） |
| `recollect_candidates.csv` | 処理後に元PSTへメールが増えた収集単位（移行直前の再回収の候補） |
| `errors.json` | エラーの詳細（例外・スタックトレース・転送の段階など） |
| `pcs.csv` | 1行＝1PC（ユーザー）。ツール/Outlook のビット数・空き容量・連絡先バックアップと実行時のエラー（1.2.4以降。各PCのログは `users\<所有者>\<PC>\logs\`） |

`ledger.csv` で見るところ：

- `status = PENDING` が何日も続く → そのPCが **ログオフされていない**（スリープ運用）。サインアウトを1回してもらう。
- `status = FAILED` → `last_error_code` を見て[エラー対応](#エラーコード)。
- `warnings` に何か出ている → 下の表を参照。
- `source_gb`（元PSTのサイズ。ログオンのたびに更新）が `logoff_snapshot_budget_gb` を超えている → ログオフ時にはコピーされない。ライブ分割で処理されるか、`live_split = off` なら保留（`LIVE_SPLIT_UNAVAILABLE`）。
- `split_mode`（`snapshot` / `live`）、`chunks`（束の数）、`source_added`（処理後に増えたメールの件数）。
- `new_pst_size` / `archive_size` は、ライブ分割なら束の合計。

### ステータス

| status | 意味 |
|---|---|
| `PENDING` | 検出済み。次のログオフ（スナップショット）または次のログオン待ち |
| `SCANNING` / `SPLITTING` / `VERIFYING` / `COMPRESSING` / `TRANSFERRING` | 処理中（または処理中に落ちた。次回続きから）。ライブ分割は束の作成〜転送の間ずっと `SPLITTING` |
| `RETRY_PENDING` | 失敗。次回自動で再試行する |
| `COMPLETED` | 完了。サーバに揃っている |
| `FAILED` | 自動では再試行しない。管理者対応が必要 |
| `SKIPPED` | 対象外（OST、旧スクリプトの出力PSTなど。`skip_reason` 参照） |

### エラーコード

一時的なもの（ネットワーク・容量など）は回数制限なしで次回再試行する。それ以外は `max_attempts`（既定3回）失敗すると `FAILED` になる。

| コード | 意味 | 対応 |
|---|---|---|
| `OUTLOOK_UNAVAILABLE` | Classic Outlook に接続できない | New Outlook のみのPCでないか確認 |
| `SOURCE_LOCKED` | ログオフ時にPSTが使用中 | 通常は次回で解消。続くなら Outlook がログオフ時に残っていないか |
| `SOURCE_MISSING` | 元PSTが見つからない | ユーザーが移動・削除していないか |
| `DISK_SPACE` | ローカルの空き不足（ふつうの方式は元PSTの約2.5倍＋5GB。ライブ分割は束を作る空きが無い） | 空きを作るか、`work_dir` を別ドライブへ |
| `NOT_PST` / `PST_OPEN_FAILED` | PSTとして開けない（破損など） | scanpst で修復したコピーを手作業で処理 |
| `COUNT_MISMATCH` | 分割後の件数が合わない（ふつうの方式） | 自動でスナップショットから取り直す。3回続いたら FAILED。ログ確認 |
| `CHUNK_MISMATCH` | 束の中身が計画と一致しない（ライブ分割） | 即 FAILED。ログで不一致の内容を確認し、`--reset` |
| `SOURCE_CHANGED_DURING_COPY` | ライブ分割の途中で、元PSTのメールが消された・移された | 即 FAILED。ユーザーに整理を止めてもらってから `--reset` |
| `LIVE_SPLIT_UNAVAILABLE` | ふつうの方式で処理できず、ライブ分割も使えない（`live_split = off`、または MAPI のセッションを張れない） | `live_split` を見直す。MAPI なら Office のビット数（64bit 前提）を確認 |
| `CHUNKS_INCOMPLETE` / `STATE_CORRUPT` | 束の記録が計画を覆っていない（通常は起きない） | ログ確認。`--reset` |
| `PST_LOCKED` | Outlook・Windows Search・ウイルス対策ソフト等が作業用PSTを手放さない | 解放を待って続行。待機上限を超えた場合は次回ログオン時に続きから |
| `OUTLOOK_LOST` | 処理中に Outlook が終了・クラッシュした（社員が閉じた等） | 次回ログオン時に続きから。放置でよい。同じPCで毎回続くなら Outlook のクラッシュをイベントログで確認 |
| `ARCHIVE_FAILED` / `ARCHIVE_TEST_FAILED` | 7z 圧縮・テストの失敗 | ログの7z出力を確認。AppLocker で `%TEMP%` 実行禁止なら設計書 §10.1 |
| `SLOT_BUSY` | 転送枠が空かない | スロットを増やすか放置 |
| `NETWORK` / `ACCESS_DENIED` | NAS に書けない | NAS の電源・VPN接続、`server_user` / `server_password`、NAS 側の読み書き権限、[繋がらないとき](#繋がらないとき) |
| `HASH_MISMATCH` | NAS転送後のファイル全体のSHA-256不一致 | `last_error.details` で期待値・NAS側・現在のローカル側のハッシュを比較。ローカルも記録と違う場合は即 FAILED、ローカルが一致する場合は再試行（上限あり）。メール単位の照合失敗とは別 |
| `ARTIFACT_MISSING` / `ARTIFACT_MISMATCH` / `ARTIFACT_METADATA` | 再開に必要な成果物が無い・内容が変わった・検証記録が不整合 | 即 FAILED。作業ファイルを保持して調査する。`ARCHIVED` は検証情報を復元できるまで再開不可 |
| `REMOTE_CONFLICT` | サーバに同名で中身の違うファイルがある（ふつうの方式） | 上書きはしない。どちらが正しいか確認して手で整理 |
| `UNMIGRATED_ITEMS` | 日付の取れないメールがある（転送自体は済んでいる） | report.json の `errors` で場所を確認。元PSTに残っているので個別対応 |
| `UNEXPECTED` | 想定外の例外 | 即 FAILED。`_admin/errors.json` またはローカルログのスタックトレースを確認し、原因解消後に個別に `--reset` |

| 警告（warnings） | 意味 |
|---|---|
| `SNAPSHOT_OVER_BUDGET` | PSTが `logoff_snapshot_budget_gb` より大きく、ログオフ時にコピーしない |
| `SOURCE_CHANGED` | ふつうの方式：完了後に元PSTのメールが増減した。ライブ分割：調査後にメールが増えた。どちらも `recollect_candidates.csv` に出る（下の再収集を参照） |
| `STALE_REMOTE_LEFT` | `cutoff` 変更・`--reset` 前の NAS フォルダーを消せなかった。`stale_remote_dirs` のフォルダーを手で削除 |
| `NEW_PST_OVER_LIMIT` | 新メールPSTが20GB超（ふつうの方式）。Purview 投入前に手で分割を検討 |
| `ARTIFACT_MISSING` / `SNAPSHOT_MISSING` | 分割までの作業ファイルが消えていたので最初からやり直した。`ARCHIVED` 以降の成果物欠落は上記エラーとして停止する |

### FAILED を再試行させる

原因を取り除いたあと、**そのPCで・そのユーザーとして**：

```powershell
Start-Process "$env:LOCALAPPDATA\PSTMigration\bin\PSTMigration.exe" -ArgumentList "--reset","<pst_key>" -Wait
```

`pst_key` は ledger.csv の列にある。

多数のPCをまとめて戻すとき（1.2.4以降）：管理者が `\\192.168.8.120\PSTBackup\_admin\reset.json` を置く。各PCは次の実行時に、自分のPC名の行のうち `before` より前に FAILED になったものだけを `--reset` と同じく戻す。戻した後にまた失敗したものは繰り返し戻さない（もう一度戻すなら `before` を新しくして置き直す）。ファイルが無い・壊れているときは何もしない。

```json
{"before": "2026-11-12T18:00:00+09:00", "targets": ["TYCO054N/8ab754353bb7", "TYCO124N/36668f3d5ade"]}
```

`targets` は `PC名/pst_key`（ledger.csv の `computer` と `pst_key`）。同じ pst_key が別のPCにあることがあるのでPC名も書く。

- ふつうの方式：途中まで済んでいる段階は保持され、続きから再開する。
- ライブ分割：最初（調査）から新しいフォルダーに作り直す。前のフォルダーは、新しいフォルダーへの転送が全部済んだ時点で NAS から消す。

### 完了後に元PSTが変わった（再収集）

ユーザーがIMAPのメールをPSTに移すと、完了済みでも `SOURCE_CHANGED` 警告が付く。既定では再収集しない。
移行直前に最終回収したい場合は、`config.json` を `"recollect_on_change": true` にして配布する。次回ログオン以降、変更のあったPSTが **別フォルダーに** 再収集される（既存は残る）。対象は `recollect_candidates.csv` と `ledger.csv` の `warnings` で確認する。

### ログの場所（各PC）

```
%LOCALAPPDATA%\PSTMigration\
    state.json          全PSTの状態（これが正）
    logs\run-*.jsonl    実行ログ（30日で自動削除）
    reports\{key}.json  完了したPSTの report（サーバと同じもの）
    work\{key}\         作業中ファイル（完了すると消える）。ライブ分割は plan.json（計画）と作業中の束
```

## 6. 設定 (config.json)

`PSTMigration.exe` と同じフォルダーの `config.json` を読む（`--config` で別のファイルも指定できる）。キーを間違えると何もせず終了する（ログにエラーが出る）。書かなかったキーは下の既定値になる。

| キー | 既定 | 説明 |
|---|---|---|
| `cutoff` | `2021-01-01T00:00:00+09:00` | 分割の基準日時（タイムゾーン必須）。出力ファイル名の日付にもなる。今の配布値は `2016-02-21T00:00:00+09:00`。**途中で変えると**、変更前の値で処理済み・処理中のPSTは次回ログオンで最初からやり直し、新しいフォルダーへの転送が済んだ時点で旧フォルダーを NAS から消す |
| `server_root` | `\\192.168.8.120\PSTBackup` | 転送先 |
| `server_user` / `server_password` | 空 / 空 | NAS のローカルアカウント。配布する config.json に書く。`server_user` が空なら、ログオンユーザーの資格情報のまま接続を試みる |
| `work_dir` | `%LOCALAPPDATA%\PSTMigration\work` | 作業フォルダー。Cドライブが小さいPCは変更 |
| `disk_factor` / `disk_reserve_gb` | 2.5 / 5 | ふつうの方式の必要空き容量 = 元PSTサイズ × factor + reserve。`disk_reserve_gb` はライブ分割でも Windows 用に残す空き |
| `logoff_snapshot_budget_gb` | 10 | 1回のログオフでコピーする合計上限（ログオフ待ちを抑える）。超えるPSTはライブ分割の対象 |
| `split_pause_ms` | 200 | 分割中、1回のコピー（最大500件）ごとの休み。Outlook が重いと言われたら増やす |
| `release_wait_sec` | 900 | Outlook・Windows Search 等が出力PSTを手放すまで待つ秒数。ライブ分割の照合・ハッシュ計算・圧縮・転送前にも適用 |
| `sevenzip_level` / `sevenzip_threads` | 5 / 2 | 7z の圧縮レベル・スレッド数 |
| `transfer_max_mbps` | 50 | 1台あたりの転送帯域上限（0で無制限） |
| `slot_wait_min` | 30 | 転送枠の空き待ち上限 |
| `new_pst_warn_gb` | 20 | 新メールPSTの警告サイズ（ふつうの方式） |
| `max_attempts` | 3 | 一時的でないエラーの再試行上限 |
| `recollect_on_change` | false | 完了後の変更で再収集するか |
| `log_subject` | false | 日付取得不能メールの件名を report に残すか（障害解析時のみ true に） |
| `live_split` | `auto` | ライブ分割。`auto`：ふつうの方式で処理できないPSTだけ / `always`：MAPI が使えれば全PST / `off`：使わない（[4-9](#4-9-ライブ分割の様子を見る止める)） |
| `chunk_max_gb` / `chunk_min_gb` | 5 / 1 | 束の目標サイズの上限・下限。空きが少ないと上限より小さくし、下限を割るなら `DISK_SPACE` |
| `chunk_disk_factor` | 2.5 | 束の目標サイズに対して必要な空きの比率（束PST＋7z＋余裕）。束の目標 =（空き − `disk_reserve_gb`）÷ この値 |
| `batch_target_mb` / `batch_items` | 100 / 500 | 1回のコピー（`CopyMessages`）の量の目安（メールサイズの合計・件数） |
| `plan_size_factor` | 1.15 | 束のサイズの見積もり用（メールサイズの合計 × この値）。容量の判断は実ファイルサイズで行う |

## 後続作業（SharePoint / Purview）

設計書 §13.1 の要約。全PC完了後（または区切りのいいところで）。

**新メール → Exchange Online**

1. `--aggregate` で出た `purview_mapping_draft.csv` を確認する。`Mailbox` 列はメールアドレスのフォルダー名から自動で入っている。M365 のアドレスと違う人がいれば直す。`Note = MAILBOX_UNKNOWN` の行は空欄なので記入する。
2. `Note = MULTIPLE_UNITS` の行（1人に複数のPSTがある。ライブ分割の束どうしは含まない）は、どれを入れるか決める。決めたら `Note` 列は削除する。
3. AzCopy で `users\` を `--include-pattern "new_mail_from_*.pst"` 付きでアップロードする（CSVの `FilePath` と同じ相対パスになる。`*.pst` だと連絡先のバックアップも上がる）。
4. Purview でインポートジョブを作る。`TargetRootFolder` は `/POP移行メール` のままにする（IMAP移行分と混ざらない）。ライブ分割の束は、同じメールボックスに複数のPSTとして取り込まれ、フォルダーは1つにまとまる。

**旧メール → SharePoint Online**

1. `users\` 配下の各収集フォルダーから `*.7z` と `report.json` を移す。
2. `archive_list.csv` のサイズ・SHA-256 と突き合わせる。

---

## 開発メモ

```powershell
uv sync                                  # 依存（pywin32。開発用に pyinstaller, pytest, hypothesis, ruff）
uv run pytest -q                         # Outlook 不要のテスト
$env:PSTM_INTEGRATION = "1"; uv run pytest tests/integration -q   # 実 MAPI・実 7z の統合テスト
uv run ruff check --fix . ; uv run ruff format .

# ソースから直接実行（ログは %LOCALAPPDATA%\PSTMigration\logs と画面に出る）
$env:PYTHONPATH = "src"; uv run python -m pstmigration --scan-only
```

統合テストは、一時フォルダーに新規作成したテスト用PSTだけを使い、一時 MAPI プロファイル（`PSTMigration-tmp-*` / `PSTM-IT-*`。終了時に削除）で開く。ユーザーの Outlook・メールボックスには接続も書き込みもしない。

| ファイル | 役割 |
|---|---|
| `src/pstmigration/__main__.py` | 引数・多重起動防止・ログ |
| `src/pstmigration/pipeline.py` | 処理の流れ（スナップショット、分割→圧縮→転送→完了、ライブ分割の組み立て） |
| `src/pstmigration/chunks.py` | ライブ分割の束：範囲・照合・作成〜転送・再開・完了判定（Outlook 不要） |
| `src/pstmigration/mapilive.py` | ライブ分割の Extended MAPI（読み取り専用オープン・一時プロファイル・コピー・束の読み直し） |
| `src/pstmigration/outlook.py` | Outlook COM（検出・マウント・調査・Move・件数検証） |
| `src/pstmigration/rules.py` | 日付判定・方式の選択などの純粋ロジック |
| `src/pstmigration/transfer.py` | SMB転送・転送枠・SHA-256 |
| `src/pstmigration/report.py` | report.json / status.json / 集計CSV |
| `src/pstmigration/state.py` / `config.py` | 状態ファイル・設定 |

| テスト | 内容 |
|---|---|
| `tests/test_core.py` | 日付判定・転送・集計など |
| `tests/test_chunks.py` | 束の範囲・照合（Hypothesis のプロパティテスト）・設定の検証・方式の選択 |
| `tests/test_resume.py` | 束の作成〜転送と中断からの再開（判定表の例と、Hypothesis の状態遷移テスト） |
| `tests/test_live_pipeline.py` / `test_live_report.py` / `test_mapilive.py` | ライブ分割の組み込み・台帳・MAPI 不要の部分 |
| `tests/integration/` | 実 MAPI・実 7z（`PSTM_INTEGRATION=1`）。通し処理とプロセス強制終了からの復旧 |

旧 `Split-PstByDate.ps1` は廃止（元PSTに `Copy()` で書き込む問題があったため。設計書 §1.1）。
