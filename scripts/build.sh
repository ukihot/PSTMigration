#!/usr/bin/env bash
# PSTMigration.exe をビルドし、GPO 配布一式を dist/PSTMigration/ に揃える（Git Bash 用）。
#
#   scripts/build.sh
#
# 必要: uv、ビルドPC上の 7-Zip（7z.exe / 7z.dll を EXE に同梱するため。配布先PCには不要）
#   7-Zip の場所は自動検出。見つからなければ SEVENZIP_DIR=/c/path/to/7-Zip を指定。
# 出力: dist/PSTMigration/{PSTMigration.exe, config.json, PSTMigration-manual.cmd/.ps1, 7-Zip-License.txt, SHA256SUMS.txt}
#   → config.json に NAS のパスワードを入れ、中身を \\<ドメイン名>\NETLOGON\PSTMigration\ に置く（設計書 §10.1）。
set -euo pipefail
cd "$(dirname "$0")/.."

if [[ -z "${SEVENZIP_DIR:-}" ]]; then
  for d in "/c/Program Files/7-Zip" "$HOME/scoop/apps/7zip/current"; do
    [[ -f "$d/7z.exe" && -f "$d/7z.dll" ]] && SEVENZIP_DIR="$d" && break
  done
fi
[[ -n "${SEVENZIP_DIR:-}" ]] || { echo "7-Zip が見つかりません。SEVENZIP_DIR を指定してください" >&2; exit 1; }
sz="$(cygpath -w "$SEVENZIP_DIR")"  # PyInstaller には Windows 形式のパスで渡す

uv sync
uv run ruff check .
uv run ruff format --check .
uv run pytest -q

uv run pyinstaller --noconfirm --clean --onefile --noconsole \
  --name PSTMigration \
  --paths src \
  --hidden-import win32timezone \
  --hidden-import win32comext.mapi.mapi \
  --add-binary "$sz\\7z.exe;." \
  --add-binary "$sz\\7z.dll;." \
  --distpath build/dist --workpath build/work --specpath build \
  src/pstmigration/__main__.py

# 32ビット版 Outlook の PC 用（64ビットの EXE からは MAPI を使えない。設計書 §18.6）。
# PSTMigration.exe が Outlook のビット数を見て、同じフォルダーのこちらに任せる。
uv python install cpython-3.12-windows-x86-none >/dev/null
py32="$(uv python find cpython-3.12-windows-x86-none)"
uv venv -q --allow-existing --python "$py32" build/venv32
uv export --frozen --no-hashes --no-emit-project > build/requirements32.txt
VIRTUAL_ENV=build/venv32 uv pip install -q -r build/requirements32.txt
build/venv32/Scripts/python.exe -c "import struct; assert struct.calcsize('P') == 4"
build/venv32/Scripts/pyinstaller --noconfirm --clean --onefile --noconsole \
  --name PSTMigration32 \
  --paths src \
  --hidden-import win32timezone \
  --hidden-import win32comext.mapi.mapi \
  --add-binary "$sz\\7z.exe;." \
  --add-binary "$sz\\7z.dll;." \
  --distpath build/dist --workpath build/work32 --specpath build \
  src/pstmigration/__main__.py

out=dist/PSTMigration
rm -rf "$out"
mkdir -p "$out"
cp build/dist/PSTMigration.exe build/dist/PSTMigration32.exe config.json \
  scripts/PSTMigration-manual.cmd scripts/PSTMigration-manual.ps1 "$out/"
[[ -f "$SEVENZIP_DIR/License.txt" ]] && cp "$SEVENZIP_DIR/License.txt" "$out/7-Zip-License.txt"
(cd "$out" && sha256sum PSTMigration.exe PSTMigration32.exe config.json > SHA256SUMS.txt)

echo "OK: $out"
cat "$out/SHA256SUMS.txt"
