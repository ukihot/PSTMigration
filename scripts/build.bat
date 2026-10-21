@echo off
rem PSTMigration.exe をビルドし、GPO 配布一式を dist\PSTMigration\ に揃える（cmd 用。中身は build.sh と同じ）。
rem 必要: uv、ビルドPC上の 7-Zip（EXE に同梱するため。配布先PCには不要）。
rem 7-Zip の場所が違う場合は  set SEVENZIP_DIR=C:\path\to\7-Zip  してから実行。
setlocal
chcp 65001 >nul
cd /d "%~dp0.."

if not defined SEVENZIP_DIR if exist "%ProgramFiles%\7-Zip\7z.exe" set "SEVENZIP_DIR=%ProgramFiles%\7-Zip"
if not defined SEVENZIP_DIR if exist "%USERPROFILE%\scoop\apps\7zip\current\7z.exe" set "SEVENZIP_DIR=%USERPROFILE%\scoop\apps\7zip\current"
if not defined SEVENZIP_DIR (echo 7-Zip が見つかりません。SEVENZIP_DIR を指定してください & exit /b 1)

uv sync || exit /b 1
uv run ruff check . || exit /b 1
uv run ruff format --check . || exit /b 1
uv run pytest -q || exit /b 1
uv run pyinstaller --noconfirm --clean --onefile --noconsole --name PSTMigration --paths src ^
  --hidden-import win32timezone ^
  --add-binary "%SEVENZIP_DIR%\7z.exe;." --add-binary "%SEVENZIP_DIR%\7z.dll;." ^
  --distpath build\dist --workpath build\work --specpath build ^
  src\pstmigration\__main__.py || exit /b 1

set "OUT=dist\PSTMigration"
if exist "%OUT%" rmdir /s /q "%OUT%"
mkdir "%OUT%"
copy /y build\dist\PSTMigration.exe "%OUT%\" >nul
copy /y config.json "%OUT%\" >nul
copy /y scripts\PSTMigration-manual.cmd "%OUT%\" >nul
copy /y scripts\PSTMigration-manual.ps1 "%OUT%\" >nul
if exist "%SEVENZIP_DIR%\License.txt" copy /y "%SEVENZIP_DIR%\License.txt" "%OUT%\7-Zip-License.txt" >nul

pushd "%OUT%"
(for %%f in (PSTMigration.exe config.json) do @for /f "skip=1 delims=" %%h in ('certutil -hashfile %%f SHA256 ^| findstr /v ":"') do @echo %%h  %%f) > SHA256SUMS.txt
popd

echo OK: %OUT%
type "%OUT%\SHA256SUMS.txt"
