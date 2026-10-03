@echo off
REM ============================================================
REM  SheetMerger - Windows build script
REM  Usage: double-click, or run "packaging\build_windows.bat" from cmd.
REM  Optional: "packaging\build_windows.bat cli" also builds SheetMerger-cli.exe
REM  Requires: Python 3.11 (64-bit) from python.org, "py" launcher available.
REM  Output:   dist\SheetMerger_v<version>\  (exe + docs + licenses)
REM ============================================================
setlocal
cd /d "%~dp0.."

where py >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python launcher "py" not found. Install Python 3.11 64-bit from python.org first.
    exit /b 1
)
py -3.11 -c "import sys; assert sys.maxsize > 2**32" >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python 3.11 64-bit is required.
    exit /b 1
)

echo [1/5] Creating build environment .venv-build ...
if not exist .venv-build py -3.11 -m venv .venv-build || exit /b 1
call .venv-build\Scripts\activate.bat || exit /b 1
python -m pip install --upgrade pip >nul
python -m pip install -r requirements.txt "pyinstaller>=6.0" || exit /b 1

echo [2/5] Running tests ...
python tests\generate_test_data.py >nul || exit /b 1
python -m pytest -q
if errorlevel 1 (
    echo [ERROR] Tests failed. Build stopped.
    exit /b 1
)

echo [3/5] Building exe with PyInstaller ...
if /i "%~1"=="cli" (set SHEETMERGER_BUILD_CLI=1) else (set SHEETMERGER_BUILD_CLI=)
pyinstaller --noconfirm --clean packaging\sheetmerger.spec || exit /b 1

echo [4/5] Smoke test ...
for /f "delims=" %%v in ('python -c "from core import __version__; print(__version__)"') do set VER=%%v
if exist dist\SheetMerger-cli.exe (
    dist\SheetMerger-cli.exe --input tests\test_data --output build\smoke.xlsx || exit /b 1
)

echo [5/5] Assembling release folder ...
set OUT=dist\SheetMerger_v%VER%
if exist "%OUT%" rmdir /s /q "%OUT%"
mkdir "%OUT%\licenses"
copy /y dist\SheetMerger.exe "%OUT%\" >nul
if exist dist\SheetMerger-cli.exe copy /y dist\SheetMerger-cli.exe "%OUT%\" >nul
copy /y README.md "%OUT%\" >nul
copy /y LICENSES.md "%OUT%\" >nul
copy /y packaging\licenses\*.txt "%OUT%\licenses\" >nul

echo.
echo Done: %OUT%
dir /b "%OUT%"
endlocal
