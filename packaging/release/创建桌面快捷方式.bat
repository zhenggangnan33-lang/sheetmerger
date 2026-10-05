@echo off
REM Create a desktop shortcut "SheetMerger <Chinese name>" for SheetMerger.exe in this folder.
REM The shortcut is saved under an ASCII name first, then renamed: WScript.Shell cannot save
REM non-ASCII file names on systems whose ANSI code page is not Chinese.
REM Keep this folder in a fixed place (e.g. D:\SheetMerger) before running; moving it breaks the shortcut.
setlocal
set "EXE=%~dp0SheetMerger.exe"
if not exist "%EXE%" (
    echo [ERROR] SheetMerger.exe not found next to this script.
    echo         Please extract the whole zip folder first, then run this script inside it.
    pause
    exit /b 1
)
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ErrorActionPreference = 'Stop';" ^
  "$name = 'SheetMerger ' + [string][char]0x591A + [char]0x8868 + [char]0x6C47 + [char]0x603B;" ^
  "$desk = [Environment]::GetFolderPath('Desktop');" ^
  "$lnk = Join-Path $desk ($name + '.lnk');" ^
  "$tmp = Join-Path $desk ('SheetMerger_tmp_' + $PID + '.lnk');" ^
  "$s = (New-Object -ComObject WScript.Shell).CreateShortcut($tmp);" ^
  "$s.TargetPath = $env:EXE; $s.WorkingDirectory = (Split-Path $env:EXE);" ^
  "$s.IconLocation = $env:EXE + ',0'; $s.Description = 'SheetMerger'; $s.Save();" ^
  "if (Test-Path -LiteralPath $lnk) { Remove-Item -LiteralPath $lnk -Force };" ^
  "Move-Item -LiteralPath $tmp -Destination $lnk;" ^
  "Write-Host ('OK: ' + $lnk)"
if errorlevel 1 (
    echo [ERROR] Failed to create the shortcut.
    pause
    exit /b 1
)
if not "%~1"=="/quiet" pause
endlocal
