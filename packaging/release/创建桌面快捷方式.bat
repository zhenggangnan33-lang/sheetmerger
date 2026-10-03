@echo off
REM Create a desktop shortcut "SheetMerger <Chinese name>" for SheetMerger.exe in this folder.
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
  "$name = 'SheetMerger ' + [string][char]0x591A + [char]0x8868 + [char]0x6C47 + [char]0x603B;" ^
  "$desk = [Environment]::GetFolderPath('Desktop');" ^
  "$lnk = Join-Path $desk ($name + '.lnk');" ^
  "$s = (New-Object -ComObject WScript.Shell).CreateShortcut($lnk);" ^
  "$s.TargetPath = $env:EXE; $s.WorkingDirectory = (Split-Path $env:EXE);" ^
  "$s.IconLocation = $env:EXE + ',0'; $s.Description = 'SheetMerger'; $s.Save();" ^
  "Write-Host ('OK: ' + $lnk)"
if errorlevel 1 (
    echo [ERROR] Failed to create the shortcut.
    pause
    exit /b 1
)
if not "%~1"=="/quiet" pause
endlocal
