@echo off
setlocal ENABLEDELAYEDEXPANSION

set BASE=%~1
if "%BASE%"=="" set BASE=大版本更新
set PREFIX=%~2
if "%PREFIX%"=="" set PREFIX=daily

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0new-daily-branch.ps1" -Base "%BASE%" -Prefix "%PREFIX%" -AutoStash

endlocal
