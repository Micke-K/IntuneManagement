@echo off
rem PowerShell 7 (pwsh). For Windows PowerShell 5.1, use Start.cmd.
cmd /c pwsh -ex bypass -File "%~DP0Start.ps1" -showui
pause
