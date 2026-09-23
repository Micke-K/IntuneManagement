@echo off
rem Windows PowerShell 5.1. For PowerShell 7, use Start-PS7.cmd.
cmd /c powershell -version 5 -ex bypass -File "%~DP0Start.ps1" -showui
pause
