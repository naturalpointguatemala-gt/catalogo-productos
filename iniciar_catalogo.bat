@echo off
title Catalogo Productos GitHub
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0vigilar_github.ps1"
pause
